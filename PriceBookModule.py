import sys
import os
import json
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QTableWidget, QTableWidgetItem, QLineEdit,
                             QPushButton, QMessageBox, QAbstractItemView,
                             QHeaderView, QApplication, QFormLayout)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QDoubleValidator

# Attempt to import SharePointExcelManager, handle gracefully if missing
try:
    from SharePointManager import SharePointExcelManager
except ImportError:
    SharePointExcelManager = None
    print("WARNING: SharePointManager could not be imported in PriceBookModule. Cannot load data.")

class PriceBookModule(QWidget):
    """
    Widget to display and search price book data from SharePoint Excel sheet,
    including calculated pricing based on exchange, markup, and margin.
    Markup/Margin update dynamically. Exchange rate, Markup, Margin persist.
    Prices formatted as currency.
    """
    reload_deal_requested = pyqtSignal(dict)

    def __init__(self, main_window=None, sharepoint_manager=None, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.sharepoint_manager = sharepoint_manager
        self.setObjectName("PriceBookModule")
        
        # Check if main_window has a data_path and use it directly
        if main_window and hasattr(main_window, 'data_path') and main_window.data_path:
            self._data_path = main_window.data_path
            print(f"DEBUG: Using data path from main_window in PriceBookModule: {self._data_path}")
        else:
            self._data_path = self._get_data_path() # Fallback to search method
            
        self.settings_file = os.path.join(self._data_path, "pricebook_settings.json") if self._data_path else None

        # Data storage
        self.price_headers = []
        self.price_data_rows = []
        self.code_col_idx = -1
        self.name_col_idx = -1
        self.cost_col_idx = -1
        # Default values before loading settings
        self.last_exchange_rate = 1.35
        self.last_markup = 20.0
        self.last_margin = 15.0 # This will be recalculated based on markup after load/init

        # Load settings BEFORE setting up UI that uses them
        self._load_settings()

        # Check if SharePoint connection is available
        if SharePointExcelManager is None:
            print("ERROR: SharePointManager class not available. Price Book disabled.")
            self._setup_error_ui("SharePoint Manager module not found.")
            return
        if self.sharepoint_manager is None:
             print("WARNING: SharePointManager instance not provided. Price Book disabled.")
             self._setup_error_ui("SharePoint connection not initialized.")
             return

        self.setup_ui()
        self.load_price_data() # Load data on initialization

    def _get_data_path(self):
        """Helper function to determine the path to the 'data' directory."""
        # Try BRIDeal data directory first (the correct location per your logs)
        current_dir = os.path.dirname(__file__)
        brideal_data_path = os.path.join(os.path.dirname(current_dir), 'data')
        if os.path.isdir(brideal_data_path):
            print(f"DEBUG: Using BRIDeal data path: {brideal_data_path}")
            return brideal_data_path
            
        # Fall back to other locations if BRIDeal data path doesn't exist
        project_root_guess = os.path.abspath(os.path.join(current_dir, os.pardir))
        data_path_parent = os.path.join(project_root_guess, 'data')
        data_path_relative = os.path.join(current_dir, 'data')
        data_path_user = os.path.join(os.path.expanduser("~"), 'data')
        data_path_assets = os.path.join(project_root_guess, 'assets')
        
        for path, name in [
            (data_path_parent, "parent data"),
            (data_path_relative, "relative data"),
            (data_path_user, "user home data"),
            (data_path_assets, "assets")
        ]:
            if os.path.isdir(path):
                print(f"DEBUG PriceBookModule: Using {name} path: {path}")
                return path
                
        print(f"CRITICAL WARNING: Could not locate 'data' or 'assets' directory near {current_dir}.")
        return None

    def _show_status_message(self, message, timeout=3000):
        """Helper to show messages on main window status bar or print."""
        if self.main_window and hasattr(self.main_window, 'statusBar'):
            self.main_window.statusBar().showMessage(message, timeout)
        else:
            print(f"Status: {message}")

    def _setup_error_ui(self, error_message):
        """Setup UI to show an error if initialization fails."""
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        error_label = QLabel(f"❌ Error: Price Book Unavailable\n\n{error_message}")
        error_label.setAlignment(Qt.AlignCenter)
        error_label.setStyleSheet("font-size: 16px; color: red;")
        layout.addWidget(error_label)

    def setup_ui(self):
        """Set up the UI elements for the price book module."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # --- Title ---
        title = QLabel("📖 Price Book Search")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #2a5d24; margin-bottom: 10px;")
        layout.addWidget(title)

        # --- Calculation Inputs ---
        calc_layout = QHBoxLayout()
        exchange_validator = QDoubleValidator(0.0, 10.0, 4)
        percent_validator = QDoubleValidator(0.0, 99.99, 2) # Margin < 100

        calc_layout.addWidget(QLabel("Exchange Rate:"))
        # Use loaded value
        self.exchange_rate_input = QLineEdit(f"{self.last_exchange_rate:.4f}")
        self.exchange_rate_input.setValidator(exchange_validator)
        self.exchange_rate_input.setFixedWidth(80)
        self.exchange_rate_input.editingFinished.connect(self._save_settings)
        calc_layout.addWidget(self.exchange_rate_input)
        calc_layout.addSpacing(20)

        calc_layout.addWidget(QLabel("Markup %:"))
        # Use loaded value
        self.markup_input = QLineEdit(f"{self.last_markup:.2f}")
        self.markup_input.setValidator(percent_validator)
        self.markup_input.setFixedWidth(60)
        self.markup_input.editingFinished.connect(self._update_margin_from_markup)
        self.markup_input.editingFinished.connect(self._save_settings) # Also save on change
        calc_layout.addWidget(self.markup_input)
        calc_layout.addSpacing(20)

        calc_layout.addWidget(QLabel("Margin %:"))
        # Use loaded value
        self.margin_input = QLineEdit(f"{self.last_margin:.2f}")
        self.margin_input.setValidator(percent_validator)
        self.margin_input.setFixedWidth(60)
        self.margin_input.editingFinished.connect(self._update_markup_from_margin)
        self.margin_input.editingFinished.connect(self._save_settings) # Also save on change
        calc_layout.addWidget(self.margin_input)
        calc_layout.addStretch()

        layout.addLayout(calc_layout)
        # Initial calculation based on loaded/default values
        self._update_markup_from_margin() # Update markup based on initial margin


        # --- Search Bar ---
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Search by Code/Name:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Enter search term and press Enter or click Search")
        self.search_input.returnPressed.connect(self._execute_search)
        search_layout.addWidget(self.search_input)

        self.search_btn = QPushButton("🔍 Search")
        self.search_btn.setToolTip("Search for matching items and calculate prices")
        self.search_btn.clicked.connect(self._execute_search)
        search_layout.addWidget(self.search_btn)

        self.refresh_btn = QPushButton("🔄 Refresh Data")
        self.refresh_btn.setToolTip("Reload all data from SharePoint (clears search results)")
        self.refresh_btn.clicked.connect(self.load_price_data)
        search_layout.addWidget(self.refresh_btn)

        layout.addLayout(search_layout)

        # --- Table Widget ---
        self.table = QTableWidget()
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers) # Read-only
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.verticalHeader().setVisible(False) # Hide row numbers
        self.table.setStyleSheet("font-size: 10pt;")
        # Updated headers including calculated fields
        self.display_headers = ["ProductCode", "Product Name", "USD Cost", "CAD Cost", "Markup Price", "Margin Price"]
        self.table.setColumnCount(len(self.display_headers))
        self.table.setHorizontalHeaderLabels(self.display_headers)
        self.table.setRowCount(0) # Start empty
        # Set resize modes
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch) # Stretch Product Name
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)

        layout.addWidget(self.table)


    def load_price_data(self):
        """Loads data from the 'App Source' sheet into memory."""
        if not self.sharepoint_manager:
            QMessageBox.critical(self, "Error", "SharePoint connection is not available.")
            self.price_data_rows = []; self.price_headers = []
            return

        self._show_status_message("Loading Price Book data from SharePoint...", 0)
        QApplication.processEvents()

        sheet_name = "App Source"
        sheet_data = self.sharepoint_manager.read_excel_sheet(sheet_name)

        self.table.setRowCount(0) # Clear table display on reload/error
        self.price_data_rows = [] # Clear internal data
        self.price_headers = []
        self.code_col_idx = self.name_col_idx = self.cost_col_idx = -1 # Reset indices

        if sheet_data is None:
            self._show_status_message(f"Error loading Price Book.", 5000)
            return
        if not sheet_data or len(sheet_data) < 1:
            self._show_status_message(f"Price Book sheet '{sheet_name}' is empty.", 3000)
            return

        # Store headers and data rows
        self.price_headers = [str(h).strip() for h in sheet_data[0]]
        self.price_data_rows = sheet_data[1:]

        # Find column indices
        try:
            # Adjust column names to match EXACTLY what's in the sheet header
            self.code_col_idx = self.price_headers.index("ProductCode")
            self.name_col_idx = self.price_headers.index("Product Name")
            self.cost_col_idx = self.price_headers.index("USD Cost")
            print(f"DEBUG: Found column indices - Code: {self.code_col_idx}, Name: {self.name_col_idx}, Cost: {self.cost_col_idx}")
        except ValueError as e:
            print(f"ERROR: Could not find required column in Price Book headers: {e}")
            QMessageBox.critical(self, "Header Error", f"Required column missing in 'App Source' sheet header: {e}\nExpected: ProductCode, Product Name, USD Cost")
            self.price_headers = []; self.price_data_rows = [] # Clear data
            self._show_status_message("Error: Price Book header mismatch.", 5000)
            return

        self._show_status_message(f"Price Book data loaded ({len(self.price_data_rows)} items). Enter search term.", 5000)
        print(f"DEBUG: Price Book data loaded ({len(self.price_data_rows)} items).")
        # Update table headers now that we know the columns to display
        self.table.setColumnCount(len(self.display_headers))
        self.table.setHorizontalHeaderLabels(self.display_headers)


    def _execute_search(self):
        """Filters the stored data based on search input, calculates prices, and displays results."""
        search_term = self.search_input.text().strip().lower()

        if not self.price_data_rows or self.code_col_idx == -1 or self.name_col_idx == -1 or self.cost_col_idx == -1:
            QMessageBox.warning(self, "Search Error", "Price book data is not loaded correctly or headers are missing.\nTry clicking 'Refresh Data'.")
            return

        # --- Get Calculation Inputs ---
        try: exchange_rate = float(self.exchange_rate_input.text() or 1.0)
        except ValueError: QMessageBox.warning(self, "Input Error", "Invalid Exchange Rate. Using 1.0."); exchange_rate = 1.0
        try: markup_percent = float(self.markup_input.text() or 0.0)
        except ValueError: QMessageBox.warning(self, "Input Error", "Invalid Markup %. Using 0."); markup_percent = 0.0
        try:
            margin_percent = float(self.margin_input.text() or 0.0)
            if margin_percent >= 100.0: QMessageBox.warning(self, "Input Error", "Margin % cannot be >= 100%. Using 0."); margin_percent = 0.0
        except ValueError: QMessageBox.warning(self, "Input Error", "Invalid Margin %. Using 0."); margin_percent = 0.0

        # --- Filter and Calculate ---
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0) # Clear previous results
        self._show_status_message(f"Searching for '{search_term}'...", 0)
        QApplication.processEvents()

        results_found = 0
        for row_data in self.price_data_rows:
            # Ensure row has enough columns
            if len(row_data) <= max(self.code_col_idx, self.name_col_idx, self.cost_col_idx): continue

            code_val = str(row_data[self.code_col_idx]).lower()
            name_val = str(row_data[self.name_col_idx]).lower()

            # Check if search term is in Code OR Name
            if not search_term or (search_term in code_val or search_term in name_val):
                if not search_term and results_found >= 500: # Limit display if search is empty
                     if results_found == 500: print("WARN: Display limit reached for empty search.")
                     continue

                # --- Perform Calculations ---
                usd_cost_str = str(row_data[self.cost_col_idx]).replace('$', '').replace(',', '').strip()
                usd_cost_display = usd_cost_str # Default display if calc fails
                cad_cost_str = "Error"
                markup_price_str = "Error"
                margin_price_str = "Error"
                try:
                    usd_cost = float(usd_cost_str) if usd_cost_str else 0.0
                    usd_cost_display = f"${usd_cost:,.2f}" # Format USD cost

                    cad_cost = usd_cost * exchange_rate
                    cad_cost_str = f"${cad_cost:,.2f}" # Format CAD cost

                    markup_price = cad_cost * (1 + markup_percent / 100.0)
                    markup_price_str = f"${markup_price:,.2f}" # Format Markup Price

                    denominator = 1.0 - margin_percent / 100.0
                    if denominator <= 0: margin_price_str = "N/A"
                    else: margin_price = cad_cost / denominator; margin_price_str = f"${margin_price:,.2f}" # Format Margin Price

                except (ValueError, TypeError) as calc_err:
                    print(f"WARN: Could not calculate price for row {row_data}: {calc_err}")
                    usd_cost_display = usd_cost_str # Show original string on error

                # Add row to table
                current_row = self.table.rowCount()
                self.table.insertRow(current_row)

                code_display = str(row_data[self.code_col_idx]) if row_data[self.code_col_idx] is not None else ""
                name_display = str(row_data[self.name_col_idx]) if row_data[self.name_col_idx] is not None else ""

                # Create items
                item_code = QTableWidgetItem(code_display)
                item_name = QTableWidgetItem(name_display)
                item_usd = QTableWidgetItem(usd_cost_display)
                item_cad = QTableWidgetItem(cad_cost_str)
                item_markup = QTableWidgetItem(markup_price_str)
                item_margin = QTableWidgetItem(margin_price_str)

                # Set alignment for price columns
                item_usd.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                item_cad.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                item_markup.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                item_margin.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

                # Populate the specific columns
                self.table.setItem(current_row, 0, item_code)      # ProductCode
                self.table.setItem(current_row, 1, item_name)      # Product Name
                self.table.setItem(current_row, 2, item_usd)       # USD Cost
                self.table.setItem(current_row, 3, item_cad)       # CAD Cost
                self.table.setItem(current_row, 4, item_markup)    # Markup Price
                self.table.setItem(current_row, 5, item_margin)    # Margin Price
                results_found += 1

        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch) # Stretch Product Name again
        self.table.setSortingEnabled(True)
        self._show_status_message(f"Found {results_found} matching item(s).", 3000)
        print(f"DEBUG: Search complete. Found {results_found} items.")

    # --- Dynamic Update Methods ---
    def _update_margin_from_markup(self):
        """Calculates Margin % when Markup % editing is finished."""
        try:
            markup_percent = float(self.markup_input.text() or 0.0)
            if 1.0 + markup_percent / 100.0 <= 0: margin_percent_str = ""
            else: margin_percent = (1.0 - (1.0 / (1.0 + markup_percent / 100.0))) * 100.0; margin_percent_str = f"{margin_percent:.2f}"
            self.margin_input.blockSignals(True); self.margin_input.setText(margin_percent_str); self.margin_input.blockSignals(False)
        except ValueError: self.margin_input.blockSignals(True); self.margin_input.clear(); self.margin_input.blockSignals(False)
        except Exception as e: print(f"Error calculating margin from markup: {e}")

    def _update_markup_from_margin(self):
        """Calculates Markup % when Margin % editing is finished."""
        try:
            margin_percent = float(self.margin_input.text() or 0.0)
            denominator = 1.0 - margin_percent / 100.0
            if denominator <= 0: markup_percent_str = ""
            else: markup_percent = ((1.0 / denominator) - 1.0) * 100.0; markup_percent_str = f"{markup_percent:.2f}"
            self.markup_input.blockSignals(True); self.markup_input.setText(markup_percent_str); self.markup_input.blockSignals(False)
        except ValueError: self.markup_input.blockSignals(True); self.markup_input.clear(); self.markup_input.blockSignals(False)
        except Exception as e: print(f"Error calculating markup from margin: {e}")

    # --- Settings Load/Save ---
    def _load_settings(self):
        """Loads last used values from JSON file."""
        defaults = {"last_exchange_rate": 1.35, "last_markup": 20.0, "last_margin": 15.0} # Define defaults
        if not self.settings_file:
            print("WARN: Settings file path not set, using default values.")
            self.last_exchange_rate = defaults["last_exchange_rate"]
            self.last_markup = defaults["last_markup"]
            self.last_margin = defaults["last_margin"]
            return

        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    # Load each value, falling back to default if key missing or type error
                    try: self.last_exchange_rate = float(settings.get("last_exchange_rate", defaults["last_exchange_rate"]))
                    except (ValueError, TypeError): self.last_exchange_rate = defaults["last_exchange_rate"]
                    try: self.last_markup = float(settings.get("last_markup_percent", defaults["last_markup"])) # Use key from save
                    except (ValueError, TypeError): self.last_markup = defaults["last_markup"]
                    try: self.last_margin = float(settings.get("last_margin_percent", defaults["last_margin"])) # Use key from save
                    except (ValueError, TypeError): self.last_margin = defaults["last_margin"]

                    print(f"DEBUG: Loaded settings: Exch={self.last_exchange_rate}, Markup={self.last_markup}, Margin={self.last_margin}")
            else:
                print("DEBUG: Settings file not found, using default values.")
                self.last_exchange_rate = defaults["last_exchange_rate"]
                self.last_markup = defaults["last_markup"]
                self.last_margin = defaults["last_margin"]
        except (IOError, json.JSONDecodeError) as e:
            print(f"WARN: Could not load or parse settings file '{self.settings_file}': {e}. Using defaults.")
            self.last_exchange_rate = defaults["last_exchange_rate"]
            self.last_markup = defaults["last_markup"]
            self.last_margin = defaults["last_margin"]
        # Recalculate margin based on loaded markup initially for consistency
        # self._update_margin_from_markup() # Or recalculate markup from margin

    def _save_settings(self):
        """Saves the current exchange rate, markup, and margin to JSON file."""
        if not self.settings_file:
            print("WARN: Settings file path not set, cannot save settings.")
            return

        settings_to_save = {}
        # Validate and store current values
        try: settings_to_save["last_exchange_rate"] = float(self.exchange_rate_input.text())
        except ValueError: settings_to_save["last_exchange_rate"] = self.last_exchange_rate # Save last known good value
        try: settings_to_save["last_markup_percent"] = float(self.markup_input.text())
        except ValueError: settings_to_save["last_markup_percent"] = self.last_markup
        try: settings_to_save["last_margin_percent"] = float(self.margin_input.text())
        except ValueError: settings_to_save["last_margin_percent"] = self.last_margin

        # Update internal 'last' values as well
        self.last_exchange_rate = settings_to_save["last_exchange_rate"]
        self.last_markup = settings_to_save["last_markup_percent"]
        self.last_margin = settings_to_save["last_margin_percent"]

        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings_to_save, f, indent=4)
            print(f"DEBUG: Saved settings: {settings_to_save}")
        except IOError as e:
            print(f"ERROR: Could not write settings file '{self.settings_file}': {e}")
            QMessageBox.warning(self, "Save Setting Error", f"Could not save settings:\n{e}")
        except Exception as e:
            print(f"ERROR: Unexpected error saving settings: {e}")


    # --- Placeholder methods for signals/buttons from RecentDeals template ---
    def _get_selected_deal_data(self): return None
    def _reload_deal(self): pass
    def _regenerate_email(self): pass
    def _regenerate_csv(self): pass


# --- Example Usage (for testing PriceBookModule independently) ---
if __name__ == '__main__':
    app = QApplication(sys.argv)

    # --- Mock SharePoint Manager for testing ---
    class MockSPManager:
         def read_excel_sheet(self, sheet_name):
             print(f"MockSPManager: Reading sheet '{sheet_name}'")
             if sheet_name == "App Source":
                 return [
                     ["ProductCode", "Product Name", "USD Cost", "Category", "Notes"], # Header row
                     ["PF1111", "Product Alpha", "19.99", "Widgets", "Note A"],
                     ["PF2222", "Product Beta", "120.50", "Gadgets", "Note B"],
                     ["PF3333", "Thingamajig", "75.00", "Widgets", "Note C"],
                     ["PF4444", "Alpha Widget", "25.95", "Widgets", "Note D"]
                 ]
             else:
                 print(f"MockSPManager: Sheet '{sheet_name}' not found.")
                 return None

    # --- Dummy MainWindow for status bar ---
    class MockMainWindow:
        def statusBar(self): return self
        def showMessage(self, msg, timeout): print(f"Mock Status: {msg} ({timeout}ms)")

    mock_main = MockMainWindow()
    mock_sp = MockSPManager()
    # --- End Mocks ---

    # Create and show the module
    test_data_path = "."
    try: os.makedirs(test_data_path, exist_ok=True)
    except: pass

    price_book = PriceBookModule(main_window=mock_main, sharepoint_manager=mock_sp, data_path=test_data_path) # Pass data_path
    price_book.setWindowTitle("Price Book Module (Standalone Test)")
    price_book.resize(800, 600) # Set a default size
    price_book.show()

    sys.exit(app.exec_())