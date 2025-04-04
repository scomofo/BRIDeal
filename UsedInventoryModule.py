import sys
import os
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QTableWidget, QTableWidgetItem, QLineEdit,
                             QPushButton, QMessageBox, QAbstractItemView,
                             QHeaderView, QApplication)
from PyQt5.QtCore import Qt

# Attempt to import SharePointExcelManager, handle gracefully if missing
try:
    # Assumes SharePointManager.py is in the same directory or accessible via PYTHONPATH
    from SharePointManager import SharePointExcelManager
except ImportError:
    SharePointExcelManager = None
    print("WARNING: SharePointManager could not be imported in UsedInventoryModule. Cannot load data.")

class UsedInventoryModule(QWidget):
    """
    Widget to display and search Used AMS inventory data from SharePoint Excel sheet.
    """
    def __init__(self, main_window=None, sharepoint_manager=None, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.sharepoint_manager = sharepoint_manager
        self.setObjectName("UsedInventoryModule")

        # Data storage
        self.inventory_headers = []
        self.inventory_data_rows = []

        # Check if SharePoint connection is available
        if SharePointExcelManager is None:
            print("ERROR: SharePointManager class not available. Used Inventory disabled.")
            self._setup_error_ui("SharePoint Manager module not found.")
            return
        if self.sharepoint_manager is None:
             print("WARNING: SharePointManager instance not provided. Used Inventory disabled.")
             self._setup_error_ui("SharePoint connection not initialized.")
             return

        self.setup_ui()
        self.load_inventory_data() # Load data on initialization

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
        error_label = QLabel(f"❌ Error: Used Inventory Unavailable\n\n{error_message}")
        error_label.setAlignment(Qt.AlignCenter)
        error_label.setStyleSheet("font-size: 16px; color: red;")
        layout.addWidget(error_label)

    def setup_ui(self):
        """Set up the UI elements for the used inventory module."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # --- Title ---
        title = QLabel("🚜 Used Inventory (from 'Used AMS' Sheet)")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #2a5d24; margin-bottom: 10px;")
        layout.addWidget(title)

        # --- Search Bar ---
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Search:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Filter table by any column...")
        self.search_input.textChanged.connect(self._filter_table) # Filter as user types
        search_layout.addWidget(self.search_input)

        self.refresh_btn = QPushButton("🔄 Refresh Data")
        self.refresh_btn.setToolTip("Reload data from SharePoint")
        self.refresh_btn.clicked.connect(self.load_inventory_data)
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
        self.table.setSortingEnabled(True) # Enable sorting
        layout.addWidget(self.table)

        # Set initial message
        self.table.setRowCount(1)
        self.table.setColumnCount(1)
        self.table.setHorizontalHeaderLabels(["Status"])
        self.table.setItem(0, 0, QTableWidgetItem("Loading data..."))
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)


    def load_inventory_data(self):
        """Loads data from the 'Used AMS' sheet via SharePointManager."""
        if not self.sharepoint_manager:
            QMessageBox.critical(self, "Error", "SharePoint connection is not available.")
            return

        self._show_status_message("Loading Used Inventory data from SharePoint...", 0) # Persistent message
        QApplication.processEvents() # Update UI

        sheet_name = "Used AMS" # Sheet name specified by user
        sheet_data = self.sharepoint_manager.read_excel_sheet(sheet_name)

        self.table.setSortingEnabled(False) # Disable sorting during population
        self.table.clearContents() # Clear existing data but keep headers

        if sheet_data is None:
            self.table.setRowCount(1); self.table.setColumnCount(1)
            self.table.setHorizontalHeaderLabels(["Error"])
            self.table.setItem(0, 0, QTableWidgetItem(f"Failed to load data from sheet '{sheet_name}'."))
            self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
            self._show_status_message(f"Error loading Used Inventory.", 5000)
            return

        if not sheet_data or len(sheet_data) < 1: # Check if sheet_data is empty or just header
            self.table.setRowCount(1); self.table.setColumnCount(1)
            self.table.setHorizontalHeaderLabels(["Info"])
            self.table.setItem(0, 0, QTableWidgetItem(f"No data found in sheet '{sheet_name}'."))
            self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
            self._show_status_message(f"Used Inventory sheet '{sheet_name}' is empty.", 3000)
            return

        # Assume first row is header
        self.inventory_headers = [str(h).strip() for h in sheet_data[0]]
        self.inventory_data_rows = sheet_data[1:] # Store data rows only

        self.table.setColumnCount(len(self.inventory_headers))
        self.table.setHorizontalHeaderLabels(self.inventory_headers)
        self.table.setRowCount(len(self.inventory_data_rows))

        for r, row_data in enumerate(self.inventory_data_rows):
            for c, cell_value in enumerate(row_data):
                 if c < len(self.inventory_headers): # Avoid writing past header count
                     item_text = str(cell_value) if cell_value is not None else ""
                     self.table.setItem(r, c, QTableWidgetItem(item_text))

        self.table.resizeColumnsToContents()
        # Optionally stretch a specific column like Description if it exists
        try:
             desc_col_index = self.inventory_headers.index("Description") # Example
             self.table.horizontalHeader().setSectionResizeMode(desc_col_index, QHeaderView.Stretch)
        except ValueError:
             self.table.horizontalHeader().setStretchLastSection(True) # Fallback

        self.table.setSortingEnabled(True) # Re-enable sorting
        self._show_status_message(f"Used Inventory loaded ({len(self.inventory_data_rows)} items).", 3000)
        print(f"DEBUG: Used Inventory loaded {len(self.inventory_data_rows)} items.")

    def _filter_table(self):
        """Hides rows that do not contain the search text in any column."""
        search_text = self.search_input.text().strip().lower()
        if not hasattr(self, 'table'): # Check if table exists (might not if init failed)
             return

        for row in range(self.table.rowCount()):
            row_matches = False
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item and search_text in item.text().lower():
                    row_matches = True
                    break # Found a match in this row, no need to check other columns
            self.table.setRowHidden(row, not row_matches) # Hide row if it doesn't match


# --- Example Usage (for testing UsedInventoryModule independently) ---
if __name__ == '__main__':
    app = QApplication(sys.argv)

    # --- Mock SharePoint Manager for testing ---
    class MockSPManager:
         def read_excel_sheet(self, sheet_name):
             print(f"MockSPManager: Reading sheet '{sheet_name}'")
             if sheet_name == "Used AMS":
                 # Return dummy data matching typical structure
                 return [
                     ["Stock#", "Year", "Make", "Model", "Serial", "Location", "List Price"], # Header row
                     ["U1234", "2018", "John Deere", "S670", "SN123", "Camrose", "250000"],
                     ["U5678", "2020", "John Deere", "8R 340", "SN456", "Killam", "450000"],
                     ["U9012", "2019", "Case IH", "9250", "SN789", "Wainwright", "380000"],
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
    inventory_module = UsedInventoryModule(main_window=mock_main, sharepoint_manager=mock_sp, data_path=".")
    inventory_module.setWindowTitle("Used Inventory Module (Standalone Test)")
    inventory_module.resize(1000, 600) # Set a default size
    inventory_module.show()

    sys.exit(app.exec_())

