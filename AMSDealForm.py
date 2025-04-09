# --- Start of PyQt5 Imports ---
# QtWidgets: Common UI elements and layouts
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QMainWindow, QDockWidget, # Added QMainWindow, QDockWidget for test
                             QGridLayout, QGroupBox, QLabel, QLineEdit,
                             QPushButton, QComboBox, QCheckBox, QTextEdit,
                             QMessageBox, QApplication, QCompleter, QListWidget,
                             QListWidgetItem, QSpinBox, QInputDialog, QScrollArea,
                             QSizePolicy) # Added QSizePolicy

# QtGui: For handling images, icons, fonts, colors etc.
from PyQt5 import QtGui
from PyQt5.QtGui import QDoubleValidator, QPixmap, QClipboard # Added QPixmap, QClipboard based on usage
from PyQt5.QtCore import QMimeData # Added for clipboard

# QtCore: Core non-GUI functionality, including the Qt namespace
from PyQt5.QtCore import Qt, pyqtSignal, QDate, QSize # Added QDate, QSize based on usage
# --- End of PyQt5 Imports ---

# --- Keep your other imports ---
import sys
import os # Import os to use getenv
import io # Import io for csv reading from string
import json # Added for Save/Load Draft
# config Imports (Ensure these are correct and config.py defines/loads them from .env)
# *** Reverted to importing SITE_ID, removed FILE_PATH import ***
# Assuming config.py exists and defines these classes or variables
try:
    from config import AzureConfig, APIConfig, SharePointConfig
except ImportError:
    print("WARNING: Could not import from config.py. Using dummy values if needed.")
    # Define dummy classes/variables if necessary for the script to run without config.py
    class DummyConfig: pass
    AzureConfig = DummyConfig()
    APIConfig = DummyConfig()
    SharePointConfig = DummyConfig()


# Standard library imports used later
import csv
from datetime import datetime # Use datetime for timestamp
from urllib.parse import quote
import webbrowser # Re-added for mailto
import requests # Needed for Graph API calls
import re # Import regex for parsing edited items
import html # <-- Ensure this import is present
import traceback # For detailed error logging

# --- Import Authentication Function ---
# *** Updated to import the correct function name from auth.py ***
try:
    from auth import get_access_token # <-- CORRECTED FUNCTION NAME
except ImportError as e:
    # Handle case where auth.py exists but the function doesn't (less likely now)
    # or if auth.py itself cannot be imported.
    print(f"CRITICAL ERROR: Could not import 'get_access_token' from auth.py: {e}. SharePoint features will fail.")
    # Define a dummy function to prevent NameError later, but show a clear error
    def get_access_token():
        print("ERROR: auth.get_access_token function is missing or auth.py failed to import!")
        raise RuntimeError("Authentication function not found or import failed.")

# --- Import SharePoint Manager ---
# Assumes SharePointManager.py is in the same directory or accessible via PYTHONPATH
try:
    from SharePointManager import SharePointExcelManager
except ImportError:
    SharePointExcelManager = None # Set to None if import fails
    print("CRITICAL ERROR: SharePointManager.py could not be imported. SharePoint features will fail.")
    # Define a dummy class if needed to prevent NameErrors later, though checks should handle None
    class SharePointExcelManager:
        def __init__(self, *args, **kwargs): # Accept args/kwargs for flexibility
             print("ERROR: Using Dummy SharePointExcelManager because import failed.")
        def update_excel_data(self, data): print("Dummy SP Manager: Update called"); return False
        # Keep dummy send method even if unused by DealForm now
        def send_html_email(self, r, s, b): print("Dummy SP Manager: Send Email called"); return False


# --- AMSDealForm Class ---
class AMSDealForm(QWidget):
    # Define signals if needed for communication (e.g., with RecentDealsModule)
    # deal_processed = pyqtSignal(dict) # Example signal

    def __init__(self, main_window=None, sharepoint_manager=None, parent=None): # Added sharepoint_manager parameter
        super().__init__(parent)
        print("DEBUG: Initializing AMSDealForm...") # Add init log
        self.main_window = main_window if main_window else self._create_mock_main_window()
        self.sharepoint_manager = sharepoint_manager # Store passed-in manager
        self.csv_lines = []
        self._data_path = self._get_data_path() # Store data path
        print(f"DEBUG: Final data path determined: {self._data_path}") # Log the determined path

        # Define paths for draft and recent deals files
        self.draft_file_path = os.path.join(self._data_path, "deal_draft.json") if self._data_path else None
        self.recent_deals_file = os.path.join(self._data_path, "recent_deals.json") if self._data_path else None
        if not self.draft_file_path: print("WARNING: Could not determine path for draft file.")
        if not self.recent_deals_file: print("WARNING: Could not determine path for recent deals file.")

        # Check if the passed manager is valid
        if SharePointExcelManager is None:
             print("ERROR: SharePointExcelManager class not imported. Disabling SharePoint features.")
             self.sharepoint_manager = None # Ensure it's None
        elif self.sharepoint_manager is None:
             # This case happens if main.py failed to create the manager instance
             print("WARNING: SharePoint Manager instance was not provided. SharePoint features disabled.")

        # Load initial data
        print("DEBUG: Loading data for AMSDealForm...")
        self.products_dict = self.load_products()
        self.parts_dict = self.load_parts()
        self.customers_dict = self.load_customers()
        self.salesmen_emails = self.load_salesmen_emails()
        print("DEBUG: Data loading checks complete for AMSDealForm.") # Renamed log message

        # Check if essential data loaded - use empty dicts as fallback
        # No critical errors here, just warnings if load failed
        if not self.products_dict: # Check if it's None OR empty
             print("WARNING: products_dict is empty after loading. Product completers will be empty.")
             if self.products_dict is None: self.products_dict = {} # Ensure it's a dict
        if not self.parts_dict:
             print("WARNING: parts_dict is empty after loading. Parts completers will be empty.")
             if self.parts_dict is None: self.parts_dict = {}
        if not self.customers_dict:
             print("WARNING: customers_dict is empty after loading. Customer completer will be empty.")
             if self.customers_dict is None: self.customers_dict = {}
        if not self.salesmen_emails:
             print("WARNING: salesmen_emails is empty after loading. Salesmen completer will be empty.")
             if self.salesmen_emails is None: self.salesmen_emails = {}

        self.setup_ui()
        self.apply_styles()
        self.last_charge_to = ""
        print("DEBUG: AMSDealForm initialization complete.")


    # --- Helper to show status messages ---
    def _show_status_message(self, message, timeout=3000):
        """Helper to show messages on main window status bar or print."""
        if self.main_window and hasattr(self.main_window, 'statusBar'):
            try:
                # Ensure statusBar() returns a valid object with showMessage
                status_bar = self.main_window.statusBar()
                if hasattr(status_bar, 'showMessage'):
                    status_bar.showMessage(message, timeout)
                else:
                     print(f"Status (Main window has status bar, but no showMessage method): {message}")
            except Exception as e:
                 print(f"Error showing status message on main window: {e}")
                 print(f"Status (Fallback): {message}")
        else:
            print(f"Status: {message}")


    def _create_mock_main_window(self):
        """Creates a mock main window if none is provided, for status bar messages."""
        class MockStatusBar:
            def showMessage(self, message, timeout=0):
                print(f"Status Bar (Mock): {message} (timeout: {timeout})")
        class MockMainWindow:
            def statusBar(self):
                return MockStatusBar()
            # Add a dummy data_path attribute for testing _get_data_path
            data_path = None
        print("Warning: No main_window provided to AMSDealForm. Using mock for status bar.")
        return MockMainWindow()

    def _get_data_path(self):
        """Helper function to determine the path to the 'data' directory."""
        print("DEBUG: Determining data path...")
        # First, check if main_window has a data_path attribute
        if self.main_window and hasattr(self.main_window, 'data_path') and self.main_window.data_path:
            main_window_path = self.main_window.data_path
            if os.path.isdir(main_window_path):
                print(f"DEBUG: Using data path from main_window: {main_window_path}")
                return main_window_path
            else:
                print(f"DEBUG: main_window.data_path provided but not a valid directory: {main_window_path}")

        # If main_window path doesn't exist or wasn't provided, try standard locations
        try:
            # Get the directory where the current script (AMSDealForm.py) is located
            current_dir = os.path.dirname(os.path.abspath(__file__))
            print(f"DEBUG: Script directory: {current_dir}")
        except NameError:
             # __file__ might not be defined if running in certain environments (e.g., interactive)
             current_dir = os.getcwd()
             print(f"DEBUG: Could not use __file__, using current working directory: {current_dir}")

        # Try relative 'data' directory next to the script first
        brideal_data_path = os.path.join(current_dir, 'data')
        print(f"DEBUG: Checking relative path: {brideal_data_path}")
        if os.path.isdir(brideal_data_path):
            print(f"DEBUG: Using relative data path: {brideal_data_path}")
            return brideal_data_path

        # Fall back to other possible locations
        data_path_parent = os.path.join(os.path.abspath(os.path.join(current_dir, os.pardir)), 'data')
        data_path_user = os.path.join(os.path.expanduser("~"), 'data')
        data_path_assets = os.path.join(os.path.abspath(os.path.join(current_dir, os.pardir)), 'assets')
        assets_path_relative = os.path.join(current_dir, 'assets')

        for path, name in [
            (data_path_parent, "parent 'data'"),
            (data_path_user, "user home 'data'"),
            (data_path_assets, "parent 'assets'"),
            (assets_path_relative, "relative 'assets'")
        ]:
            print(f"DEBUG: Checking fallback path ({name}): {path}")
            if os.path.isdir(path):
                print(f"DEBUG: Using fallback {name} path: {path}")
                return path

        print(f"CRITICAL WARNING: Could not locate 'data' or 'assets' directory relative to {current_dir} or other fallback locations.")
        QMessageBox.critical(self, "Data Directory Error",
                             f"Could not find the 'data' directory.\n"
                             f"Checked relative path: {brideal_data_path}\n"
                             f"Checked user home: {data_path_user}\n"
                             f"And other fallbacks.\n\nPlease ensure the 'data' folder exists next to the application.")
        return None

    # --- Data Loading Methods (Using _data_path) ---
    def _load_csv_generic(self, filename, required_headers=None, key_column=None, value_column=None, is_dict=True):
        """
        Generic CSV loader with encoding fallback, header check, and dict/list output.
        More flexible header checking and enhanced debugging.
        """
        print(f"DEBUG Load ({filename}): --- ENTERING _load_csv_generic ---")

        data = {} if is_dict else []
        if not self._data_path:
            print(f"ERROR ({filename}): Cannot load - data path not determined.")
            print(f"DEBUG Load ({filename}): Returning None because _data_path is not set.")
            return None # Return None to indicate failure

        file_path = os.path.join(self._data_path, filename)
        print(f"DEBUG Load ({filename}): Attempting to load from: {file_path}")
        encodings = ['utf-8', 'latin1', 'windows-1252']
        loaded = False
        rows_processed = 0
        last_error = "File not found" # Default error if check fails

        if not os.path.exists(file_path):
            print(f"Warning ({filename}): File not found at expected path.")
            print(f"DEBUG Load ({filename}): Returning None because file does not exist at: {file_path}")
            return None # Return None

        for encoding in encodings:
            print(f"DEBUG Load ({filename}): Trying encoding '{encoding}'...")
            try:
                with open(file_path, mode='r', newline='', encoding=encoding) as infile:
                    # --- IMPROVED DEBUGGING: Log first few lines ---
                    try:
                         preview = "".join(infile.readline() for _ in range(3))
                         print(f"DEBUG Load ({filename}): File preview (first ~3 lines with {encoding}):\n'''\n{preview}\n'''")
                         infile.seek(0) # IMPORTANT: Reset file pointer after preview
                    except Exception as preview_err:
                         print(f"DEBUG Load ({filename}): Could not read file preview: {preview_err}")
                         infile.seek(0) # Ensure reset even if preview failed
                    # --- End Improved Debugging ---

                    # Use DictReader if we need a dictionary or have headers
                    if is_dict or required_headers:
                        reader = csv.DictReader(infile)
                        actual_fieldnames = reader.fieldnames or []
                        print(f"DEBUG Load ({filename}): Actual headers found with {encoding}: {actual_fieldnames}")

                        if not actual_fieldnames:
                             print(f"Warning ({filename}): No headers found with {encoding}. Trying next encoding.")
                             last_error = f"No headers found with {encoding}"
                             continue # Try next encoding

                        actual_headers_lower_stripped = {h.lower().strip() for h in actual_fieldnames}

                        # --- MORE FLEXIBLE HEADER CHECK ---
                        key_header_actual = None
                        value_header_actual = None

                        if key_column:
                            key_header_actual = next((h for h in actual_fieldnames if h.lower().strip() == key_column.lower()), None)
                            if not key_header_actual:
                                print(f"Warning ({filename}): Required key column '{key_column}' (case-insensitive search for '{key_column.lower()}') not found in actual headers {actual_fieldnames} with {encoding}. Trying next encoding.")
                                last_error = f"Key column '{key_column}' not found with {encoding}"
                                continue # Try next encoding
                            else:
                                print(f"DEBUG Load ({filename}): Found actual key header '{key_header_actual}' for requested key column '{key_column}'.")


                        if value_column:
                            value_header_actual = next((h for h in actual_fieldnames if h.lower().strip() == value_column.lower()), None)
                            if not value_header_actual:
                                print(f"Warning ({filename}): Required value column '{value_column}' (case-insensitive search for '{value_column.lower()}') not found in actual headers {actual_fieldnames} with {encoding}. Trying next encoding.")
                                last_error = f"Value column '{value_column}' not found with {encoding}"
                                continue # Try next encoding
                            else:
                                print(f"DEBUG Load ({filename}): Found actual value header '{value_header_actual}' for requested value column '{value_column}'.")


                        if required_headers:
                             missing_required = set(h.lower() for h in required_headers) - actual_headers_lower_stripped
                             if missing_required:
                                  print(f"DEBUG Load ({filename}): Note - some optional 'required_headers' were not found: {missing_required}")
                        # --- END FLEXIBLE HEADER CHECK ---

                        temp_data = {}
                        current_row_num = 1 # Start from 1 for DictReader header row
                        for row in reader:
                            current_row_num += 1
                            if all(v is None for v in row.values()):
                                 print(f"DEBUG Load ({filename}): Skipping empty row {current_row_num}.")
                                 continue

                            # Use the actual found header name to get the key
                            key = row.get(key_header_actual, "").strip() if key_header_actual else f"row_{current_row_num}"
                            if not key:
                                print(f"DEBUG Load ({filename}): Skipping row {current_row_num} due to empty key.")
                                continue

                            if value_column:
                                # Use the actual found header name to get the value
                                value = row.get(value_header_actual, "").strip()
                                temp_data[key] = value
                            else:
                                # Clean keys and values when storing the whole row
                                temp_data[key] = {k.strip() if k else f"unknown_header_{idx}": v.strip() if v is not None else "" for idx, (k, v) in enumerate(row.items())}

                        data = temp_data
                        loaded = True
                        rows_processed = len(data)
                        print(f"DEBUG Load ({filename}): Successfully processed {rows_processed} data rows with encoding '{encoding}'.")
                        break # Break encoding loop

                    else: # Simple list loading
                        plain_reader = csv.reader(infile)
                        header_row = None
                        try:
                             header_row = next(plain_reader)
                             print(f"DEBUG Load ({filename}): Skipped header row for list loading: {header_row}")
                        except StopIteration:
                             print(f"DEBUG Load ({filename}): File is empty.")
                             loaded = True
                             break
                        except Exception as e_hdr:
                             print(f"DEBUG Load ({filename}): Error reading header for list loading: {e_hdr}")

                        temp_data = [row[0].strip() for row in plain_reader if row and len(row) > 0 and row[0] is not None and row[0].strip()]
                        rows_processed = len(temp_data)
                        data = temp_data
                        loaded = True
                        print(f"DEBUG Load ({filename}): Successfully processed {rows_processed} list rows with encoding '{encoding}'.")
                        break # Break encoding loop

            except UnicodeDecodeError:
                print(f"DEBUG Load ({filename}): Failed decoding with {encoding}.")
                last_error = f"UnicodeDecodeError with {encoding}"
                rows_processed = 0
                data = {} if is_dict else []
                continue # Try next encoding
            except csv.Error as e_csv:
                 print(f"CSV Error loading {filename} with {encoding} near line {reader.line_num if 'reader' in locals() and hasattr(reader, 'line_num') else 'unknown'}: {e_csv}")
                 last_error = f"CSV Error with {encoding}: {e_csv}"
                 rows_processed = 0
                 data = {} if is_dict else []
                 continue # Try next encoding
            except Exception as e:
                print(f"Unexpected Error loading {filename} with {encoding}: {e}")
                traceback.print_exc() # Print full traceback
                last_error = f"Unexpected error with {encoding}: {e}"
                print(f"DEBUG Load ({filename}): Returning None due to unexpected exception during file processing with {encoding}.")
                return None # Return None on critical error

        if not loaded:
            print(f"Failed to decode or load {filename} with any supported encoding or required columns missing. Last error: {last_error}")
            print(f"DEBUG Load ({filename}): Returning None because loading failed for all encodings.")
            return None # Indicate failure

        print(f"DEBUG Load ({filename}): Successfully loaded. Final data size: {len(data)}")
        return data

    def load_customers(self):
        """Loads customer data using the generic loader."""
        print("--- Loading Customers ---")
        # Ensure fallback to empty dict if loading fails
        return self._load_csv_generic("customers.csv", required_headers=["Name"], key_column="Name", is_dict=True) or {}

    def load_products(self, filename='products.csv'):
        """Loads product data using the generic loader and post-processes."""
        print("--- Loading Products ---")
        # --- CORRECTED key_column and required_headers ---
        product_rows = self._load_csv_generic(filename,
                                              required_headers=["ProductName", "ProductCode", "Price"], # Updated optional headers
                                              key_column="ProductName", # CORRECTED column name
                                              is_dict=True)
        # --- END CORRECTION ---

        # Return {} if load fails
        if product_rows is None:
             print(f"ERROR ({filename}): _load_csv_generic returned None. Returning empty dict.")
             return {} # Handle load failure explicitly, return empty dict

        products_dict = {}
        # Determine actual headers used for Code and Price (case-insensitive)
        first_row_keys = list(product_rows.get(next(iter(product_rows), ''), {}).keys())
        # Use the actual headers found in the file for lookups
        code_header = next((k for k in first_row_keys if k.lower().strip() == 'productcode'), None) # Match actual header
        price_header = next((k for k in first_row_keys if k.lower().strip() == 'price'), None) # Match actual header

        if not code_header: print(f"Warning ({filename}): Could not find 'ProductCode' column header in loaded data.")
        if not price_header: print(f"Warning ({filename}): Could not find 'Price' column header in loaded data.")

        processed_count = 0
        for name, row_dict in product_rows.items():
             # Use the found headers (code_header, price_header) to get values
            code = row_dict.get(code_header, "") if code_header else ""
            price_str = row_dict.get(price_header, "0.00") if price_header else "0.00"
            price_str_cleaned = re.sub(r'[^\d.-]', '', price_str)
            try:
                 _ = float(price_str_cleaned) if price_str_cleaned else 0.0
                 products_dict[name] = (code, price_str)
                 processed_count += 1
            except ValueError:
                 print(f"Warning ({filename}): Invalid price format '{price_str}' for product '{name}'. Using '0.00'.")
                 products_dict[name] = (code, "0.00")
                 processed_count += 1 # Count even if price is bad

        print(f"DEBUG ({filename}): Processed {processed_count} product rows into final dict size: {len(products_dict)}")
        return products_dict # Return the processed dict (might be empty)

    def load_parts(self):
        """Loads parts data using the generic loader."""
        print("--- Loading Parts ---")
        # Return {} if load fails
        return self._load_csv_generic("parts.csv",
                                      required_headers=["Part Name", "Part Number"],
                                      key_column="Part Name",
                                      value_column="Part Number",
                                      is_dict=True) or {}

    def load_salesmen_emails(self):
        """Loads salesmen data using the generic loader."""
        print("--- Loading Salesmen ---")
        # Ensure fallback to empty dict if loading fails
        return self._load_csv_generic("salesmen.csv",
                                      required_headers=["Name", "Email"],
                                      key_column="Name",
                                      value_column="Email",
                                      is_dict=True) or {}

    # --- End of Data Loading Methods ---


    # --- setup_ui METHOD with QScrollArea ---
    def setup_ui(self):
        # This method now relies on the dictionaries being non-None (they are initialized to {} if loading fails)
        main_layout_for_self = QVBoxLayout(self); main_layout_for_self.setContentsMargins(0,0,0,0)
        scroll_area = QScrollArea(self); scroll_area.setWidgetResizable(True); scroll_area.setStyleSheet("QScrollArea { border: none; background-color: #f1f3f5; }")
        content_container_widget = QWidget(); content_layout = QVBoxLayout(content_container_widget); content_layout.setSpacing(15); content_layout.setContentsMargins(15, 15, 15, 15)
        header = QWidget(); header.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #367C2B, stop:1 #4a9c3e); padding: 15px; border-bottom: 2px solid #2a5d24;"); header_layout = QHBoxLayout(header)

        # Logo
        logo_label = QLabel(self)
        logo_path_local = None
        if self._data_path: # Try loading logo from data path first
             logo_path_local = os.path.join(self._data_path, "logo.png")
             if not os.path.exists(logo_path_local):
                  logo_path_local = None # Fallback to script dir if not in data
        if not logo_path_local: # Try script directory
             try:
                  script_dir = os.path.dirname(os.path.abspath(__file__))
                  logo_path_local = os.path.join(script_dir, "logo.png")
             except NameError: # If __file__ not defined
                  logo_path_local = "logo.png" # Assume current dir

        logo_pixmap = QtGui.QPixmap(logo_path_local)
        if not logo_pixmap.isNull():
             logo_label.setPixmap(logo_pixmap.scaled(80, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation))
             print(f"DEBUG: Logo loaded from {logo_path_local}")
        else:
             logo_label.setText("Logo Missing"); logo_label.setStyleSheet("color: white;")
             print(f"DEBUG: Logo file not found or invalid at {logo_path_local}")

        title_label = QLabel("AMS Deal Form"); title_label.setStyleSheet("color: white; font-size: 40px; font-weight: bold; font-family: Arial;"); header_layout.addWidget(logo_label); header_layout.addWidget(title_label); header_layout.addStretch(); content_layout.addWidget(header)

        # Customer/Salesperson Group
        customer_group = QGroupBox("Customer & Salesperson"); customer_layout = QHBoxLayout(customer_group); self.customer_name = QLineEdit(); self.customer_name.setPlaceholderText("Customer Name"); customer_completer = QCompleter(list(self.customers_dict.keys()), self); customer_completer.setCaseSensitivity(Qt.CaseInsensitive); customer_completer.setFilterMode(Qt.MatchContains); self.customer_name.setCompleter(customer_completer); self.salesperson = QLineEdit(); self.salesperson.setPlaceholderText("Salesperson"); salesperson_completer = QCompleter(list(self.salesmen_emails.keys()), self); salesperson_completer.setCaseSensitivity(Qt.CaseInsensitive); salesperson_completer.setFilterMode(Qt.MatchContains); self.salesperson.setCompleter(salesperson_completer); customer_layout.addWidget(self.customer_name); customer_layout.addWidget(self.salesperson); content_layout.addWidget(customer_group)

        # Equipment Group
        equipment_group = QGroupBox("Equipment"); equipment_layout = QVBoxLayout(equipment_group); equipment_input_layout = QHBoxLayout(); self.equipment_product_name = QLineEdit(); self.equipment_product_name.setPlaceholderText("Product Name"); equipment_completer = QCompleter(list(self.products_dict.keys()), self); equipment_completer.setCaseSensitivity(Qt.CaseInsensitive); equipment_completer.setFilterMode(Qt.MatchContains); self.equipment_product_name.setCompleter(equipment_completer); equipment_completer.activated.connect(self.on_equipment_selected); self.equipment_product_code = QLineEdit(); self.equipment_product_code.setPlaceholderText("Product Code"); self.equipment_product_code.setReadOnly(True); self.equipment_product_code.setToolTip("Product Code from products.csv (auto-filled)"); self.equipment_manual_stock = QLineEdit(); self.equipment_manual_stock.setPlaceholderText("Stock # (Manual)"); self.equipment_manual_stock.setToolTip("Enter the specific Stock Number for this deal"); self.equipment_price = QLineEdit(); self.equipment_price.setPlaceholderText("$0.00"); validator_eq = QDoubleValidator(0.0, 9999999.99, 2); validator_eq.setNotation(QDoubleValidator.StandardNotation); self.equipment_price.setValidator(validator_eq); self.equipment_price.editingFinished.connect(self.format_price); equipment_add_btn = QPushButton("Add"); equipment_add_btn.clicked.connect(lambda: self.add_item("equipment")); equipment_input_layout.addWidget(self.equipment_product_name, 3); equipment_input_layout.addWidget(self.equipment_product_code, 1); equipment_input_layout.addWidget(self.equipment_manual_stock, 1); equipment_input_layout.addWidget(self.equipment_price, 1); equipment_input_layout.addWidget(equipment_add_btn, 0); self.equipment_list = QListWidget(); self.equipment_list.setSelectionMode(QListWidget.SingleSelection); self.equipment_list.itemDoubleClicked.connect(self.edit_equipment_item); equipment_layout.addLayout(equipment_input_layout); equipment_layout.addWidget(self.equipment_list); content_layout.addWidget(equipment_group)

        # Trades Group
        trades_group = QGroupBox("Trades"); trades_layout = QVBoxLayout(trades_group); trades_input_layout = QHBoxLayout(); self.trade_name = QLineEdit(); self.trade_name.setPlaceholderText("Trade Item"); trade_completer = QCompleter(list(self.products_dict.keys()), self); trade_completer.setCaseSensitivity(Qt.CaseInsensitive); trade_completer.setFilterMode(Qt.MatchContains); self.trade_name.setCompleter(trade_completer); trade_completer.activated.connect(self.on_trade_selected); self.trade_stock = QLineEdit(); self.trade_stock.setPlaceholderText("Stock #"); self.trade_amount = QLineEdit(); self.trade_amount.setPlaceholderText("$0.00"); validator_tr = QDoubleValidator(0.0, 9999999.99, 2); validator_tr.setNotation(QDoubleValidator.StandardNotation); self.trade_amount.setValidator(validator_tr); self.trade_amount.editingFinished.connect(self.format_amount); trades_add_btn = QPushButton("Add"); trades_add_btn.clicked.connect(lambda: self.add_item("trade")); trades_input_layout.addWidget(self.trade_name); trades_input_layout.addWidget(self.trade_stock); trades_input_layout.addWidget(self.trade_amount); trades_input_layout.addWidget(trades_add_btn); self.trade_list = QListWidget(); self.trade_list.setSelectionMode(QListWidget.SingleSelection); self.trade_list.itemDoubleClicked.connect(self.edit_trade_item); trades_layout.addLayout(trades_input_layout); trades_layout.addWidget(self.trade_list); content_layout.addWidget(trades_group)

        # Parts Group
        parts_group = QGroupBox("Parts"); parts_layout = QVBoxLayout(parts_group); parts_input_layout = QHBoxLayout(); self.part_quantity = QSpinBox(); self.part_quantity.setValue(1); self.part_quantity.setMinimum(1); self.part_quantity.setMaximum(999); self.part_quantity.setFixedWidth(60); self.part_number = QLineEdit(); self.part_number.setPlaceholderText("Part #"); part_number_completer = QCompleter(list(self.parts_dict.values()), self); part_number_completer.setCaseSensitivity(Qt.CaseInsensitive); part_number_completer.setFilterMode(Qt.MatchContains); self.part_number.setCompleter(part_number_completer); part_number_completer.activated.connect(self.on_part_number_selected); self.part_name = QLineEdit(); self.part_name.setPlaceholderText("Part Name"); part_name_completer = QCompleter(list(self.parts_dict.keys()), self); part_name_completer.setCaseSensitivity(Qt.CaseInsensitive); part_name_completer.setFilterMode(Qt.MatchContains); self.part_name.setCompleter(part_name_completer); part_name_completer.activated.connect(self.on_part_selected); self.part_location = QComboBox(); self.part_location.addItems(["", "Camrose", "Killam", "Wainwright", "Provost"]); self.part_charge_to = QLineEdit(); self.part_charge_to.setPlaceholderText("Charge to:"); parts_add_btn = QPushButton("Add"); parts_add_btn.clicked.connect(lambda: self.add_item("part")); parts_input_layout.addWidget(self.part_quantity); parts_input_layout.addWidget(self.part_number); parts_input_layout.addWidget(self.part_name); parts_input_layout.addWidget(self.part_location); parts_input_layout.addWidget(self.part_charge_to); parts_input_layout.addWidget(parts_add_btn); self.part_list = QListWidget(); self.part_list.setSelectionMode(QListWidget.SingleSelection); self.part_list.itemDoubleClicked.connect(self.edit_part_item); parts_layout.addLayout(parts_input_layout); parts_layout.addWidget(self.part_list); content_layout.addWidget(parts_group)

        # Work Order Group
        work_order_group = QGroupBox("Work Order & Options"); work_order_layout = QHBoxLayout(work_order_group); self.work_order_required = QCheckBox("Work Order Req'd?"); self.work_order_charge_to = QLineEdit(); self.work_order_charge_to.setPlaceholderText("Charge to:"); self.work_order_charge_to.setFixedWidth(150); self.work_order_hours = QLineEdit(); self.work_order_hours.setPlaceholderText("Duration (hours)"); self.multi_line_csv = QCheckBox("Multiple CSV Lines"); self.paid_checkbox = QCheckBox("Paid"); self.paid_checkbox.setStyleSheet("font-size: 16px; color: #333;"); self.paid_checkbox.setChecked(False); work_order_layout.addWidget(self.work_order_required); work_order_layout.addWidget(self.work_order_charge_to); work_order_layout.addWidget(self.work_order_hours); work_order_layout.addWidget(self.multi_line_csv); work_order_layout.addStretch(); work_order_layout.addWidget(self.paid_checkbox); content_layout.addWidget(work_order_group)

        # Buttons Layout
        buttons_layout = QHBoxLayout(); self.delete_line_btn = QPushButton("Delete Selected Line"); self.delete_line_btn.setToolTip("Delete the selected line from the list above (Equipment, Trade, or Part)"); self.delete_line_btn.clicked.connect(self.delete_selected_list_item); buttons_layout.addWidget(self.delete_line_btn); buttons_layout.addStretch(1); self.save_draft_btn = QPushButton("Save Draft"); self.save_draft_btn.setToolTip("Save the current form entries as a draft"); self.save_draft_btn.clicked.connect(self.save_draft); buttons_layout.addWidget(self.save_draft_btn); self.load_draft_btn = QPushButton("Load Draft"); self.load_draft_btn.setToolTip("Load the last saved draft"); self.load_draft_btn.clicked.connect(self.load_draft); buttons_layout.addWidget(self.load_draft_btn); buttons_layout.addSpacing(20); self.generate_csv_btn = QPushButton("Gen. CSV & Save"); self.generate_csv_btn.clicked.connect(self.generate_csv); self.generate_email_btn = QPushButton("Gen. Email"); self.generate_email_btn.clicked.connect(self.generate_email); self.generate_both_btn = QPushButton("Generate All"); self.generate_both_btn.clicked.connect(self.generate_csv_and_email); self.reset_btn = QPushButton("Reset Form"); self.reset_btn.clicked.connect(self.reset_form); self.reset_btn.setObjectName("reset_btn"); buttons_layout.addWidget(self.generate_csv_btn); buttons_layout.addSpacing(10); buttons_layout.addWidget(self.generate_email_btn); buttons_layout.addSpacing(10); buttons_layout.addWidget(self.generate_both_btn); buttons_layout.addSpacing(10); buttons_layout.addWidget(self.reset_btn); content_layout.addLayout(buttons_layout)
        content_layout.addStretch(); scroll_area.setWidget(content_container_widget); main_layout_for_self.addWidget(scroll_area)

    # --- Helper methods ---
    # ... (update_charge_to_default, format_price, format_amount unchanged) ...
    def update_charge_to_default(self):
        if self.equipment_list.count() > 0:
            line = self.equipment_list.item(0).text()
            match = re.search(r'STK#(\S+)', line)
            if match:
                stock_number = match.group(1)
                if not self.part_charge_to.text():
                    self.part_charge_to.setText(stock_number)
                self.last_charge_to = stock_number
            else:
                print(f"Error parsing equipment line (regex failed): {line}")

    def format_price(self):
        sender = self.sender(); text = sender.text(); cleaned_text = ''.join(c for c in text if c.isdigit() or c == '.' or (c == '-' and text.startswith('-')))
        try: value = float(cleaned_text) if cleaned_text and cleaned_text != '-' else 0.0; sender.setText(f"${value:,.2f}")
        except ValueError: sender.setText("$0.00")

    def format_amount(self):
        sender = self.sender(); text = sender.text(); cleaned_text = ''.join(c for c in text if c.isdigit() or c == '.' or (c == '-' and text.startswith('-')))
        try: value = float(cleaned_text) if cleaned_text and cleaned_text != '-' else 0.0; sender.setText(f"${value:,.2f}")
        except ValueError: sender.setText("$0.00")

    # --- Edit item methods (Using QListWidget now) ---
    # ... (Edit item methods remain unchanged) ...
    def edit_equipment_item(self, item):
        current_text = item.text(); pattern = r'"(.*)"\s+\(Code:\s*(.*?)\)\s+STK#(.*?)\s+\$(.*)'; match = re.match(pattern, current_text)
        if not match: QMessageBox.warning(self, "Edit Error", "Could not parse item text."); return
        name, code, manual_stock, price_str = match.groups(); price = price_str.strip()
        new_name, ok = QInputDialog.getText(self, "Edit Equipment", "Product Name:", text=name);
        if not ok: return
        # Use .get with a default tuple matching the structure if key not found
        new_code, new_default_price_str = self.products_dict.get(new_name, (code, price_str)) # Provide defaults
        new_manual_stock, ok = QInputDialog.getText(self, "Edit Equipment", "Manual Stock #:", text=manual_stock);
        if not ok: return
        current_price_val_str = price.replace(',', '')
        new_price_input_str, ok = QInputDialog.getText(self, "Edit Equipment", "Price (e.g., 1234.50):", text=current_price_val_str);
        if not ok: return
        try: price_value = float(new_price_input_str) if new_price_input_str else 0.0; new_price_formatted = f"{price_value:,.2f}"
        except ValueError: QMessageBox.warning(self, "Invalid Price", f"Invalid price: '{new_price_input_str}'."); new_price_formatted = "0.00"
        item.setText(f'"{new_name}" (Code: {new_code}) STK#{new_manual_stock} ${new_price_formatted}')
        self.update_charge_to_default()

    def edit_trade_item(self, item):
        current_text = item.text()
        try: name_part, stock_part = current_text.split(" STK#", 1); name = name_part.strip('" '); stock, amount_str = stock_part.split(" $", 1); amount = amount_str.strip()
        except ValueError: QMessageBox.warning(self, "Edit Error", "Could not parse item text."); return
        new_name, ok = QInputDialog.getText(self, "Edit Trade", "Trade Item:", text=name);
        if not ok: return
        new_stock, ok = QInputDialog.getText(self, "Edit Trade", "Stock #:", text=stock);
        if not ok: return
        current_amount_val_str = amount.replace(',', '')
        new_amount_str, ok = QInputDialog.getText(self, "Edit Trade", "Amount (e.g., 500.00):", text=current_amount_val_str);
        if not ok: return
        try: amount_value = float(new_amount_str) if new_amount_str else 0.0; new_amount_formatted = f"${amount_value:,.2f}"
        except ValueError: QMessageBox.warning(self, "Invalid Amount", f"Invalid amount: '{new_amount_str}'."); new_amount_formatted = "$0.00"
        item.setText(f'"{new_name}" STK#{new_stock} {new_amount_formatted}')

    def edit_part_item(self, item):
        current_text = item.text(); parts = current_text.split(" ", 4)
        if len(parts) < 4: QMessageBox.warning(self, "Edit Error", "Could not parse item text."); return
        try: qty = int(parts[0].rstrip('x'))
        except ValueError: qty = 1
        number = parts[1]; name = parts[2]; location = parts[3]; charge_to = parts[4] if len(parts) > 4 else ""
        new_qty, ok = QInputDialog.getInt(self, "Edit Part", "Quantity:", qty, 1, 999);
        if not ok: return
        new_number, ok = QInputDialog.getText(self, "Edit Part", "Part #:", text=number);
        if not ok: return
        new_name, ok = QInputDialog.getText(self, "Edit Part", "Part Name:", text=name);
        if not ok: return
        locations = ["", "Camrose", "Killam", "Wainwright", "Provost"]; current_loc_index = locations.index(location) if location in locations else 0
        new_location, ok = QInputDialog.getItem(self, "Edit Part", "Location:", locations, current=current_loc_index, editable=False);
        if not ok: return
        new_charge_to, ok = QInputDialog.getText(self, "Edit Part", "Charge to:", text=charge_to);
        if not ok: return
        item.setText(f"{new_qty}x {new_number} {new_name} {new_location} {new_charge_to}")


    # --- Autocompletion Handlers ---
    # ... (Autocompletion handlers remain unchanged) ...
    def on_equipment_selected(self, text):
        """Handles selection from equipment completer. Fills Code and Price."""
        print(f"Equipment selected: '{text}'")
        # Use .get with a default tuple to avoid KeyError if text not found
        product_data = self.products_dict.get(text, (None, None))
        code, price_str = product_data

        if code is not None: # Check if product was found
            current_price = self.equipment_price.text().replace('$','').replace(',','')
            self.equipment_product_code.setText(code)
            try:
                # Only set price if current price is empty or zero
                if not current_price or float(current_price) == 0.0:
                    price_value = float(re.sub(r'[^\d.-]', '', price_str)) if price_str else 0.0
                    self.equipment_price.setText(f"${price_value:,.2f}")
            except ValueError:
                print(f"Invalid price format '{price_str}' for product '{text}'")
                self.equipment_price.setText("$0.00")
            self.equipment_manual_stock.setFocus() # Move focus even if price wasn't updated
        else:
            print(f"Product '{text}' not found in products_dict.")
            self.equipment_product_code.clear()


    def on_trade_selected(self, text):
        """Handles selection from trade completer, fills amount/stock if empty."""
        print(f"Trade selected: '{text}'")
        # Use .get with a default tuple
        product_data = self.products_dict.get(text, (None, None))
        code, price_str = product_data

        if code is not None:
            current_amount = self.trade_amount.text().replace('$','').replace(',','')
            current_stock = self.trade_stock.text()
            try:
                # Only set amount if current amount is empty or zero
                if not current_amount or float(current_amount) == 0.0:
                    price_value = float(re.sub(r'[^\d.-]', '', price_str)) if price_str else 0.0
                    self.trade_amount.setText(f"${price_value:,.2f}")
            except ValueError:
                print(f"Invalid price format '{price_str}' for product '{text}'")
                self.trade_amount.setText("$0.00")
            # Only set stock if current stock is empty
            if not current_stock and code:
                self.trade_stock.setText(code)
        else:
             print(f"Product '{text}' not found in products_dict for trade.")


    def on_part_selected(self, text):
        """Handles selection from part name completer, fills part number."""
        text = text.strip(); print(f"Part name selected: '{text}'")
        # Use .get with a default value (e.g., None or "")
        part_number = self.parts_dict.get(text)
        if part_number:
             # Only fill if part number field is empty
             if not self.part_number.text():
                  self.part_number.setText(part_number)
        else: print(f"Part name '{text}' not found in parts_dict.")

    def on_part_number_selected(self, text):
        """Handles selection from part number completer, fills part name."""
        text = text.strip(); print(f"Part number selected: '{text}'")
        found = False
        # Check if parts_dict exists and is iterable
        if self.parts_dict and isinstance(self.parts_dict, dict):
            for part_name, part_num in self.parts_dict.items():
                if part_num == text:
                     # Only fill if part name field is empty
                     if not self.part_name.text():
                          self.part_name.setText(part_name)
                     found = True
                     break # Found the first match
        if not found: print(f"Part number '{text}' not found in parts_dict values.")


    # --- Add Item Logic (Using QListWidget now) ---
    # ... (Add item logic remains unchanged) ...
    def add_item(self, item_type):
        """Adds an item to the corresponding list widget."""
        if item_type == "equipment":
            name = self.equipment_product_name.text().strip()
            code = self.equipment_product_code.text().strip()
            manual_stock = self.equipment_manual_stock.text().strip()
            price_text = self.equipment_price.text().strip()
            if not name: QMessageBox.warning(self, "Missing Info", "Please enter or select a Product Name."); return
            if not manual_stock: QMessageBox.warning(self, "Missing Info", "Please enter a manual Stock Number."); return
            try: price = price_text if price_text.startswith('$') else f"${float(price_text.replace(',', '')) if price_text else 0.0:,.2f}"
            except ValueError: price = "$0.00"
            item_text = f'"{name}" (Code: {code}) STK#{manual_stock} {price}'
            QListWidgetItem(item_text, self.equipment_list)
            self.equipment_product_name.clear(); self.equipment_product_code.clear(); self.equipment_manual_stock.clear(); self.equipment_price.clear()
            self.update_charge_to_default(); self.equipment_product_name.setFocus()
        elif item_type == "trade":
            name = self.trade_name.text().strip(); stock = self.trade_stock.text().strip(); amount_text = self.trade_amount.text().strip()
            if name:
                try: amount = amount_text if amount_text.startswith('$') else f"${float(amount_text.replace(',', '')) if amount_text else 0.0:,.2f}"
                except ValueError: amount = "$0.00"
                item_text = f'"{name}" STK#{stock} {amount}'
                QListWidgetItem(item_text, self.trade_list)
                self.trade_name.clear(); self.trade_stock.clear(); self.trade_amount.clear(); self.trade_name.setFocus()
        elif item_type == "part":
            qty = str(self.part_quantity.value()); number = self.part_number.text().strip(); name = self.part_name.text().strip()
            location = self.part_location.currentText().strip(); charge_to = self.part_charge_to.text().strip()
            if name or number:
                item_text = f"{qty}x {number} {name} {location} {charge_to}"
                QListWidgetItem(item_text, self.part_list)
                if not self.last_charge_to and charge_to: self.last_charge_to = charge_to
                self.part_name.clear(); self.part_number.clear(); self.part_quantity.setValue(1)
                self.part_charge_to.setText(self.last_charge_to); self.part_number.setFocus()


    # --- Form Actions ---

    # --- reset_form (Modified to refresh completers) ---
    def reset_form(self):
        """Clears all input fields and lists, and refreshes completers."""
        reply = QMessageBox.question(self, 'Confirm Reset', "Are you sure?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            # Clear text fields
            self.customer_name.clear(); self.salesperson.clear()
            self.equipment_product_name.clear(); self.equipment_product_code.clear()
            self.equipment_manual_stock.clear(); self.equipment_price.clear()
            self.trade_name.clear(); self.trade_stock.clear(); self.trade_amount.clear()
            self.part_number.clear(); self.part_name.clear(); self.part_charge_to.clear()
            self.work_order_charge_to.clear(); self.work_order_hours.clear()

            # Clear lists
            self.equipment_list.clear(); self.trade_list.clear(); self.part_list.clear()

            # Reset other controls
            self.part_quantity.setValue(1); self.part_location.setCurrentIndex(0)
            self.work_order_required.setChecked(False); self.multi_line_csv.setChecked(False); self.paid_checkbox.setChecked(False)

            # --- Refresh Completers (with individual try/except and checks for None/empty) ---
            print("DEBUG: Refreshing completer models after reset...")
            # Customer
            try:
                customer_keys = list(self.customers_dict.keys()) if self.customers_dict else []
                print(f"DEBUG Reset: Customer completer list size: {len(customer_keys)}")
                completer = self.customer_name.completer()
                if completer and hasattr(completer, 'model') and callable(completer.model):
                    completer.model().setStringList(customer_keys)
                else: print("DEBUG Reset: Customer completer or model not found/valid.")
            except Exception as e: print(f"ERROR: Could not refresh Customer completer model: {e}")
            # Salesperson
            try:
                salesmen_keys = list(self.salesmen_emails.keys()) if self.salesmen_emails else []
                print(f"DEBUG Reset: Salesperson completer list size: {len(salesmen_keys)}")
                completer = self.salesperson.completer()
                if completer and hasattr(completer, 'model') and callable(completer.model):
                     completer.model().setStringList(salesmen_keys)
                else: print("DEBUG Reset: Salesperson completer or model not found/valid.")
            except Exception as e: print(f"ERROR: Could not refresh Salesperson completer model: {e}")
            # Equipment / Trade (Product Names)
            try:
                product_keys = list(self.products_dict.keys()) if self.products_dict else []
                print(f"DEBUG Reset: Equipment/Trade completer list size: {len(product_keys)}")
                eq_completer = self.equipment_product_name.completer()
                if eq_completer and hasattr(eq_completer, 'model') and callable(eq_completer.model):
                    eq_completer.model().setStringList(product_keys)
                else: print("DEBUG Reset: Equipment completer or model not found/valid.")
                tr_completer = self.trade_name.completer()
                if tr_completer and hasattr(tr_completer, 'model') and callable(tr_completer.model):
                    tr_completer.model().setStringList(product_keys)
                else: print("DEBUG Reset: Trade completer or model not found/valid.")
            except Exception as e: print(f"ERROR: Could not refresh Equipment/Trade completer models: {e}")
            # Part Name
            try:
                part_name_keys = list(self.parts_dict.keys()) if self.parts_dict else []
                print(f"DEBUG Reset: Part Name completer list size: {len(part_name_keys)}")
                completer = self.part_name.completer()
                if completer and hasattr(completer, 'model') and callable(completer.model):
                    completer.model().setStringList(part_name_keys)
                else: print("DEBUG Reset: Part Name completer or model not found/valid.")
            except Exception as e: print(f"ERROR: Could not refresh Part Name completer model: {e}")
            # Part Number
            try:
                # Ensure values are strings for the completer
                part_number_values = [str(v) for v in self.parts_dict.values()] if self.parts_dict else []
                print(f"DEBUG Reset: Part Number completer list size: {len(part_number_values)}")
                completer = self.part_number.completer()
                if completer and hasattr(completer, 'model') and callable(completer.model):
                    completer.model().setStringList(part_number_values)
                else: print("DEBUG Reset: Part Number completer or model not found/valid.")
            except Exception as e: print(f"ERROR: Could not refresh Part Number completer model: {e}")

            print("DEBUG: Completer models refresh attempt finished.")
            # --- End Refresh Completers ---

            # Reset internal state
            self.last_charge_to = ""; self.csv_lines = []
            print("Form reset!")
            self._show_status_message("Form Reset", 3000)

    # ... (delete_selected_list_item, save_draft, load_draft, populate_form unchanged) ...
    def delete_selected_list_item(self):
        """Deletes the selected item from whichever list has focus/selection."""
        focused_widget = QApplication.focusWidget(); target_list = None; list_name = ""
        # Determine target list based on focus or selection
        if isinstance(focused_widget, QListWidget):
             if focused_widget == self.equipment_list: target_list, list_name = self.equipment_list, "Equipment"
             elif focused_widget == self.trade_list: target_list, list_name = self.trade_list, "Trade"
             elif focused_widget == self.part_list: target_list, list_name = self.part_list, "Part"

        # Fallback if focus is not on a list, check selection
        if target_list is None:
             if self.equipment_list.currentRow() >= 0: target_list, list_name = self.equipment_list, "Equipment"
             elif self.trade_list.currentRow() >= 0: target_list, list_name = self.trade_list, "Trade"
             elif self.part_list.currentRow() >= 0: target_list, list_name = self.part_list, "Part"

        if target_list is None: QMessageBox.warning(self, "Delete Line", "Please select a line to delete."); return
        current_row = target_list.currentRow()
        if current_row < 0: QMessageBox.warning(self, "Delete Line", f"Please select a line in the {list_name} list."); return
        item = target_list.item(current_row); item_text = item.text() if item else f"Row {current_row + 1}"
        reply = QMessageBox.question(self, 'Confirm Delete', f"Delete this {list_name} line?\n\n'{item_text}'", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes: target_list.takeItem(current_row); self._show_status_message(f"{list_name} line deleted.", 3000)

    def save_draft(self):
        """Saves the current state of the form to a JSON file."""
        if not self.draft_file_path: QMessageBox.critical(self, "Error", "Draft file path not configured."); return
        draft_data = self._get_current_deal_data()
        try:
            with open(self.draft_file_path, 'w', encoding='utf-8') as f: json.dump(draft_data, f, indent=4)
            self._show_status_message("Draft saved successfully.", 3000); print(f"Draft saved to {self.draft_file_path}")
        except Exception as e: print(f"Error saving draft: {e}"); QMessageBox.critical(self, "Save Draft Error", f"Could not write draft file:\n{e}"); self._show_status_message(f"Error saving draft: {e}", 5000)

    def load_draft(self):
        """Loads the form state from the saved JSON draft file."""
        if not self.draft_file_path: QMessageBox.critical(self, "Error", "Draft file path not configured."); return
        if not os.path.exists(self.draft_file_path): QMessageBox.information(self, "Load Draft", "No draft file found."); self._show_status_message("No draft file found.", 3000); return
        try:
            with open(self.draft_file_path, 'r', encoding='utf-8') as f: draft_data = json.load(f)
            self.populate_form(draft_data); self._show_status_message("Draft loaded successfully.", 3000); print(f"Draft loaded from {self.draft_file_path}")
        except Exception as e: print(f"Error loading/populating draft: {e}"); QMessageBox.critical(self, "Load Draft Error", f"Could not read or parse draft file:\n{e}"); self._show_status_message(f"Error loading draft: {e}", 5000)

    def populate_form(self, deal_data):
        """Populates the form widgets with data from a dictionary."""
        print(f"DEBUG: Populating form with data for customer: {deal_data.get('customer_name')}")
        try:
            self.equipment_list.clear(); self.trade_list.clear(); self.part_list.clear()
            self.customer_name.setText(deal_data.get("customer_name", "")); self.salesperson.setText(deal_data.get("salesperson", ""))
            for item_text in deal_data.get("equipment", []): QListWidgetItem(item_text, self.equipment_list)
            for item_text in deal_data.get("trades", []): QListWidgetItem(item_text, self.trade_list)
            for item_text in deal_data.get("parts", []): QListWidgetItem(item_text, self.part_list)
            self.work_order_required.setChecked(deal_data.get("work_order_required", False)); self.work_order_charge_to.setText(deal_data.get("work_order_charge_to", "")); self.work_order_hours.setText(deal_data.get("work_order_hours", ""))
            self.multi_line_csv.setChecked(deal_data.get("multi_line_csv", False)); self.paid_checkbox.setChecked(deal_data.get("paid", False))
            self.part_location.setCurrentIndex(deal_data.get("part_location_index", 0)); self.last_charge_to = deal_data.get("last_charge_to", ""); self.part_charge_to.setText(self.last_charge_to)
            self.update_charge_to_default()
            self.equipment_product_name.clear(); self.equipment_product_code.clear(); self.equipment_manual_stock.clear(); self.equipment_price.clear()
            self.trade_name.clear(); self.trade_stock.clear(); self.trade_amount.clear()
            self.part_name.clear(); self.part_number.clear(); self.part_quantity.setValue(1)
            self._show_status_message("Form populated with selected deal data.", 3000)
        except Exception as e: print(f"Error populating form: {e}"); QMessageBox.critical(self, "Populate Error", f"An error occurred populating the form:\n{e}")


    # --- save_to_csv (SharePoint Save) ---
    def save_to_csv(self, csv_lines):
        """Save the CSV data to SharePoint"""
        if not self.sharepoint_manager:
             print("SharePoint Manager not initialized. Cannot save CSV.");
             self._show_status_message("SP connection failed.", 5000);
             # Optionally try saving locally as a fallback
             # self.save_csv_locally(csv_lines)
             return False
        if not csv_lines: QMessageBox.warning(self, "Save Error", "No data generated."); return False
        try:
            data_to_save = []; headers = ["Payment", "Customer", "Equipment", "Stock Number", "Amount", "Trade", "Attached to stk#", "Trade STK#", "Amount2", "Salesperson", "Email Date", "Status", "Timestamp"]
            csv_data = "\n".join(csv_lines); csvfile = io.StringIO(csv_data); reader = csv.reader(csvfile, delimiter=',', quotechar='"')
            for fields in reader:
                if len(fields) == len(headers): data_to_save.append(dict(zip(headers, fields)))
                else: print(f"Warning: Skipping malformed CSV line: {fields}")
            if not data_to_save: QMessageBox.warning(self, "Save Error", "Could not parse generated data."); return False

            # Ensure sharepoint_manager has the update_excel_data method
            if not hasattr(self.sharepoint_manager, 'update_excel_data'):
                 print("ERROR: SharePoint Manager object does not have 'update_excel_data' method.")
                 self._show_status_message("SP Save Error: Method missing.", 5000);
                 return False

            success = self.sharepoint_manager.update_excel_data(data_to_save)
            if success: self._show_status_message("Data saved to SharePoint!", 5000); return True
            else: self._show_status_message("Failed to save to SharePoint", 5000); return False
        except AttributeError as ae:
             print(f"AttributeError in save_to_csv (likely SP Manager issue): {ae}")
             QMessageBox.critical(self, "Save Error", f"SharePoint connection error:\n{ae}");
             self._show_status_message(f"SP Save Error: {ae}", 5000);
             return False
        except Exception as e:
             print(f"Error in save_to_csv: {e}");
             import traceback; traceback.print_exc() # Log full traceback
             QMessageBox.critical(self, "Save Error", f"Unexpected error during save:\n{e}");
             self._show_status_message(f"SP Save Error: {e}", 5000);
             return False


    # --- generate_csv (Modified to save recent deal) ---
    def generate_csv(self):
        """Generates the CSV content, saves deal to recent list, and attempts to save to SP."""
        print("Generating CSV...")
        customer = self.customer_name.text().strip(); salesperson = self.salesperson.text().strip()
        if not customer: QMessageBox.warning(self, "Missing Info", "Please enter a Customer Name."); return False
        if not salesperson: QMessageBox.warning(self, "Missing Info", "Please enter a Salesperson."); return False
        if self.equipment_list.count() == 0: QMessageBox.warning(self, "Missing Info", "Please add at least one Equipment item."); return False

        equipment_items = [self.equipment_list.item(i).text() for i in range(self.equipment_list.count())]
        trade_items = [self.trade_list.item(i).text() for i in range(self.trade_list.count())]
        today = QDate.currentDate().toString("yyyy-MM-dd"); status = "Paid" if self.paid_checkbox.isChecked() else "Not Paid"
        payment_icon = "🟩" if self.paid_checkbox.isChecked() else "🟥"; timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S PDT")
        self.csv_lines = []

        try:
            eq_pattern = r'"(.*)"\s+\(Code:\s*(.*?)\)\s+STK#(.*?)\s+\$(.*)'
            if self.multi_line_csv.isChecked():
                for line in equipment_items:
                    if line:
                        match = re.match(eq_pattern, line)
                        if match: equipment, _code, manual_stock, amount_str = match.groups(); amount = amount_str.strip()
                        else: raise ValueError(f"Cannot parse equipment line: {line}")
                        output = io.StringIO(); writer = csv.writer(output, quoting=csv.QUOTE_ALL)
                        writer.writerow([payment_icon, customer, equipment, manual_stock, amount, "", "", "", "", salesperson, today, status, timestamp])
                        self.csv_lines.append(output.getvalue().strip())
                first_eq_manual_stock = "";
                if equipment_items and equipment_items[0]: match = re.match(eq_pattern, equipment_items[0]);
                if match: first_eq_manual_stock = match.group(3)
                for line in trade_items:
                    if line:
                        # Handle potential ValueError if split fails
                        try:
                             name_part, stock_part = line.split(" STK#", 1); trade = name_part.strip('" '); trade_stock, trade_amount_str = stock_part.split(" $", 1); trade_amount = trade_amount_str.strip()
                        except ValueError as split_err:
                             raise ValueError(f"Cannot parse trade line: {line} ({split_err})")
                        attached_to = first_eq_manual_stock if first_eq_manual_stock else "N/A"
                        output = io.StringIO(); writer = csv.writer(output, quoting=csv.QUOTE_ALL)
                        writer.writerow([payment_icon, customer, "", "", "", trade, attached_to, trade_stock, trade_amount, salesperson, today, status, timestamp])
                        self.csv_lines.append(output.getvalue().strip())
            else: # Single line logic
                payment = payment_icon; equipment = stock_number = amount = trade = attached_to = trade_stock = trade_amount = ""
                if equipment_items and equipment_items[0]:
                    match = re.match(eq_pattern, equipment_items[0])
                    if match: equipment, _code, stock_number, amount_str = match.groups(); amount = amount_str.strip()
                    else: raise ValueError(f"Cannot parse equipment line: {equipment_items[0]}")
                if trade_items and trade_items[0]:
                    line = trade_items[0]
                    try:
                         name_part, stock_part = line.split(" STK#", 1); trade = name_part.strip('" '); trade_stock, trade_amount_str = stock_part.split(" $", 1); trade_amount = trade_amount_str.strip();
                    except ValueError as split_err:
                         raise ValueError(f"Cannot parse trade line: {line} ({split_err})")
                    attached_to = stock_number if stock_number else "N/A"
                output = io.StringIO(); writer = csv.writer(output, quoting=csv.QUOTE_ALL)
                writer.writerow([payment, customer, equipment, stock_number, amount, trade, attached_to, trade_stock, trade_amount, salesperson, today, status, timestamp])
                self.csv_lines.append(output.getvalue().strip())

            if not self.csv_lines: QMessageBox.warning(self, "CSV Error", "No valid CSV lines generated."); return False

            try: deal_data = self._get_current_deal_data(); self._save_deal_to_recent(deal_data)
            except Exception as recent_err: print(f"ERROR: Failed to save deal to recent list: {recent_err}"); QMessageBox.warning(self, "Recent Deals Error", f"Could not save deal to recent list:\n{recent_err}")

            print("CSV generated successfully! Attempting to save...");
            if self.save_to_csv(self.csv_lines): return True
            else: return False

        except Exception as e: print(f"Error during CSV generation: {e}"); import traceback; traceback.print_exc(); QMessageBox.critical(self, "CSV Generation Error", f"Error generating CSV:\n{e}"); self.csv_lines = []; return False

    # --- NEW: Get Current Deal Data ---
    def _get_current_deal_data(self):
        """Collects current form data into a dictionary."""
        part_items = [self.part_list.item(i).text() for i in range(self.part_list.count())]
        deal_data = {
            "timestamp": datetime.now().isoformat(),
            "customer_name": self.customer_name.text(), "salesperson": self.salesperson.text(),
            "equipment": [self.equipment_list.item(i).text() for i in range(self.equipment_list.count())],
            "trades": [self.trade_list.item(i).text() for i in range(self.trade_list.count())],
            "parts": part_items,
            "work_order_required": self.work_order_required.isChecked(), "work_order_charge_to": self.work_order_charge_to.text(), "work_order_hours": self.work_order_hours.text(),
            "multi_line_csv": self.multi_line_csv.isChecked(), "paid": self.paid_checkbox.isChecked(),
            "part_location_index": self.part_location.currentIndex(), "last_charge_to": self.last_charge_to
        }
        return deal_data

    # --- NEW: Save Deal to Recent List ---
    def _save_deal_to_recent(self, deal_data):
        """Loads recent deals, prepends new deal, trims list, saves back."""
        if not self.recent_deals_file: print("ERROR: Recent deals file path not set."); return
        recent_deals = []
        try:
            if os.path.exists(self.recent_deals_file):
                with open(self.recent_deals_file, 'r', encoding='utf-8') as f:
                    recent_deals = json.load(f)
                    if not isinstance(recent_deals, list): recent_deals = []
        except json.JSONDecodeError as json_err:
             print(f"Warning: Could not parse recent deals file (JSON invalid): {json_err}. Starting fresh list.")
             recent_deals = []
        except Exception as e:
             print(f"Warning: Could not load recent deals file: {e}. Starting fresh list.")
             recent_deals = []

        # Ensure deal_data is serializable (basic check)
        try:
             _ = json.dumps(deal_data) # Test serialization
        except TypeError as serial_err:
             print(f"ERROR: Deal data is not JSON serializable: {serial_err}")
             QMessageBox.critical(self, "Save Error", "Could not save recent deal: data contains non-serializable types.")
             return # Don't proceed if data is bad

        recent_deals.insert(0, deal_data); recent_deals = recent_deals[:10] # Keep max 10
        try:
            with open(self.recent_deals_file, 'w', encoding='utf-8') as f: json.dump(recent_deals, f, indent=4)
            print(f"Saved deal to recent list ({len(recent_deals)} items).")
        except Exception as e: print(f"Error saving recent deals file: {e}")

    # --- IMPROVED: Enhanced Email Template ---
    def generate_email(self):
        """Generate HTML email body with improved styling, show in preview dock, and open mailto link."""
        print("Generating email (Enhanced HTML Preview + Mailto)...")
        customer_name = self.customer_name.text().strip()
        salesperson = self.salesperson.text().strip()
        equipment_items = [self.equipment_list.item(i).text() for i in range(self.equipment_list.count())]
        trade_items = [self.trade_list.item(i).text() for i in range(self.trade_list.count())]
        part_items = [self.part_list.item(i).text() for i in range(self.part_list.count())]

        if not customer_name:
            QMessageBox.warning(self, "Missing Info", "Please enter a Customer Name.")
            return
        if not salesperson:
            QMessageBox.warning(self, "Missing Info", "Please enter a Salesperson.")
            return

        # --- Get First Equipment Info ---
        first_product_name = "N/A"
        first_stock_number = "N/A"
        eq_pattern = r'"(.*)"\s+\(Code:\s*(.*?)\)\s+STK#(.*?)\s+\$(.*)'
        if equipment_items:
            match = re.match(eq_pattern, equipment_items[0])
            if match:
                first_product_name, _, first_stock_number, _ = match.groups()
            else:
                try:
                    # Basic fallback if regex fails
                    first_product_name = equipment_items[0].split('"')[1] if '"' in equipment_items[0] else equipment_items[0].split(' ')[0]
                except IndexError:
                    first_product_name = equipment_items[0][:30] # Truncate if split fails

        # --- Subject ---
        subject = f"AMS Deal ({customer_name} - {first_product_name})"

        # --- Recipients ---
        fixed_recipients = [
            'bstdenis@briltd.com', 'cgoodrich@briltd.com', 'dvriend@briltd.com',
            'rbendfeld@briltd.com', 'bfreadrich@briltd.com', 'vedwards@briltd.com'
        ]
        # Use .get on the dictionary which handles None if loading failed
        salesman_email = self.salesmen_emails.get(salesperson) if self.salesmen_emails else None
        if salesman_email:
            fixed_recipients.append(salesman_email)
        else:
            print(f"Warning: Email for salesperson '{salesperson}' not found.")
            QMessageBox.warning(self, "Salesperson Email", f"Email for {salesperson} not found.")

        if part_items:
            fixed_recipients.append('rkrys@briltd.com')

        recipient_list = [r for r in fixed_recipients if r and '@' in r]
        recipients_string = ";".join(recipient_list)

        # --- HTML Body Construction with Enhanced Styling ---
        # CSS Styles
        styles = """
        <style>
            body {
                font-family: Arial, sans-serif;
                font-size: 12pt;
                color: #333333;
                line-height: 1.4;
                max-width: 800px;
                margin: 0 auto;
            }
            .header {
                background-color: #367C2B;
                color: white;
                padding: 15px;
                margin-bottom: 20px;
                border-radius: 5px;
            }
            .header h1 {
                font-size: 20pt;
                margin: 0;
            }
            .info-section {
                background-color: #f8f9fa;
                padding: 15px;
                margin-bottom: 20px;
                border-radius: 5px;
                border-left: 5px solid #367C2B;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 20px;
            }
            th {
                background-color: #367C2B;
                color: white;
                font-weight: bold;
                text-align: left;
                padding: 10px;
                border: 1px solid #ddd;
            }
            td {
                padding: 10px;
                border: 1px solid #ddd;
                vertical-align: top;
            }
            tr:nth-child(even) {
                background-color: #f2f2f2;
            }
            .section-title {
                background-color: #6c757d;
                color: white;
                padding: 8px 15px;
                font-size: 14pt;
                font-weight: bold;
                margin: 20px 0 10px 0;
                border-radius: 5px;
            }
            .money {
                text-align: right;
                font-family: 'Courier New', monospace;
            }
            .footer {
                background-color: #f8f9fa;
                padding: 15px;
                margin-top: 20px;
                border-top: 1px solid #ddd;
                font-style: italic;
                border-radius: 5px;
            }
            .paid {
                background-color: #d4edda;
                color: #155724;
                padding: 5px 10px;
                border-radius: 3px;
                font-weight: bold;
                display: inline-block;
            }
            .not-paid {
                background-color: #f8d7da;
                color: #721c24;
                padding: 5px 10px;
                border-radius: 3px;
                font-weight: bold;
                display: inline-block;
            }
            .note {
                background-color: #fff3cd;
                color: #856404;
                padding: 10px;
                margin: 10px 0;
                border-radius: 3px;
                border-left: 5px solid #ffeeba;
            }
        </style>
        """

        # Start building HTML
        body = []
        body.append("<html><head>")
        body.append(styles)
        body.append("</head><body>")

        # Header with AMS logo and deal info
        body.append("<div class='header'>")
        body.append(f"<h1>AMS Deal: {html.escape(customer_name)}</h1>")
        body.append("</div>")

        # Customer and salesperson info section
        body.append("<div class='info-section'>")
        body.append(f"<p><strong>Customer:</strong> {html.escape(customer_name)}</p>")
        body.append(f"<p><strong>Sales:</strong> {html.escape(salesperson)}</p>")

        # Payment status
        if self.paid_checkbox.isChecked():
            body.append("<p><span class='paid'>PAID</span></p>")
        else:
            body.append("<p><span class='not-paid'>NOT PAID</span></p>")
        body.append("</div>")

        # Equipment section
        body.append("<div class='section-title'>Equipment</div>")
        total_equipment_price = 0.0 # Initialize here
        if equipment_items:
            body.append("<table>")
            body.append("<thead><tr><th>Name</th><th>Stock #</th><th>Price</th></tr></thead><tbody>")

            for line in equipment_items:
                match = re.match(eq_pattern, line)
                if match:
                    name, _code, manual_stock, price_str = match.groups()
                    try:
                        price_float = float(price_str.replace(',', ''))
                        total_equipment_price += price_float
                    except ValueError:
                        price_float = 0.0

                    body.append(f"<tr>")
                    body.append(f"<td>{html.escape(name)}</td>")
                    body.append(f"<td>{html.escape(manual_stock)}</td>")
                    body.append(f"<td class='money'>${html.escape(price_str)}</td>")
                    body.append(f"</tr>")
                else:
                    body.append(f"<tr><td colspan='3'><i>Error parsing: {html.escape(line)}</i></td></tr>")

            # Add total row
            body.append(f"<tr><td colspan='2'><strong>Total Equipment</strong></td>")
            body.append(f"<td class='money'><strong>${total_equipment_price:,.2f}</strong></td></tr>")
            body.append("</tbody></table>")
        else:
            body.append("<p>No equipment items added.</p>")

        # Trades section
        total_trade_amount = 0.0 # Initialize here
        if trade_items:
            body.append("<div class='section-title'>Trade-Ins</div>")
            body.append("<table>")
            body.append("<thead><tr><th>Item</th><th>Stock #</th><th>Amount</th></tr></thead><tbody>")

            for line in trade_items:
                try:
                    name_part, stock_part = line.split(" STK#", 1)
                    name = name_part.strip('" ')
                    stock, amount_str = stock_part.split(" $", 1)
                    amount_str = amount_str.strip()

                    try:
                        amount_float = float(amount_str.replace(',', ''))
                        total_trade_amount += amount_float
                    except ValueError:
                        amount_float = 0.0

                    body.append(f"<tr>")
                    body.append(f"<td>{html.escape(name)}</td>")
                    body.append(f"<td>{html.escape(stock)}</td>")
                    body.append(f"<td class='money'>${html.escape(amount_str)}</td>")
                    body.append(f"</tr>")
                except ValueError:
                    body.append(f"<tr><td colspan='3'><i>Error parsing: {html.escape(line)}</i></td></tr>")

            # Add total row
            body.append(f"<tr><td colspan='2'><strong>Total Trade-In</strong></td>")
            body.append(f"<td class='money'><strong>${total_trade_amount:,.2f}</strong></td></tr>")
            body.append("</tbody></table>")

        # Calculate net amount if both equipment and trades exist
        if equipment_items and trade_items:
            try:
                net_amount = total_equipment_price - total_trade_amount
                body.append("<div class='note'>")
                body.append(f"<p><strong>Net Amount After Trade:</strong> <span style='font-size: 14pt;'>${net_amount:,.2f}</span></p>")
                body.append("</div>")
            except Exception as e:
                print(f"Error calculating net amount: {e}")

        # Parts section (grouped by location)
        if part_items:
            body.append("<div class='section-title'>Parts</div>")

            part_groups = {}
            for line in part_items:
                parts = line.split(" ", 4)
                if len(parts) >= 4:
                    qty = parts[0].rstrip('x')
                    number = parts[1]
                    name = parts[2]
                    location = parts[3]
                    charge_to = parts[4] if len(parts) > 4 else first_stock_number

                    loc_key = location if location else "Unspecified Location"
                    part_groups.setdefault(loc_key, []).append({
                        "qty": qty,
                        "number": number,
                        "name": name,
                        "charge_to": charge_to
                    })
                else:
                    part_groups.setdefault("Format Error", []).append(line)

            for location, parts_list in part_groups.items():
                body.append(f"<h3>Parts from {html.escape(location)}</h3>")
                body.append("<table>")
                body.append("<thead><tr><th>Qty</th><th>Part #</th><th>Description</th><th>Charge To</th></tr></thead><tbody>")

                for part in parts_list:
                    if isinstance(part, dict):  # Properly parsed part
                        body.append(f"<tr>")
                        body.append(f"<td style='text-align: center;'>{html.escape(part['qty'])}</td>")
                        body.append(f"<td>{html.escape(part['number'])}</td>")
                        body.append(f"<td>{html.escape(part['name'])}</td>")
                        body.append(f"<td>{html.escape(part['charge_to'])}</td>")
                        body.append(f"</tr>")
                    else:  # Error case
                        body.append(f"<tr><td colspan='4'><i>Error parsing: {html.escape(part)}</i></td></tr>")

                body.append("</tbody></table>")

        # Work Order section
        if self.work_order_required.isChecked():
            duration = self.work_order_hours.text().strip() or "[Duration Not Specified]"
            charge_to_wo = self.work_order_charge_to.text().strip() or first_stock_number

            body.append("<div class='section-title'>Work Order</div>")
            body.append("<div class='note'>")
            body.append(f"<p><strong>Please create a work order:</strong></p>")
            body.append(f"<ul>")
            body.append(f"<li><strong>Duration:</strong> {html.escape(duration)} hours</li>")
            body.append(f"<li><strong>Charge to:</strong> {html.escape(charge_to_wo)}</li>")
            body.append(f"</ul>")
            body.append("</div>")

        # Footer with additional information
        body.append("<div class='footer'>")
        body.append(f"<p>PFW and spreadsheet have been updated. {html.escape(salesperson)} to collect.</p>")
        body.append(f"<p><small>Generated by AMS Deal Form on {datetime.now().strftime('%Y-%m-%d %H:%M')}</small></p>")
        body.append("</div>")

        body.append("</body></html>")
        html_body = "".join(body)

        # --- Action 1: Show HTML Preview in Dock ---
        if self.main_window and hasattr(self.main_window, 'email_preview_view') and hasattr(self.main_window, 'email_preview_dock'):
            try:
                # Use setHtml for QTextEdit to render HTML
                self.main_window.email_preview_view.setHtml(html_body)
                self.main_window.email_preview_dock.setVisible(True)
                self.main_window.email_preview_dock.raise_()
                self._show_status_message("Email preview generated in dock.", 3000)

                # Add a copy button to the dock if it exists and doesn't already have one
                # Check if the dock's widget is the container we expect or just the text edit
                dock_widget = self.main_window.email_preview_dock.widget()
                if not isinstance(dock_widget, QWidget) or not dock_widget.findChild(QPushButton, "emailCopyButton"):
                     try:
#                         from PyQt5.QtWidgets import QPushButton, QVBoxLayout

                         # Create a container widget and layout
                         container = QWidget()
                         layout = QVBoxLayout(container)
                         layout.setContentsMargins(0,0,0,0) # Remove margins from container layout
                         layout.setSpacing(5) # Add some spacing

                         # If the current widget is the QTextEdit, move it to the container
                         if dock_widget == self.main_window.email_preview_view:
                              # Need to remove it from the dock first if it was set directly
                              self.main_window.email_preview_dock.setWidget(None)
                              layout.addWidget(self.main_window.email_preview_view)
                         elif dock_widget is not None:
                              # Assume it's already a container, just add the button
                              # This might need adjustment if the existing layout is complex
                              existing_layout = dock_widget.layout()
                              if existing_layout:
                                   # We'll add the button below the existing content
                                   pass # Button added after this block
                              else: # No layout, just add the text view if it wasn't the direct widget
                                   if self.main_window.email_preview_view.parent() != container:
                                        layout.addWidget(self.main_window.email_preview_view)


                         # Create and add the button
                         copy_button = QPushButton("Copy HTML to Clipboard")
                         copy_button.setObjectName("emailCopyButton") # Set object name for future checks
                         copy_button.clicked.connect(lambda: self._copy_email_to_clipboard(html_body))
                         layout.addWidget(copy_button)

                         # Set the container as the new widget for the dock
                         self.main_window.email_preview_dock.setWidget(container)
                         print("Added/Updated copy button in email preview dock")

                     except Exception as button_err:
                         print(f"Could not add copy button to dock: {button_err}")


            except Exception as e:
                print(f"Error showing email preview dock: {e}")
                QMessageBox.warning(self, "Preview Error", f"Could not display email preview:\n{e}")
        else:
            print("Warning: Cannot show email preview dock. Main window reference missing or dock not set up.")
            QMessageBox.warning(self, "Preview Error", "Could not find the email preview window.")

        # --- Action 2: Open Mailto Link (with minimal body) ---
        mailto_body = "Please copy and paste the formatted content from the Email Preview window here.\n\n"
        mailto_url = f"mailto:{recipients_string}?subject={quote(subject)}&body={quote(mailto_body)}"

        try:
            if len(mailto_url) > 1800:  # Be conservative with mailto length limits
                print(f"Warning: Mailto URL is very long ({len(mailto_url)} chars), may exceed client limits.")
                QMessageBox.warning(self, "Email Link Warning",
                                  "The generated email link is very long and might not open correctly or be truncated in your email client.")

            success = webbrowser.open_new_tab(mailto_url)
            if success:
                print(f"Attempted to open mail client for draft.")
            else:
                print("webbrowser.open_new_tab returned False. Mail client might not be configured or URL too long/invalid.")
                QMessageBox.warning(self, "Email Client Error",
                                  "Could not automatically open email client.\nIs a default email application set up?")
        except Exception as e:
            print(f"Error opening mailto link: {e}")
            QMessageBox.critical(self, "Email Client Error", f"Could not open email client:\n{e}")

    def _copy_email_to_clipboard(self, html_content):
        """Helper method to copy the HTML email content to clipboard."""
        try:
            clipboard = QApplication.clipboard()
            # Set HTML content on the clipboard
            mime_data = clipboard.mimeData(mode=QClipboard.Clipboard)
            if mime_data is None:
                 mime_data = QMimeData() # Create if None

            mime_data.setHtml(html_content)
            # Provide plain text fallback
            # Basic text conversion (replace some tags) - could be improved
            plain_text = re.sub('<br\s*/?>', '\n', html_content)
            plain_text = re.sub('<p>', '\n', plain_text, flags=re.IGNORECASE) # Handle <p> tags
            plain_text = re.sub('</p>', '\n', plain_text, flags=re.IGNORECASE)
            plain_text = re.sub('<li>', '\n - ', plain_text, flags=re.IGNORECASE) # Handle lists
            plain_text = re.sub('<[^>]+>', '', plain_text) # Strip remaining tags
            plain_text = html.unescape(plain_text) # Decode HTML entities
            mime_data.setText(plain_text.strip())

            clipboard.setMimeData(mime_data, mode=QClipboard.Clipboard)

            self._show_status_message("Email HTML copied to clipboard!", 3000)
        except Exception as e:
            print(f"Error copying to clipboard: {e}")
            # Fallback to plain text copy if HTML fails
            try:
                 clipboard = QApplication.clipboard()
                 clipboard.setText(html_content) # Copy raw HTML as fallback
                 self._show_status_message("Copied raw HTML to clipboard (MIME data failed).", 3000)
            except Exception as e2:
                 print(f"Fallback text copy also failed: {e2}")
                 self._show_status_message("Failed to copy to clipboard", 3000)


    def generate_csv_and_email(self):
        """Generates CSV, saves to SharePoint, and then generates email."""
        if self.generate_csv():
            self.generate_email()
        else:
            print("CSV generation/saving failed, email generation skipped.")


    def apply_styles(self):
        """Applies stylesheets to the form widgets."""
        # Fix the SyntaxWarning for invalid escape sequence \s by using raw string r'...'
        self.setStyleSheet(r""" AMSDealForm { background-color: #f1f3f5; } QGroupBox { font-weight: bold; font-size: 16px; color: #367C2B; background-color: #ffffff; border: 1px solid #d1d5db; border-radius: 8px; padding: 25px 10px 10px 10px; margin-top: 10px; } QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0 5px; left: 10px; color: #111827; } QLineEdit, QComboBox, QSpinBox { border: 1px solid #d1d5db; border-radius: 6px; padding: 8px; font-size: 14px; background-color: #ffffff; color: #1f2937; } QLineEdit:focus, QComboBox:focus, QSpinBox:focus { border: 2px solid #367C2B; padding: 7px; } QLineEdit[readOnly="true"] { background-color: #e9ecef; color: #6b7280; } QPushButton { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #FFDE00, stop:1 #e6c700); color: #000000; border: 1px solid #dca100; padding: 8px 15px; border-radius: 6px; font-size: 14px; font-weight: bold; } QPushButton:hover { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffe633, stop:1 #FFDE00); border: 1px solid #e6c700; } QPushButton:pressed { background: #e6c700; } QPushButton#reset_btn { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #f87171, stop:1 #dc2626); color: white; border: 1px solid #b91c1c; } QPushButton#reset_btn:hover { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ef4444, stop:1 #dc2626); border: 1px solid #991b1b; } QPushButton#reset_btn:pressed { background: #b91c1c; } QPushButton#delete_line_btn, QPushButton#save_draft_btn, QPushButton#load_draft_btn { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #d1d5db, stop:1 #9ca3af); color: #1f2937; border: 1px solid #6b7280; } QPushButton#delete_line_btn:hover, QPushButton#save_draft_btn:hover, QPushButton#load_draft_btn:hover { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #e5e7eb, stop:1 #d1d5db); border: 1px solid #4b5563; } QPushButton#delete_line_btn:pressed, QPushButton#save_draft_btn:pressed, QPushButton#load_draft_btn:pressed { background: #9ca3af; } QListWidget { background-color: #ffffff; border: 1px solid #d1d5db; border-radius: 6px; padding: 5px; font-size: 13px; color: #374151; } QCompleter::popup { border: 1px solid #9ca3af; border-radius: 6px; background-color: #ffffff; font-size: 13px; padding: 2px; } QCompleter::popup::item:selected { background-color: #e0f2fe; color: #0c4a6e; } QCheckBox { font-size: 14px; color: #333; spacing: 5px; } QCheckBox::indicator { width: 16px; height: 16px; } QCheckBox::indicator:checked { background-color: #367C2B; border: 1px solid #2a5d24; border-radius: 3px; } QCheckBox::indicator:unchecked { background-color: white; border: 1px solid #9ca3af; border-radius: 3px; } QLabel#errorLabel { color: #dc2626; font-size: 12px; } """)
        # Ensure object names are set for specific styling
        self.reset_btn.setObjectName("reset_btn");
        self.delete_line_btn.setObjectName("delete_line_btn");
        self.save_draft_btn.setObjectName("save_draft_btn");
        self.load_draft_btn.setObjectName("load_draft_btn")
        # Re-polish buttons to apply object name specific styles if needed
        for btn in [self.reset_btn, self.delete_line_btn, self.save_draft_btn, self.load_draft_btn, self.generate_csv_btn, self.generate_email_btn, self.generate_both_btn]:
             self.style().unpolish(btn); self.style().polish(btn)

    def get_csv_lines(self):
        """Returns the generated CSV lines (if any)."""
        if not self.csv_lines: print("Warning: get_csv_lines called but no CSV lines were generated.")
        return self.csv_lines

# --- CSV Output Dialog ---
# This is just a placeholder now as CSV is saved directly
class CSVOutputDialog(QMessageBox):
     def __init__(self, csv_content, parent=None, form=None):
        super().__init__(parent); self.setWindowTitle("Generated CSV Preview"); self.setIcon(QMessageBox.Information)
        self.setText("CSV Content Generated (See console log or saved file)."); self.setStandardButtons(QMessageBox.Ok); print("--- CSV DIALOG (Placeholder - Full content logged/saved) ---")

# --- Main execution (for testing AMSDealForm independently) ---
if __name__ == '__main__':
    # Attempt to import QWebEngineView for the main app context, but handle if missing
    try: from PyQt5.QtWebEngineWidgets import QWebEngineView
    except ImportError:
        print("WARNING: PyQtWebEngine not installed. Email preview might not work if main app relies on it.")
        QWebEngineView = None # Set to None so checks later don't cause NameError

    # Load environment variables if dotenv is installed
    try:
        from dotenv import load_dotenv
        script_dir = os.path.dirname(__file__) if '__file__' in locals() else os.getcwd()
        dotenv_path_script = os.path.join(script_dir, '.env')
        dotenv_path_parent = os.path.join(os.path.abspath(os.path.join(script_dir, os.pardir)), '.env')
        dotenv_path_to_use = None
        if os.path.exists(dotenv_path_script): dotenv_path_to_use = dotenv_path_script
        elif os.path.exists(dotenv_path_parent): dotenv_path_to_use = dotenv_path_parent
        if dotenv_path_to_use: print(f"Loading .env from: {dotenv_path_to_use}"); load_dotenv(dotenv_path=dotenv_path_to_use, verbose=True)
        else: print(f"Warning: .env file not found in script or parent directory.")
    except ImportError: print("Warning: python-dotenv not installed. Cannot load .env file.")

    app = QApplication(sys.argv)
    if hasattr(Qt, 'AA_EnableHighDpiScaling'): QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'): QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app.setStyle("Fusion")

    # Create a Mock Main Window that includes the necessary dock widget for testing
    class MockMainWindowForTest(QMainWindow):
        def __init__(self):
            super().__init__()
            self.setWindowTitle("Mock Main Window for Test")
            self.statusBar = self.statusBar() # Create status bar instance
            self.statusBar.showMessage("Mock status bar ready.")

            # Set a dummy data_path for testing the _get_data_path logic if needed
            # self.data_path = "path/to/test/data" # Uncomment and set if needed

            # Add dummy preview dock for testing email generation
            self.email_preview_dock = QDockWidget("Email Preview", self)
            self.email_preview_view = QTextEdit() # Use QTextEdit for basic HTML rendering in test
            self.email_preview_view.setReadOnly(True)
            # We'll add the copy button dynamically in generate_email if needed
            self.email_preview_dock.setWidget(self.email_preview_view)
            self.addDockWidget(Qt.BottomDockWidgetArea, self.email_preview_dock)
            self.email_preview_dock.setVisible(False) # Start hidden

    window = MockMainWindowForTest()

    # Use the dummy SharePoint manager if the real one failed to import or isn't configured
    mock_sp_manager = None
    if SharePointExcelManager:
         try:
              # Attempt to instantiate the real one - might need config from .env
              # For testing, often better to use a mock/dummy anyway
              print("Note: Using Dummy SharePointExcelManager for standalone test.")
              # Pass dummy args if needed, or handle config loading here
              mock_sp_manager = SharePointExcelManager() # Use dummy if real one needs args/fails
         except Exception as sp_init_err:
              print(f"Error initializing real SharePointExcelManager: {sp_init_err}. Using dummy.")
              # Define or re-use dummy class definition
              if 'SharePointExcelManager' not in locals() or SharePointExcelManager is None:
                   class SharePointExcelManager:
                        def __init__(self, *args, **kwargs): print("ERROR: Using Dummy SharePointExcelManager because import/init failed.")
                        def update_excel_data(self, data): print("Dummy SP Manager: Update called"); return False
                        def send_html_email(self, r, s, b): print("Dummy SP Manager: Send Email called"); return False
              mock_sp_manager = SharePointExcelManager()

    else: # SharePointExcelManager was None from the start
         # Define or re-use dummy class definition
         if 'SharePointExcelManager' not in locals() or SharePointExcelManager is None:
              class SharePointExcelManager:
                   def __init__(self, *args, **kwargs): print("ERROR: Using Dummy SharePointExcelManager because import failed.")
                   def update_excel_data(self, data): print("Dummy SP Manager: Update called"); return False
                   def send_html_email(self, r, s, b): print("Dummy SP Manager: Send Email called"); return False
         mock_sp_manager = SharePointExcelManager()


    # Instantiate the form, passing the mock window and manager
    test_form = AMSDealForm(main_window=window, sharepoint_manager=mock_sp_manager)

    # Set the central widget of the mock main window to be the form
    window.setCentralWidget(test_form)
    window.resize(1100, 800) # Resize main window
    window.show()

    sys.exit(app.exec_())
