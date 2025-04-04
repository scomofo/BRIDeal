# --- Start of PyQt5 Imports ---
# QtWidgets: Common UI elements and layouts
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QGridLayout, QGroupBox, QLabel, QLineEdit,
                             QPushButton, QComboBox, QCheckBox, QTextEdit,
                             QMessageBox, QApplication, QCompleter, QListWidget,
                             QListWidgetItem, QSpinBox, QInputDialog, QScrollArea,
                             QSizePolicy) # Added QSizePolicy

# QtGui: For handling images, icons, fonts, colors etc.
from PyQt5 import QtGui
from PyQt5.QtGui import QDoubleValidator, QPixmap, QClipboard # Added QPixmap, QClipboard based on usage

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
from config import AzureConfig, APIConfig, SharePointConfig


# Standard library imports used later
import csv
from datetime import datetime # Use datetime for timestamp
from urllib.parse import quote
import webbrowser # Re-added for mailto
import requests # Needed for Graph API calls
import re # Import regex for parsing edited items
import html # <-- Ensure this import is present

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
        def __init__(self): print("ERROR: Using Dummy SharePointExcelManager because import failed.")
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
        self.customers_dict = self.load_customers() # <--- This calls the modified method
        self.salesmen_emails = self.load_salesmen_emails()
        print("DEBUG: Data loading complete for AMSDealForm.")

        self.setup_ui()
        self.apply_styles()
        self.last_charge_to = ""
        print("DEBUG: AMSDealForm initialization complete.")


    # --- Helper to show status messages ---
    def _show_status_message(self, message, timeout=3000):
        """Helper to show messages on main window status bar or print."""
        if self.main_window and hasattr(self.main_window, 'statusBar'):
            self.main_window.statusBar().showMessage(message, timeout)
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
        print("Warning: No main_window provided to AMSDealForm. Using mock for status bar.")
        return MockMainWindow()

    def _get_data_path(self):
        """Helper function to determine the path to the 'data' directory."""
        current_dir = os.path.dirname(__file__); data_path = None; project_root_guess = os.path.abspath(os.path.join(current_dir, os.pardir))
        data_path_parent = os.path.join(project_root_guess, 'data'); data_path_relative = os.path.join(current_dir, 'data'); data_path_assets = os.path.join(project_root_guess, 'assets')
        if os.path.isdir(data_path_parent): data_path = data_path_parent
        elif os.path.isdir(data_path_relative): data_path = data_path_relative
        elif os.path.isdir(data_path_assets): print(f"Warning: 'data' dir not found. Using 'assets': {data_path_assets}"); data_path = data_path_assets
        if data_path is None: print(f"CRITICAL WARNING: Could not locate 'data' or 'assets' directory near {current_dir}."); return None
        else: print(f"DEBUG: Data path set to: {data_path}")
        return data_path

    # --- Data Loading Methods (Using _get_data_path) ---
    def load_customers(self):
        """Loads customer data from customers.csv in the data directory."""
        customers_dict = {}; data_dir = self._get_data_path();
        if not data_dir: print("ERROR: Cannot load customers - data dir not found."); return {}
        file_path = os.path.join(data_dir, "customers.csv"); print(f"DEBUG Customer Load: Attempting to load from: {file_path}"); encodings = ['utf-8', 'latin1', 'windows-1252']; loaded = False; rows_processed = 0
        for encoding in encodings:
            print(f"DEBUG Customer Load: Trying encoding '{encoding}'...")
            try:
                if os.path.exists(file_path):
                    with open(file_path, mode='r', newline='', encoding=encoding) as infile:
                        reader = csv.DictReader(infile); print(f"DEBUG Customer Load: Detected headers: {reader.fieldnames}")
                        if not reader.fieldnames or ("Name" not in reader.fieldnames and "name" not in reader.fieldnames):
                            print(f"Warning: 'Name' column not found in header using {encoding}. Trying index 0 fallback."); infile.seek(0); plain_reader = csv.reader(infile); headers = next(plain_reader, None); print(f"DEBUG Customer Load: Fallback headers skipped: {headers}")
                            for i, row in enumerate(plain_reader):
                                rows_processed += 1
                                if row: customer_name = row[0].strip();
                                if customer_name: customers_dict[customer_name] = {'Name': customer_name}
                        else:
                            name_key = "Name" if "Name" in reader.fieldnames else "name"; print(f"DEBUG Customer Load: Using '{name_key}' as key.")
                            for i, row in enumerate(reader):
                                rows_processed += 1; customer_name = row.get(name_key, "").strip()
                                if customer_name: customers_dict[customer_name] = row
                                else: print(f"DEBUG Customer Load: Skipping row {i+1} due to empty name.")
                    print(f"DEBUG Customer Load: Successfully processed {rows_processed} rows with encoding {encoding}.")
                    loaded = True; break
                else: print(f"Warning: Customer file not found at {file_path}"); QMessageBox.warning(self, "File Not Found", f"Required file not found:\n{file_path}"); return {}
            except UnicodeDecodeError: print(f"DEBUG Customer Load: Failed decoding with {encoding}."); rows_processed = 0; customers_dict = {}; continue
            except Exception as e: print(f"Error loading customers.csv with {encoding}: {e}"); QMessageBox.warning(self, "File Load Error", f"Could not load customer data from {file_path}:\n{e}"); return {}
        if not loaded: print(f"Failed to decode or load {file_path} with any supported encoding.")
        print(f"DEBUG Customer Load: Final dictionary size: {len(customers_dict)}")
        return customers_dict

    def load_csv(self, filename): # Generic loader
        names = []; data_dir = self._get_data_path();
        if not data_dir: return []
        file_path = os.path.join(data_dir, filename); encodings = ['utf-8', 'latin1', 'windows-1252']; loaded = False
        for encoding in encodings:
            try:
                if os.path.exists(file_path):
                    with open(file_path, newline='', encoding=encoding) as csvfile: reader = csv.reader(csvfile); names = [row[0].strip() for row in reader if row]
                    loaded = True; break
            except UnicodeDecodeError: continue
            except Exception as e: print(f"Error loading {filename}: {e}"); return []
        if not loaded and not os.path.exists(file_path): print(f"Warning: File not found: {file_path}!")
        elif not loaded: print(f"Failed to decode {filename}.")
        return names

    def load_salesmen_emails(self):
        salesmen_emails = {}; data_dir = self._get_data_path();
        if not data_dir: return {}
        file_path = os.path.join(data_dir, "salesmen.csv"); encodings = ['utf-8', 'latin1', 'windows-1252']; loaded = False
        for encoding in encodings:
            try:
                if os.path.exists(file_path):
                    with open(file_path, newline='', encoding=encoding) as csvfile:
                        reader = csv.DictReader(csvfile); headers = [h.lower() for h in reader.fieldnames or []]
                        if "name" not in headers or "email" not in headers: print(f"Error: salesmen.csv missing Name/Email column"); QMessageBox.critical(self, "File Format Error", f"Salesmen file missing Name/Email"); return {}
                        name_key = next((k for k in reader.fieldnames if k.lower() == 'name'), None); email_key = next((k for k in reader.fieldnames if k.lower() == 'email'), None)
                        for row in reader: name = row.get(name_key, "").strip(); email = row.get(email_key, "").strip();
                        if name and email: salesmen_emails[name] = email
                    loaded = True; break
            except UnicodeDecodeError: continue
            except Exception as e: print(f"Error loading salesmen.csv: {e}"); QMessageBox.warning(self, "File Load Error", f"Could not load salesmen data:\n{e}"); return {}
        if not loaded and not os.path.exists(file_path): print("Warning: salesmen.csv not found!"); QMessageBox.warning(self, "File Not Found", f"File not found:\nsalesmen.csv")
        elif not loaded: print("Failed to decode salesmen.csv.")
        return salesmen_emails

    def load_products(self, filename='products.csv'):
        products_dict = {}; data_dir = self._get_data_path();
        if not data_dir: return {}
        file_path = os.path.join(data_dir, filename); encodings = ['windows-1252', 'utf-8', 'latin1']; loaded = False
        for encoding in encodings:
            try:
                if os.path.exists(file_path):
                    with open(file_path, newline='', encoding=encoding) as csvfile:
                        reader = csv.DictReader(csvfile); required_headers_lower = {"productcode", "productname", "price"}; actual_headers_lower = {h.lower().strip() for h in reader.fieldnames or []}
                        if required_headers_lower.issubset(actual_headers_lower):
                            code_key = next((k for k in reader.fieldnames if k.lower().strip() == 'productcode'), None); name_key = next((k for k in reader.fieldnames if k.lower().strip() == 'productname'), None); price_key = next((k for k in reader.fieldnames if k.lower().strip() == 'price'), None)
                            for row in reader: name = row.get(name_key,"").strip(); code = row.get(code_key,"").strip(); price = row.get(price_key,"").strip() or "0.00";
                            if name and code: products_dict[name] = (code, price)
                            loaded = True; break
                        else: print(f"Warning: Missing columns in {filename} with {encoding}.")
            except UnicodeDecodeError: continue
            except Exception as e: print(f"Error loading {filename}: {e}"); QMessageBox.warning(self, "File Load Error", f"Could not load product data:\n{e}"); return {}
        if not loaded and not os.path.exists(file_path): print(f"Warning: {filename} not found!"); QMessageBox.warning(self, "File Not Found", f"File not found:\n{filename}")
        elif not loaded: print(f"Failed to decode {filename} or headers missing."); QMessageBox.warning(self, "File Format Error", f"Could not decode {filename} or required headers missing.")
        return products_dict

    def load_parts(self):
        parts_dict = {}; data_dir = self._get_data_path();
        if not data_dir: return {}
        file_path = os.path.join(data_dir, "parts.csv"); encodings = ['utf-8', 'latin1', 'windows-1252']; loaded = False
        for encoding in encodings:
            try:
                if os.path.exists(file_path):
                    with open(file_path, newline='', encoding=encoding) as csvfile:
                        reader = csv.DictReader(csvfile); required_headers_lower = {"part number", "part name"}; actual_headers_lower = {h.lower().strip() for h in reader.fieldnames or []}
                        if required_headers_lower.issubset(actual_headers_lower):
                            number_key = next((k for k in reader.fieldnames if k.lower().strip() == 'part number'), None); name_key = next((k for k in reader.fieldnames if k.lower().strip() == 'part name'), None)
                            for row in reader: name = row.get(name_key, "").strip(); number = row.get(number_key, "").strip();
                            if name and number: parts_dict[name] = number
                            loaded = True; break
                        else: # Index fallback
                            print(f"Warning: Missing columns in parts.csv with {encoding}. Trying index fallback."); csvfile.seek(0); plain_reader = csv.reader(csvfile); headers = next(plain_reader, None); temp_dict = {}
                            for row in plain_reader:
                                if row and len(row) >= 2: number = row[0].strip(); name = row[1].strip();
                                if name and number: temp_dict[name] = number
                            if temp_dict: parts_dict = temp_dict; loaded = True; break
            except UnicodeDecodeError: continue
            except Exception as e: print(f"Error loading parts.csv: {e}"); QMessageBox.warning(self, "File Load Error", f"Could not load parts data:\n{e}"); return {}
        if not loaded and not os.path.exists(file_path): print(f"Warning: {file_path} not found!"); QMessageBox.warning(self, "File Not Found", f"File not found:\n{file_path}")
        elif not loaded: print("Failed to decode parts.csv or find columns.")
        return parts_dict
    # --- End of Data Loading Methods ---


    # --- setup_ui METHOD with QScrollArea ---
    def setup_ui(self):
        # ... (UI setup remains unchanged from previous version, includes draft/delete buttons) ...
        main_layout_for_self = QVBoxLayout(self); main_layout_for_self.setContentsMargins(0,0,0,0)
        scroll_area = QScrollArea(self); scroll_area.setWidgetResizable(True); scroll_area.setStyleSheet("QScrollArea { border: none; background-color: #f1f3f5; }")
        content_container_widget = QWidget(); content_layout = QVBoxLayout(content_container_widget); content_layout.setSpacing(15); content_layout.setContentsMargins(15, 15, 15, 15)
        header = QWidget(); header.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #367C2B, stop:1 #4a9c3e); padding: 15px; border-bottom: 2px solid #2a5d24;"); header_layout = QHBoxLayout(header)
        logo_label = QLabel(self); logo_path_local = os.path.join(os.path.dirname(__file__), "logo.png"); logo_pixmap = QtGui.QPixmap(logo_path_local)
        if not logo_pixmap.isNull(): logo_label.setPixmap(logo_pixmap.scaled(80, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else: logo_label.setText("Logo Missing"); logo_label.setStyleSheet("color: white;")
        title_label = QLabel("AMS Deal Form"); title_label.setStyleSheet("color: white; font-size: 40px; font-weight: bold; font-family: Arial;"); header_layout.addWidget(logo_label); header_layout.addWidget(title_label); header_layout.addStretch(); content_layout.addWidget(header)
        customer_group = QGroupBox("Customer & Salesperson"); customer_layout = QHBoxLayout(customer_group); self.customer_name = QLineEdit(); self.customer_name.setPlaceholderText("Customer Name"); customer_completer = QCompleter(list(self.customers_dict.keys()), self); customer_completer.setCaseSensitivity(Qt.CaseInsensitive); customer_completer.setFilterMode(Qt.MatchContains); self.customer_name.setCompleter(customer_completer); self.salesperson = QLineEdit(); self.salesperson.setPlaceholderText("Salesperson"); salesperson_completer = QCompleter(list(self.salesmen_emails.keys()), self); salesperson_completer.setCaseSensitivity(Qt.CaseInsensitive); salesperson_completer.setFilterMode(Qt.MatchContains); self.salesperson.setCompleter(salesperson_completer); customer_layout.addWidget(self.customer_name); customer_layout.addWidget(self.salesperson); content_layout.addWidget(customer_group)
        equipment_group = QGroupBox("Equipment"); equipment_layout = QVBoxLayout(equipment_group); equipment_input_layout = QHBoxLayout(); self.equipment_product_name = QLineEdit(); self.equipment_product_name.setPlaceholderText("Product Name"); equipment_completer = QCompleter(list(self.products_dict.keys()), self); equipment_completer.setCaseSensitivity(Qt.CaseInsensitive); equipment_completer.setFilterMode(Qt.MatchContains); self.equipment_product_name.setCompleter(equipment_completer); equipment_completer.activated.connect(self.on_equipment_selected); self.equipment_product_code = QLineEdit(); self.equipment_product_code.setPlaceholderText("Product Code"); self.equipment_product_code.setReadOnly(True); self.equipment_product_code.setToolTip("Product Code from products.csv (auto-filled)"); self.equipment_manual_stock = QLineEdit(); self.equipment_manual_stock.setPlaceholderText("Stock # (Manual)"); self.equipment_manual_stock.setToolTip("Enter the specific Stock Number for this deal"); self.equipment_price = QLineEdit(); self.equipment_price.setPlaceholderText("$0.00"); validator_eq = QDoubleValidator(0.0, 9999999.99, 2); validator_eq.setNotation(QDoubleValidator.StandardNotation); self.equipment_price.setValidator(validator_eq); self.equipment_price.editingFinished.connect(self.format_price); equipment_add_btn = QPushButton("Add"); equipment_add_btn.clicked.connect(lambda: self.add_item("equipment")); equipment_input_layout.addWidget(self.equipment_product_name, 3); equipment_input_layout.addWidget(self.equipment_product_code, 1); equipment_input_layout.addWidget(self.equipment_manual_stock, 1); equipment_input_layout.addWidget(self.equipment_price, 1); equipment_input_layout.addWidget(equipment_add_btn, 0); self.equipment_list = QListWidget(); self.equipment_list.setSelectionMode(QListWidget.SingleSelection); self.equipment_list.itemDoubleClicked.connect(self.edit_equipment_item); equipment_layout.addLayout(equipment_input_layout); equipment_layout.addWidget(self.equipment_list); content_layout.addWidget(equipment_group)
        trades_group = QGroupBox("Trades"); trades_layout = QVBoxLayout(trades_group); trades_input_layout = QHBoxLayout(); self.trade_name = QLineEdit(); self.trade_name.setPlaceholderText("Trade Item"); trade_completer = QCompleter(list(self.products_dict.keys()), self); trade_completer.setCaseSensitivity(Qt.CaseInsensitive); trade_completer.setFilterMode(Qt.MatchContains); self.trade_name.setCompleter(trade_completer); trade_completer.activated.connect(self.on_trade_selected); self.trade_stock = QLineEdit(); self.trade_stock.setPlaceholderText("Stock #"); self.trade_amount = QLineEdit(); self.trade_amount.setPlaceholderText("$0.00"); validator_tr = QDoubleValidator(0.0, 9999999.99, 2); validator_tr.setNotation(QDoubleValidator.StandardNotation); self.trade_amount.setValidator(validator_tr); self.trade_amount.editingFinished.connect(self.format_amount); trades_add_btn = QPushButton("Add"); trades_add_btn.clicked.connect(lambda: self.add_item("trade")); trades_input_layout.addWidget(self.trade_name); trades_input_layout.addWidget(self.trade_stock); trades_input_layout.addWidget(self.trade_amount); trades_input_layout.addWidget(trades_add_btn); self.trade_list = QListWidget(); self.trade_list.setSelectionMode(QListWidget.SingleSelection); self.trade_list.itemDoubleClicked.connect(self.edit_trade_item); trades_layout.addLayout(trades_input_layout); trades_layout.addWidget(self.trade_list); content_layout.addWidget(trades_group)
        parts_group = QGroupBox("Parts"); parts_layout = QVBoxLayout(parts_group); parts_input_layout = QHBoxLayout(); self.part_quantity = QSpinBox(); self.part_quantity.setValue(1); self.part_quantity.setMinimum(1); self.part_quantity.setMaximum(999); self.part_quantity.setFixedWidth(60); self.part_number = QLineEdit(); self.part_number.setPlaceholderText("Part #"); part_number_completer = QCompleter(list(self.parts_dict.values()), self); part_number_completer.setCaseSensitivity(Qt.CaseInsensitive); part_number_completer.setFilterMode(Qt.MatchContains); self.part_number.setCompleter(part_number_completer); part_number_completer.activated.connect(self.on_part_number_selected); self.part_name = QLineEdit(); self.part_name.setPlaceholderText("Part Name"); part_name_completer = QCompleter(list(self.parts_dict.keys()), self); part_name_completer.setCaseSensitivity(Qt.CaseInsensitive); part_name_completer.setFilterMode(Qt.MatchContains); self.part_name.setCompleter(part_name_completer); part_name_completer.activated.connect(self.on_part_selected); self.part_location = QComboBox(); self.part_location.addItems(["", "Camrose", "Killam", "Wainwright", "Provost"]); self.part_charge_to = QLineEdit(); self.part_charge_to.setPlaceholderText("Charge to:"); parts_add_btn = QPushButton("Add"); parts_add_btn.clicked.connect(lambda: self.add_item("part")); parts_input_layout.addWidget(self.part_quantity); parts_input_layout.addWidget(self.part_number); parts_input_layout.addWidget(self.part_name); parts_input_layout.addWidget(self.part_location); parts_input_layout.addWidget(self.part_charge_to); parts_input_layout.addWidget(parts_add_btn); self.part_list = QListWidget(); self.part_list.setSelectionMode(QListWidget.SingleSelection); self.part_list.itemDoubleClicked.connect(self.edit_part_item); parts_layout.addLayout(parts_input_layout); parts_layout.addWidget(self.part_list); content_layout.addWidget(parts_group)
        work_order_group = QGroupBox("Work Order & Options"); work_order_layout = QHBoxLayout(work_order_group); self.work_order_required = QCheckBox("Work Order Req'd?"); self.work_order_charge_to = QLineEdit(); self.work_order_charge_to.setPlaceholderText("Charge to:"); self.work_order_charge_to.setFixedWidth(150); self.work_order_hours = QLineEdit(); self.work_order_hours.setPlaceholderText("Duration (hours)"); self.multi_line_csv = QCheckBox("Multiple CSV Lines"); self.paid_checkbox = QCheckBox("Paid"); self.paid_checkbox.setStyleSheet("font-size: 16px; color: #333;"); self.paid_checkbox.setChecked(False); work_order_layout.addWidget(self.work_order_required); work_order_layout.addWidget(self.work_order_charge_to); work_order_layout.addWidget(self.work_order_hours); work_order_layout.addWidget(self.multi_line_csv); work_order_layout.addStretch(); work_order_layout.addWidget(self.paid_checkbox); content_layout.addWidget(work_order_group)
        buttons_layout = QHBoxLayout(); self.delete_line_btn = QPushButton("Delete Selected Line"); self.delete_line_btn.setToolTip("Delete the selected line from the list above (Equipment, Trade, or Part)"); self.delete_line_btn.clicked.connect(self.delete_selected_list_item); buttons_layout.addWidget(self.delete_line_btn); buttons_layout.addStretch(1); self.save_draft_btn = QPushButton("Save Draft"); self.save_draft_btn.setToolTip("Save the current form entries as a draft"); self.save_draft_btn.clicked.connect(self.save_draft); buttons_layout.addWidget(self.save_draft_btn); self.load_draft_btn = QPushButton("Load Draft"); self.load_draft_btn.setToolTip("Load the last saved draft"); self.load_draft_btn.clicked.connect(self.load_draft); buttons_layout.addWidget(self.load_draft_btn); buttons_layout.addSpacing(20); self.generate_csv_btn = QPushButton("Gen. CSV & Save"); self.generate_csv_btn.clicked.connect(self.generate_csv); self.generate_email_btn = QPushButton("Gen. Email"); self.generate_email_btn.clicked.connect(self.generate_email); self.generate_both_btn = QPushButton("Generate All"); self.generate_both_btn.clicked.connect(self.generate_csv_and_email); self.reset_btn = QPushButton("Reset Form"); self.reset_btn.clicked.connect(self.reset_form); self.reset_btn.setObjectName("reset_btn"); buttons_layout.addWidget(self.generate_csv_btn); buttons_layout.addSpacing(10); buttons_layout.addWidget(self.generate_email_btn); buttons_layout.addSpacing(10); buttons_layout.addWidget(self.generate_both_btn); buttons_layout.addSpacing(10); buttons_layout.addWidget(self.reset_btn); content_layout.addLayout(buttons_layout)
        content_layout.addStretch(); scroll_area.setWidget(content_container_widget); main_layout_for_self.addWidget(scroll_area)

    # --- Helper methods ---
    # ... (update_charge_to_default, format_price, format_amount unchanged) ...
    def update_charge_to_default(self):
        if self.equipment_list.count() > 0: line = self.equipment_list.item(0).text(); match = re.search(r'STK#(\S+)', line)
        if match: stock_number = match.group(1);
        if stock_number:
            if not self.part_charge_to.text(): self.part_charge_to.setText(stock_number)
            self.last_charge_to = stock_number
        else: print(f"Error parsing equipment line (regex failed): {line}")

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
        new_code, new_default_price_str = self.products_dict.get(new_name, (code, price_str))
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
        except ValueError: QMessageBox.warning(self, "Invalid Amount", f"Invalid amount: '{new_amount_str}'."); new_amount_formatted = "0.00"
        item.setText(f'"{new_name}" STK#{new_stock} ${new_amount_formatted}')

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
        if text in self.products_dict:
            code, price_str = self.products_dict[text]
            current_price = self.equipment_price.text().replace('$','').replace(',','')
            self.equipment_product_code.setText(code)
            try:
                 if not current_price or float(current_price) == 0.0:
                     price_value = float(price_str) if price_str else 0.0
                     self.equipment_price.setText(f"${price_value:,.2f}")
            except ValueError:
                 print(f"Invalid price format '{price_str}' for product '{text}'")
                 self.equipment_price.setText("$0.00")
            self.equipment_manual_stock.setFocus()
        else:
            self.equipment_product_code.clear()


    def on_trade_selected(self, text):
        """Handles selection from trade completer, fills amount/stock if empty."""
        print(f"Trade selected: '{text}'")
        if text in self.products_dict:
            code, price_str = self.products_dict[text]
            current_amount = self.trade_amount.text().replace('$','').replace(',','')
            current_stock = self.trade_stock.text()
            try:
                if not current_amount or float(current_amount) == 0.0:
                    price_value = float(price_str) if price_str else 0.0
                    self.trade_amount.setText(f"${price_value:,.2f}")
            except ValueError:
                print(f"Invalid price format '{price_str}' for product '{text}'")
                self.trade_amount.setText("$0.00")
            if not current_stock and code:
                self.trade_stock.setText(code)

    def on_part_selected(self, text):
        """Handles selection from part name completer, fills part number."""
        text = text.strip(); print(f"Part name selected: '{text}'")
        part_number = self.parts_dict.get(text)
        if part_number: self.part_number.setText(part_number)
        else: print(f"Part name '{text}' not found in parts_dict.")

    def on_part_number_selected(self, text):
        """Handles selection from part number completer, fills part name."""
        text = text.strip(); print(f"Part number selected: '{text}'")
        found = False
        for part_name, part_num in self.parts_dict.items():
            if part_num == text: self.part_name.setText(part_name); found = True; break
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

            # --- Refresh Completers (with individual try/except) ---
            print("DEBUG: Refreshing completer models after reset...")
            # Customer
            try:
                customer_keys = list(self.customers_dict.keys())
                print(f"DEBUG Reset: Customer completer list size: {len(customer_keys)}")
                if hasattr(self.customer_name, 'completer') and self.customer_name.completer():
                    self.customer_name.completer().model().setStringList(customer_keys)
                else: print("DEBUG Reset: Customer completer not found.")
            except Exception as e: print(f"ERROR: Could not refresh Customer completer model: {e}")
            # Salesperson
            try:
                salesmen_keys = list(self.salesmen_emails.keys())
                print(f"DEBUG Reset: Salesperson completer list size: {len(salesmen_keys)}")
                if hasattr(self.salesperson, 'completer') and self.salesperson.completer():
                    self.salesperson.completer().model().setStringList(salesmen_keys)
                else: print("DEBUG Reset: Salesperson completer not found.")
            except Exception as e: print(f"ERROR: Could not refresh Salesperson completer model: {e}")
            # Equipment / Trade (Product Names)
            try:
                product_keys = list(self.products_dict.keys())
                print(f"DEBUG Reset: Equipment/Trade completer list size: {len(product_keys)}")
                if hasattr(self.equipment_product_name, 'completer') and self.equipment_product_name.completer():
                    self.equipment_product_name.completer().model().setStringList(product_keys)
                else: print("DEBUG Reset: Equipment completer not found.")
                if hasattr(self.trade_name, 'completer') and self.trade_name.completer():
                    self.trade_name.completer().model().setStringList(product_keys)
                else: print("DEBUG Reset: Trade completer not found.")
            except Exception as e: print(f"ERROR: Could not refresh Equipment/Trade completer models: {e}")
            # Part Name
            try:
                part_name_keys = list(self.parts_dict.keys())
                print(f"DEBUG Reset: Part Name completer list size: {len(part_name_keys)}")
                if hasattr(self.part_name, 'completer') and self.part_name.completer():
                    self.part_name.completer().model().setStringList(part_name_keys)
                else: print("DEBUG Reset: Part Name completer not found.")
            except Exception as e: print(f"ERROR: Could not refresh Part Name completer model: {e}")
            # Part Number
            try:
                part_number_values = list(self.parts_dict.values())
                print(f"DEBUG Reset: Part Number completer list size: {len(part_number_values)}")
                if hasattr(self.part_number, 'completer') and self.part_number.completer():
                    self.part_number.completer().model().setStringList(part_number_values)
                else: print("DEBUG Reset: Part Number completer not found.")
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
        if focused_widget == self.equipment_list: target_list, list_name = self.equipment_list, "Equipment"
        elif focused_widget == self.trade_list: target_list, list_name = self.trade_list, "Trade"
        elif focused_widget == self.part_list: target_list, list_name = self.part_list, "Part"
        else:
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
        if not self.sharepoint_manager: print("SharePoint Manager not initialized."); self._show_status_message("SP connection failed.", 5000); return False
        if not csv_lines: QMessageBox.warning(self, "Save Error", "No data generated."); return False
        try:
            data_to_save = []; headers = ["Payment", "Customer", "Equipment", "Stock Number", "Amount", "Trade", "Attached to stk#", "Trade STK#", "Amount2", "Salesperson", "Email Date", "Status", "Timestamp"]
            csv_data = "\n".join(csv_lines); csvfile = io.StringIO(csv_data); reader = csv.reader(csvfile, delimiter=',', quotechar='"')
            for fields in reader:
                if len(fields) == len(headers): data_to_save.append(dict(zip(headers, fields)))
                else: print(f"Warning: Skipping malformed CSV line: {fields}")
            if not data_to_save: QMessageBox.warning(self, "Save Error", "Could not parse generated data."); return False
            success = self.sharepoint_manager.update_excel_data(data_to_save)
            if success: self._show_status_message("Data saved to SharePoint!", 5000); return True
            else: self._show_status_message("Failed to save to SharePoint", 5000); return False
        except Exception as e: print(f"Error in save_to_csv: {e}"); QMessageBox.critical(self, "Save Error", f"Unexpected error during save:\n{e}"); self._show_status_message(f"SP Save Error: {e}", 5000); return False


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
                        name_part, stock_part = line.split(" STK#", 1); trade = name_part.strip('" '); trade_stock, trade_amount_str = stock_part.split(" $", 1); trade_amount = trade_amount_str.strip()
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
                    line = trade_items[0]; name_part, stock_part = line.split(" STK#", 1); trade = name_part.strip('" '); trade_stock, trade_amount_str = stock_part.split(" $", 1); trade_amount = trade_amount_str.strip();
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
        except Exception as e: print(f"Warning: Could not load recent deals file: {e}."); recent_deals = []
        recent_deals.insert(0, deal_data); recent_deals = recent_deals[:10]
        try:
            with open(self.recent_deals_file, 'w', encoding='utf-8') as f: json.dump(recent_deals, f, indent=4)
            print(f"Saved deal to recent list ({len(recent_deals)} items).")
        except Exception as e: print(f"Error saving recent deals file: {e}")


    # --- generate_email (Mailto + HTML Preview Dock Logic) ---
    # *** This is the method to replace ***
    def generate_email(self):
        """Generate HTML email body, show in preview dock, and open mailto link."""
        print("Generating email (HTML Preview + Mailto)...")
        customer_name = self.customer_name.text().strip()
        salesperson = self.salesperson.text().strip()
        equipment_items = [self.equipment_list.item(i).text() for i in range(self.equipment_list.count())]
        trade_items = [self.trade_list.item(i).text() for i in range(self.trade_list.count())]
        part_items = [self.part_list.item(i).text() for i in range(self.part_list.count())]

        if not customer_name: QMessageBox.warning(self, "Missing Info", "Please enter a Customer Name."); return
        if not salesperson: QMessageBox.warning(self, "Missing Info", "Please enter a Salesperson."); return

        # --- Get First Equipment Info ---
        first_product_name = "N/A"; first_stock_number = "N/A"
        eq_pattern = r'"(.*)"\s+\(Code:\s*(.*?)\)\s+STK#(.*?)\s+\$(.*)'
        if equipment_items:
            match = re.match(eq_pattern, equipment_items[0])
            if match: first_product_name, _, first_stock_number, _ = match.groups()
            else:
                try: first_product_name = equipment_items[0].split('"')[1]
                except IndexError: first_product_name = equipment_items[0][:30]

        # --- Subject ---
        subject = f"AMS Deal ({customer_name} - {first_product_name})" # Use user's requested format

        # --- Recipients ---
        fixed_recipients = ['bstdenis@briltd.com', 'cgoodrich@briltd.com', 'dvriend@briltd.com','rbendfeld@briltd.com', 'bfreadrich@briltd.com', 'vedwards@briltd.com']
        salesman_email = self.salesmen_emails.get(salesperson)
        if salesman_email: fixed_recipients.append(salesman_email)
        else: print(f"Warning: Email for salesperson '{salesperson}' not found."); QMessageBox.warning(self, "Salesperson Email", f"Email for {salesperson} not found.")
        if part_items: fixed_recipients.append('rkrys@briltd.com')
        recipient_list = [r for r in fixed_recipients if r and '@' in r]
        recipients_string = ";".join(recipient_list)

        # --- HTML Body Construction (for Preview Dock) ---
        table_style = 'border-collapse: collapse; width: 95%; border: 1px solid #ADADAD; margin-bottom: 15px; font-family: sans-serif; font-size: 10pt;'
        th_style = 'border: 1px solid #ADADAD; padding: 6px 8px; text-align: left; background-color: #EAEAEA; font-weight: bold;'
        td_style = 'border: 1px solid #ADADAD; padding: 6px 8px; vertical-align: top;'
        td_right_style = 'border: 1px solid #ADADAD; padding: 6px 8px; vertical-align: top; text-align: right;'

        body = []
        body.append("<html><head><style> body { font-family: sans-serif; font-size: 10pt; } </style></head><body>") # Basic styling
        # Customer / Salesperson (Matches user template)
        body.append(f"<p>Customer: {html.escape(customer_name)}<br>")
        body.append(f"Sales: {html.escape(salesperson)}</p>")

        # Equipment Table (Matches user template format)
        body.append("<p><b>Equipment</b></p>") # Header like user template
        if equipment_items:
            body.append(f'<table style="{table_style}">')
            # Headers: Name, Stock #, Price (as requested implicitly by user template)
            body.append(f"<thead><tr><th style='{th_style}'>Name</th><th style='{th_style}'>Stock #</th><th style='{th_style} text-align: right;'>Price</th></tr></thead><tbody>")
            for line in equipment_items:
                match = re.match(eq_pattern, line)
                # Row Format: {productname} {stock#} {price}
                if match: name, _code, manual_stock, price_str = match.groups(); body.append(f"<tr><td style='{td_style}'>{html.escape(name)}</td><td style='{td_style}'>{html.escape(manual_stock)}</td><td style='{td_right_style}'>${html.escape(price_str)}</td></tr>")
                else: body.append(f"<tr><td colspan='3' style='{td_style}'><i>Error parsing: {html.escape(line)}</i></td></tr>")
            body.append("</tbody></table>")
        else:
            body.append("<p>N/A</p>") # Should not happen based on earlier checks

        # Parts Table (Conditional, using user logic)
        if part_items:
            part_groups = {}
            for line in part_items:
                 parts = line.split(" ", 4)
                 if len(parts) >= 4: qty, number, name, location = parts[0].rstrip('x'), parts[1], parts[2], parts[3]; loc_key = location if location else "N/A"; part_groups.setdefault(loc_key, []).append(f"{qty} x {number} {name}")
                 else: part_groups.setdefault("Format Error", []).append(line)

            for location, parts_list in part_groups.items():
                # Header: PARTS (From {location} charge to {1st stock listed in deal})
                body.append(f"<p style='margin-top: 15px;'><b>PARTS (From {html.escape(location)} charge to {html.escape(first_stock_number)})</b></p>")
                # Simple list format within table for parts as requested
                body.append(f'<table style="{table_style}">')
                # body.append(f"<thead><tr><th style='{th_style}'>Details (Qty x Number Name)</th></tr></thead><tbody>") # Optional header
                body.append("<tbody>")
                for part_line in parts_list:
                    body.append(f"<tr><td style='{td_style}'>{html.escape(part_line)}</td></tr>")
                body.append("</tbody></table>")

        # Trades Table (Conditional, using user logic)
        if trade_items:
            body.append("<p style='margin-top: 15px;'><b>Trade</b></p>") # Header like user template
            body.append(f'<table style="{table_style}">')
            # Headers: Item Name, Stock #, Amount (as requested implicitly)
            body.append(f"<thead><tr><th style='{th_style}'>Item</th><th style='{th_style}'>Stock #</th><th style='{th_style} text-align: right;'>Amount</th></tr></thead><tbody>")
            for line in trade_items:
                try:
                    # Format: {productname} {stock#} {amount}
                    name_part, stock_part = line.split(" STK#", 1); name = name_part.strip('" '); stock, amount_str = stock_part.split(" $", 1); body.append(f"<tr><td style='{td_style}'>{html.escape(name)}</td><td style='{td_style}'>{html.escape(stock)}</td><td style='{td_right_style}'>${html.escape(amount_str)}</td></tr>")
                except ValueError:
                    body.append(f"<tr><td colspan='3' style='{td_style}'><i>Error parsing: {html.escape(line)}</i></td></tr>")
            body.append("</tbody></table>")

        # Work Order (Conditional, using user logic)
        if self.work_order_required.isChecked():
            duration = self.work_order_hours.text().strip() or "[Duration Not Specified]"
            charge_to_wo = first_stock_number # Charge to first equipment stock#
            body.append(f"<p style='margin-top: 15px;'>Please create a work order for {html.escape(duration)} and charge to {html.escape(charge_to_wo)}</p>")

        # Signature (Matches user template)
        body.append(f"<p style='margin-top: 15px;'>PFW and spreadsheet have been updated. {html.escape(salesperson)} to collect.</p>")
        body.append("</body></html>")
        html_body = "".join(body)

        # --- Action 1: Show HTML Preview in Dock ---
        if self.main_window and hasattr(self.main_window, 'email_preview_view') and hasattr(self.main_window, 'email_preview_dock'):
            try:
                # Use setHtml for QTextEdit to render basic HTML
                self.main_window.email_preview_view.setHtml(html_body)
                self.main_window.email_preview_dock.setVisible(True)
                self.main_window.email_preview_dock.raise_()
                self._show_status_message("Email preview generated in dock.", 3000)
            except Exception as e:
                print(f"Error showing email preview dock: {e}")
                QMessageBox.warning(self, "Preview Error", f"Could not display email preview:\n{e}")
        else:
            print("Warning: Cannot show email preview dock. Main window reference missing or dock not set up.")
            QMessageBox.warning(self, "Preview Error", "Could not find the email preview window.")

        # --- Action 2: Open Mailto Link (with minimal body) ---
        mailto_body = "Please copy and paste the formatted content from the Email Preview window here.\n\n" # Added newline for spacing
        mailto_url = f"mailto:{recipients_string}?subject={quote(subject)}&body={quote(mailto_body)}"
        try:
            if len(mailto_url) > 1800: # Be conservative with mailto length limits
                 print(f"Warning: Mailto URL is very long ({len(mailto_url)} chars), may exceed client limits.")
                 QMessageBox.warning(self, "Email Link Warning", "The generated email link is very long and might not open correctly or be truncated in your email client.")

            success = webbrowser.open_new_tab(mailto_url)
            if success:
                 print(f"Attempted to open mail client for draft.")
                 # Don't overwrite status bar message from preview
                 # self._show_status_message("Opening default email client...", 5000)
            else:
                 print("webbrowser.open_new_tab returned False. Mail client might not be configured or URL too long/invalid.")
                 QMessageBox.warning(self, "Email Client Error", "Could not automatically open email client.\nIs a default email application set up?")
        except Exception as e:
            print(f"Error opening mailto link: {e}")
            QMessageBox.critical(self, "Email Client Error", f"Could not open email client:\n{e}")


    def generate_csv_and_email(self):
        """Generates CSV, saves to SharePoint, and then generates email."""
        if self.generate_csv(): self.generate_email()
        else: print("CSV generation/saving failed, email generation skipped.")

    def apply_styles(self):
        """Applies stylesheets to the form widgets."""
        self.setStyleSheet(""" AMSDealForm { background-color: #f1f3f5; } QGroupBox { font-weight: bold; font-size: 16px; color: #367C2B; background-color: #ffffff; border: 1px solid #d1d5db; border-radius: 8px; padding: 25px 10px 10px 10px; margin-top: 10px; } QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0 5px; left: 10px; color: #111827; } QLineEdit, QComboBox, QSpinBox { border: 1px solid #d1d5db; border-radius: 6px; padding: 8px; font-size: 14px; background-color: #ffffff; color: #1f2937; } QLineEdit:focus, QComboBox:focus, QSpinBox:focus { border: 2px solid #367C2B; padding: 7px; } QLineEdit[readOnly="true"] { background-color: #e9ecef; color: #6b7280; } QPushButton { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #FFDE00, stop:1 #e6c700); color: #000000; border: 1px solid #dca100; padding: 8px 15px; border-radius: 6px; font-size: 14px; font-weight: bold; } QPushButton:hover { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffe633, stop:1 #FFDE00); border: 1px solid #e6c700; } QPushButton:pressed { background: #e6c700; } QPushButton#reset_btn { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #f87171, stop:1 #dc2626); color: white; border: 1px solid #b91c1c; } QPushButton#reset_btn:hover { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ef4444, stop:1 #dc2626); border: 1px solid #991b1b; } QPushButton#reset_btn:pressed { background: #b91c1c; } QPushButton#delete_line_btn, QPushButton#save_draft_btn, QPushButton#load_draft_btn { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #d1d5db, stop:1 #9ca3af); color: #1f2937; border: 1px solid #6b7280; } QPushButton#delete_line_btn:hover, QPushButton#save_draft_btn:hover, QPushButton#load_draft_btn:hover { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #e5e7eb, stop:1 #d1d5db); border: 1px solid #4b5563; } QPushButton#delete_line_btn:pressed, QPushButton#save_draft_btn:pressed, QPushButton#load_draft_btn:pressed { background: #9ca3af; } QListWidget { background-color: #ffffff; border: 1px solid #d1d5db; border-radius: 6px; padding: 5px; font-size: 13px; color: #374151; } QCompleter::popup { border: 1px solid #9ca3af; border-radius: 6px; background-color: #ffffff; font-size: 13px; padding: 2px; } QCompleter::popup::item:selected { background-color: #e0f2fe; color: #0c4a6e; } QCheckBox { font-size: 14px; color: #333; spacing: 5px; } QCheckBox::indicator { width: 16px; height: 16px; } QCheckBox::indicator:checked { background-color: #367C2B; border: 1px solid #2a5d24; border-radius: 3px; } QCheckBox::indicator:unchecked { background-color: white; border: 1px solid #9ca3af; border-radius: 3px; } QLabel#errorLabel { color: #dc2626; font-size: 12px; } """)
        self.reset_btn.setObjectName("reset_btn"); self.delete_line_btn.setObjectName("delete_line_btn"); self.save_draft_btn.setObjectName("save_draft_btn"); self.load_draft_btn.setObjectName("load_draft_btn")
        for btn in [self.reset_btn, self.delete_line_btn, self.save_draft_btn, self.load_draft_btn]: self.style().unpolish(btn); self.style().polish(btn)

    def get_csv_lines(self):
        """Returns the generated CSV lines (if any)."""
        if not self.csv_lines: print("Warning: get_csv_lines called but no CSV lines were generated.")
        return self.csv_lines

# --- CSV Output Dialog ---
class CSVOutputDialog(QMessageBox):
     def __init__(self, csv_content, parent=None, form=None):
         super().__init__(parent); self.setWindowTitle("Generated CSV Preview"); self.setIcon(QMessageBox.Information)
         self.setText("CSV Content Generated (See console log or saved file)."); self.setStandardButtons(QMessageBox.Ok); print("--- CSV DIALOG (Placeholder - Full content logged/saved) ---")

# --- Main execution (for testing AMSDealForm independently) ---
if __name__ == '__main__':
    try: from PyQt5.QtWebEngineWidgets import QWebEngineView
    except ImportError: print("ERROR: PyQtWebEngine not installed."); app_temp = QApplication(sys.argv); QMessageBox.critical(None, "Missing Dependency", "PyQtWebEngine not installed."); sys.exit(1)
    try: from dotenv import load_dotenv; script_dir = os.path.dirname(__file__); dotenv_path_script = os.path.join(script_dir, '.env'); dotenv_path_parent = os.path.join(os.path.abspath(os.path.join(script_dir, os.pardir)), '.env'); dotenv_path_to_use = None
    except ImportError: print("Warning: python-dotenv not installed.")
    if os.path.exists(dotenv_path_script): dotenv_path_to_use = dotenv_path_script
    elif os.path.exists(dotenv_path_parent): dotenv_path_to_use = dotenv_path_parent
    if dotenv_path_to_use: print(f"Loading .env from: {dotenv_path_to_use}"); load_dotenv(dotenv_path=dotenv_path_to_use, verbose=True)
    else: print(f"Warning: .env file not found.")
    app = QApplication(sys.argv)
    if hasattr(Qt, 'AA_EnableHighDpiScaling'): QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'): QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app.setStyle("Fusion")
    class MockMainWindowForTest(QMainWindow): # Need main window context for status bar
        def __init__(self):
            super().__init__()
            self.setStatusBar(QLabel("Mock status bar"))
            # Add dummy preview dock for testing
            self.email_preview_dock = QDockWidget("Email Preview", self)
            self.email_preview_view = QTextEdit() # Use QTextEdit for testing
            self.email_preview_dock.setWidget(self.email_preview_view)
            self.addDockWidget(Qt.BottomDockWidgetArea, self.email_preview_dock)
            self.email_preview_dock.setVisible(False)
    window = MockMainWindowForTest()
    mock_sp_manager = SharePointExcelManager() # Use the dummy if real one not available/configured
    test_form = AMSDealForm(main_window=window, sharepoint_manager=mock_sp_manager)
    test_form.setWindowTitle("AMS Deal Form (Standalone Test)"); test_form.resize(1000, 750); test_form.show()
    sys.exit(app.exec_())

