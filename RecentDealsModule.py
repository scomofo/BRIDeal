import sys
import os
import json
import io # For csv processing
import csv # For csv processing
import re # For parsing equipment lines
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QListWidget, QListWidgetItem, QPushButton,
                             QMessageBox, QApplication)
from PyQt5.QtCore import Qt, pyqtSignal, QDate # Added QDate

# Attempt to import SharePointExcelManager, handle gracefully if missing
try:
    # Assumes SharePointManager.py is in the same directory
    from SharePointManager import SharePointExcelManager
except ImportError:
    SharePointExcelManager = None # Set to None if import fails
    print("WARNING: SharePointManager could not be imported in RecentDealsModule. Regenerate CSV will fail.")


class RecentDealsModule(QWidget):
    """
    A widget to display recent deals and allow reloading or reprocessing.
    """
    # Signal to request reloading a deal into the main Deal Form
    # Emits the dictionary containing the deal data
    reload_deal_requested = pyqtSignal(dict)

    # *** Ensure sharepoint_manager is accepted here ***
    def __init__(self, main_window=None, data_path=None, sharepoint_manager=None, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.data_path = data_path
        self.sharepoint_manager = sharepoint_manager # Store the shared manager instance
        self.recent_deals_file = os.path.join(self.data_path, "recent_deals.json") if self.data_path else None
        if not self.recent_deals_file:
             print("ERROR: Data path not provided to RecentDealsModule. Cannot load/save recent deals.")
        if SharePointExcelManager is None: # Check if import failed
            print("ERROR: SharePointManager class not available. Regenerate CSV disabled.")
        elif self.sharepoint_manager is None:
             print("WARNING: SharePointManager instance not provided to RecentDealsModule. Regenerate CSV disabled.")


        self.setup_ui()
        self.load_recent_deals()

    def _show_status_message(self, message, timeout=3000):
        """Helper to show messages on main window status bar or print."""
        if self.main_window and hasattr(self.main_window, 'statusBar'):
            self.main_window.statusBar().showMessage(message, timeout)
        else:
            print(f"Status: {message}")

    def setup_ui(self):
        """Set up the UI elements for the recent deals module."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        title = QLabel("🕒 Recent Deals")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #2a5d24; margin-bottom: 10px;")
        layout.addWidget(title)

        self.deals_list = QListWidget()
        self.deals_list.setStyleSheet("font-size: 14px;")
        self.deals_list.itemDoubleClicked.connect(self._reload_deal) # Allow double-click to reload
        layout.addWidget(self.deals_list)

        # Button layout
        btn_layout = QHBoxLayout()
        btn_layout.addStretch() # Push buttons to the right

        self.refresh_btn = QPushButton("🔄 Refresh List")
        self.refresh_btn.setToolTip("Reload the list of recent deals from the file.")
        self.refresh_btn.clicked.connect(self.load_recent_deals)
        btn_layout.addWidget(self.refresh_btn)

        self.reload_btn = QPushButton("📂 Reload Deal")
        self.reload_btn.setToolTip("Load the selected deal's data back into the Deal Form.")
        self.reload_btn.clicked.connect(self._reload_deal)
        btn_layout.addWidget(self.reload_btn)

        # Regenerate CSV button
        self.regen_csv_btn = QPushButton("📄 Regenerate CSV")
        self.regen_csv_btn.setToolTip("Regenerate CSV & Save to SharePoint for the selected deal")
        self.regen_csv_btn.clicked.connect(self._regenerate_csv)
        # Enable only if sharepoint manager is available
        self.regen_csv_btn.setEnabled(self.sharepoint_manager is not None)
        btn_layout.addWidget(self.regen_csv_btn)

        # Regenerate Email button (still placeholder)
        self.regen_email_btn = QPushButton("📧 Regenerate Email")
        self.regen_email_btn.setToolTip("Regenerate Email for the selected deal (Not Implemented)")
        self.regen_email_btn.clicked.connect(self._regenerate_email)
        self.regen_email_btn.setEnabled(False) # Disabled for now
        btn_layout.addWidget(self.regen_email_btn)

        layout.addLayout(btn_layout)

    def load_recent_deals(self):
        """Loads the recent deals from the JSON file and populates the list."""
        self.deals_list.clear()
        if not self.recent_deals_file:
            self.deals_list.addItem("Error: Recent deals file path not set.")
            return

        recent_deals = []
        try:
            if os.path.exists(self.recent_deals_file):
                with open(self.recent_deals_file, 'r', encoding='utf-8') as f:
                    recent_deals = json.load(f)
                    if not isinstance(recent_deals, list):
                         print(f"Warning: Recent deals file {self.recent_deals_file} does not contain a list. Resetting.")
                         recent_deals = []
            else:
                print("Recent deals file not found. List will be empty.")

        except (IOError, json.JSONDecodeError) as e:
            print(f"Error loading or parsing recent deals file {self.recent_deals_file}: {e}")
            QMessageBox.warning(self, "Load Error", f"Could not load recent deals:\n{e}")
            recent_deals = []
        except Exception as e:
             print(f"Unexpected error loading recent deals: {e}")
             recent_deals = []

        if not recent_deals:
            self.deals_list.addItem("No recent deals found.")
            return

        # Populate the list
        for deal_data in recent_deals:
            try:
                # Try to parse timestamp for display
                ts_str = deal_data.get("timestamp", "N/A")
                try:
                    # Handle potential timezone info if present (Python < 3.11 might need different parsing)
                    if '+' in ts_str and ts_str.rfind('+') > 15: # Basic check for timezone offset
                         ts_str = ts_str[:ts_str.rfind('+')]
                    elif 'Z' in ts_str:
                         ts_str = ts_str.replace('Z', '')

                    # Handle potential microseconds
                    if '.' in ts_str:
                        ts_str = ts_str[:ts_str.find('.')]

                    # Try parsing common formats
                    try:
                        ts_obj = datetime.fromisoformat(ts_str)
                    except ValueError:
                        ts_obj = datetime.strptime(ts_str, "%Y-%m-%dT%H:%M:%S") # Fallback format

                    display_ts = ts_obj.strftime("%Y-%m-%d %H:%M") # Format for display
                except ValueError as ts_err:
                    print(f"Warning: Could not parse timestamp '{ts_str}': {ts_err}")
                    display_ts = ts_str # Show raw string if parsing fails

                customer = deal_data.get("customer_name", "N/A")
                salesperson = deal_data.get("salesperson", "N/A")
                display_text = f"{display_ts} - {customer} ({salesperson})"
                item = QListWidgetItem(display_text)
                # Store the full deal data with the item
                item.setData(Qt.UserRole, deal_data)
                self.deals_list.addItem(item)
            except Exception as item_err:
                 print(f"Error processing recent deal item: {item_err}. Data: {deal_data}")
                 self.deals_list.addItem(f"Error loading item: {item_err}")

        self._show_status_message("Recent deals list refreshed.", 2000)


    def _get_selected_deal_data(self):
        """Gets the full data dictionary stored with the currently selected list item."""
        currentItem = self.deals_list.currentItem()
        if currentItem:
            deal_data = currentItem.data(Qt.UserRole)
            if isinstance(deal_data, dict):
                return deal_data
            else:
                 QMessageBox.warning(self, "Error", "Selected item has invalid data associated with it.")
                 return None
        else:
            QMessageBox.warning(self, "No Selection", "Please select a deal from the list first.")
            return None

    def _reload_deal(self):
        """Gets selected deal data and emits signal to reload it in Deal Form."""
        deal_data = self._get_selected_deal_data()
        if deal_data:
            print(f"Requesting reload for deal: {deal_data.get('customer_name')}")
            self.reload_deal_requested.emit(deal_data)

    def _regenerate_email(self):
        """Placeholder for regenerating email."""
        deal_data = self._get_selected_deal_data()
        if deal_data:
            QMessageBox.information(self, "Not Implemented", "Regenerate Email function is not yet implemented.")
            # TODO: Implement email regeneration using deal_data

    def _regenerate_csv(self):
        """Regenerates CSV lines from selected deal and saves to SharePoint."""
        deal_data = self._get_selected_deal_data()
        if not deal_data:
            return # Error message shown by _get_selected_deal_data

        if not self.sharepoint_manager:
            QMessageBox.critical(self, "Error", "SharePoint Manager is not available. Cannot save.")
            return

        print(f"Regenerating CSV for deal: {deal_data.get('customer_name')}")
        self._show_status_message(f"Regenerating CSV for {deal_data.get('customer_name')}...", 5000)

        # --- Replicate CSV Generation Logic ---
        # Extract data from the deal_data dictionary
        customer = deal_data.get("customer_name", "")
        salesperson = deal_data.get("salesperson", "")
        equipment_items = deal_data.get("equipment", [])
        trade_items = deal_data.get("trades", [])
        is_paid = deal_data.get("paid", False)
        is_multi_line = deal_data.get("multi_line_csv", False)

        # Use current date/time for regeneration? Or original timestamp? Let's use current.
        today = QDate.currentDate().toString("yyyy-MM-dd")
        status = "Paid" if is_paid else "Not Paid"
        payment_icon = "🟩" if is_paid else "🟥"
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S PDT")

        csv_lines = []
        try:
            eq_pattern = r'"(.*)"\s+\(Code:\s*(.*?)\)\s+STK#(.*?)\s+\$(.*)'
            if is_multi_line:
                for line in equipment_items:
                    if line:
                        match = re.match(eq_pattern, line)
                        if match:
                            equipment, _code, manual_stock, amount_str = match.groups()
                            amount = amount_str.strip()
                            output = io.StringIO(); writer = csv.writer(output, quoting=csv.QUOTE_ALL)
                            writer.writerow([payment_icon, customer, equipment, manual_stock, amount, "", "", "", "", salesperson, today, status, timestamp])
                            csv_lines.append(output.getvalue().strip())
                        else: raise ValueError(f"Cannot parse equipment line: {line}")

                first_eq_manual_stock = ""
                if equipment_items and equipment_items[0]:
                     match = re.match(eq_pattern, equipment_items[0])
                     if match: first_eq_manual_stock = match.group(3)

                for line in trade_items:
                    if line:
                        name_part, stock_part = line.split(" STK#", 1)
                        trade = name_part.strip('" '); trade_stock, trade_amount_str = stock_part.split(" $", 1)
                        trade_amount = trade_amount_str.strip()
                        attached_to = first_eq_manual_stock if first_eq_manual_stock else "N/A"
                        output = io.StringIO(); writer = csv.writer(output, quoting=csv.QUOTE_ALL)
                        writer.writerow([payment_icon, customer, "", "", "", trade, attached_to, trade_stock, trade_amount, salesperson, today, status, timestamp])
                        csv_lines.append(output.getvalue().strip())
            else: # Single line logic
                payment = payment_icon # Use correct variable name
                equipment = stock_number = amount = trade = attached_to = trade_stock = trade_amount = ""
                if equipment_items and equipment_items[0]:
                    match = re.match(eq_pattern, equipment_items[0])
                    if match: equipment, _code, stock_number, amount_str = match.groups(); amount = amount_str.strip()
                    else: raise ValueError(f"Cannot parse equipment line: {equipment_items[0]}")
                if trade_items and trade_items[0]:
                    line = trade_items[0]; name_part, stock_part = line.split(" STK#", 1); trade = name_part.strip('" '); trade_stock, trade_amount_str = stock_part.split(" $", 1); trade_amount = trade_amount_str.strip();
                    attached_to = stock_number if stock_number else "N/A"

                output = io.StringIO(); writer = csv.writer(output, quoting=csv.QUOTE_ALL)
                writer.writerow([payment, customer, equipment, stock_number, amount, trade, attached_to, trade_stock, trade_amount, salesperson, today, status, timestamp])
                csv_lines.append(output.getvalue().strip())

            if not csv_lines:
                 QMessageBox.warning(self, "Regenerate Error", "No valid CSV lines could be generated from selected deal data.")
                 return

            # --- Replicate Payload Preparation Logic ---
            data_to_save = []
            headers = ["Payment", "Customer", "Equipment", "Stock Number", "Amount",
                       "Trade", "Attached to stk#", "Trade STK#", "Amount2",
                       "Salesperson", "Email Date", "Status", "Timestamp"]
            csv_data = "\n".join(csv_lines)
            csvfile = io.StringIO(csv_data)
            reader = csv.reader(csvfile, delimiter=',', quotechar='"')
            for fields in reader:
                if len(fields) == len(headers):
                    row_dict = dict(zip(headers, fields))
                    data_to_save.append(row_dict)
                else:
                    print(f"Warning: Skipping malformed regenerated CSV line: {fields}")

            if not data_to_save:
                QMessageBox.warning(self, "Regenerate Error", "Could not parse regenerated CSV data for saving.")
                return

            # --- Call SharePoint Manager ---
            print(f"Attempting to save regenerated CSV data ({len(data_to_save)} rows) to SharePoint...")
            success = self.sharepoint_manager.update_excel_data(data_to_save)

            if success:
                self._show_status_message(f"Successfully regenerated and saved CSV for {customer}", 3000)
                QMessageBox.information(self, "Success", f"Successfully regenerated and saved CSV data for:\n{customer}")
            else:
                # Error message should be shown by update_excel_data
                self._show_status_message(f"Failed to save regenerated CSV for {customer}", 5000)

        except Exception as e:
            print(f"Error during CSV regeneration or saving: {e}")
            import traceback; traceback.print_exc()
            QMessageBox.critical(self, "Regeneration Error", f"An unexpected error occurred during CSV regeneration:\n{e}")
            self._show_status_message(f"Error regenerating CSV: {e}", 5000)


# --- Example Usage (for testing RecentDealsModule independently) ---
if __name__ == '__main__':
    app = QApplication(sys.argv)

    # Create dummy data path and recent deals file for testing
    test_data_path = "test_recent_data"
    os.makedirs(test_data_path, exist_ok=True)
    test_recent_file = os.path.join(test_data_path, "recent_deals.json")
    dummy_deals = [
        { "timestamp": datetime.now().isoformat(), "customer_name": "Test Customer 1", "salesperson": "Sales Guy", "equipment": ["\"EQ1\" (Code: C1) STK#S1 $100"], "trades": [], "parts": [], "paid": False, "multi_line_csv": False, "work_order_required": False, "work_order_charge_to": "", "work_order_hours": "", "part_location_index": 0, "last_charge_to": ""},
        { "timestamp": (datetime.now() - timedelta(days=1)).isoformat(), "customer_name": "Test Customer 2", "salesperson": "Sales Gal", "equipment": ["\"EQ2\" (Code: C2) STK#S2 $200"], "trades": ["\"Trade1\" STK#T1 $50"], "parts": ["1x PN1 Name1 Loc1 Chg1"], "paid": True, "multi_line_csv": True, "work_order_required": True, "work_order_charge_to": "S2", "work_order_hours": "2", "part_location_index": 1, "last_charge_to": "S2"}
    ]
    try:
        from datetime import timedelta # Import for dummy data
        with open(test_recent_file, 'w', encoding='utf-8') as f:
            json.dump(dummy_deals, f, indent=4)
        print(f"Created dummy recent deals file: {test_recent_file}")
    except Exception as e:
        print(f"Could not create dummy file: {e}")

    # --- Dummy MainWindow for testing signal ---
    class MockMainWindow:
        def statusBar(self):
            return self # Use self for mock status bar
        def showMessage(self, msg, timeout):
            print(f"Mock Status: {msg} ({timeout}ms)")
        def handle_reload_deal(self, deal_data):
            print("\n--- MockMainWindow received reload_deal_requested signal ---")
            print(json.dumps(deal_data, indent=2))
            print("--- Would now populate Deal Form and switch view ---")
            QMessageBox.information(None, "Reload Requested", f"Reload requested for customer:\n{deal_data.get('customer_name')}")

    mock_main = MockMainWindow()
    # --- End Dummy MainWindow ---

    # Create and show the editor
    # Need a mock sharepoint manager for testing _regenerate_csv
    class MockSPManager:
         def update_excel_data(self, data):
             print("--- MockSPManager received data to update ---")
             print(json.dumps(data, indent=2))
             print("--- Returning mock success ---")
             # return False # Test failure
             return True

    recent_module = RecentDealsModule(main_window=mock_main, data_path=test_data_path, sharepoint_manager=MockSPManager())
    recent_module.reload_deal_requested.connect(mock_main.handle_reload_deal) # Connect signal
    recent_module.setWindowTitle("Recent Deals Module (Standalone Test)")
    recent_module.resize(700, 500) # Set a default size
    recent_module.show()

    sys.exit(app.exec_())

