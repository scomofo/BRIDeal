import csv
import os # Import os for path operations
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem,
                             QPushButton, QHBoxLayout, QMessageBox, QApplication, QHeaderView) # Added QMessageBox, QApplication, QHeaderView
from PyQt5.QtCore import Qt # Added Qt

class CSVEditor(QWidget):
    """
    A QWidget for loading, viewing, editing, adding, deleting, and saving CSV data.
    """
    def __init__(self, filename, headers, main_window=None, parent=None):
        """
        Initializes the CSVEditor.

        Args:
            filename (str): The path to the CSV file.
            headers (list): A list of strings for the table column headers.
            main_window (QMainWindow, optional): Reference to the main window for status bar messages. Defaults to None.
            parent (QWidget, optional): The parent widget. Defaults to None.
        """
        super().__init__(parent)
        self.filename = filename
        self.headers = headers
        self.main_window = main_window # Store reference to main window
        self._data_changed = False # Flag to track unsaved changes

        # --- Main Layout ---
        layout = QVBoxLayout(self)

        # --- Table Widget ---
        self.table = QTableWidget()
        self.table.itemChanged.connect(self._mark_changed) # Connect itemChanged signal
        layout.addWidget(self.table)

        # --- Button Layout ---
        btn_layout = QHBoxLayout()

        # Add Row Button
        self.add_row_btn = QPushButton("Add Row")
        self.add_row_btn.setToolTip("Add a new blank row to the end of the table.")
        self.add_row_btn.clicked.connect(self.add_row)
        btn_layout.addWidget(self.add_row_btn)

        # Delete Row Button
        self.delete_row_btn = QPushButton("Delete Selected Row")
        self.delete_row_btn.setToolTip("Delete the currently selected row.")
        self.delete_row_btn.clicked.connect(self.delete_row)
        btn_layout.addWidget(self.delete_row_btn)

        # Spacer to push save button to the right (optional)
        btn_layout.addStretch()

        # Save Button
        self.save_btn = QPushButton("Save Changes")
        self.save_btn.setToolTip(f"Save changes back to {os.path.basename(self.filename)}")
        self.save_btn.clicked.connect(self.save_csv)
        btn_layout.addWidget(self.save_btn)

        layout.addLayout(btn_layout)

        # --- Load Initial Data ---
        self.load_csv()

    def _mark_changed(self, item):
        """Slot connected to itemChanged signal to mark data as modified."""
        # print(f"Item changed: row {item.row()}, col {item.column()}, text '{item.text()}'") # Debugging
        self._data_changed = True
        # Optionally enable save button only when changes are made
        # self.save_btn.setEnabled(True)

    def _show_status_message(self, message, timeout=3000):
        """Helper to show messages on main window status bar or print."""
        if self.main_window and hasattr(self.main_window, 'statusBar'):
            self.main_window.statusBar().showMessage(message, timeout)
        else:
            print(f"Status: {message}")

    def load_csv(self):
        """Loads data from the CSV file into the table widget."""
        self._data_changed = False # Reset change flag on load
        try:
            # Clear existing table contents before loading
            self.table.setRowCount(0)
            self.table.setColumnCount(len(self.headers)) # Set column count based on headers
            self.table.setHorizontalHeaderLabels(self.headers)

            # Check if file exists before trying to open
            if not os.path.exists(self.filename):
                 print(f"Warning: CSV file not found: {self.filename}. Starting with an empty table.")
                 # Optionally create the file with headers if it doesn't exist?
                 # self.save_csv(save_empty=True) # Be careful with this
                 return # Start with empty table

            # Attempt to read with utf-8, fallback to latin-1 if needed
            data = []
            encodings_to_try = ['utf-8', 'latin-1', 'windows-1252']
            file_loaded = False
            for encoding in encodings_to_try:
                try:
                    with open(self.filename, mode='r', newline='', encoding=encoding) as f:
                        # Use csv.reader for robust handling of quotes and commas
                        reader = csv.reader(f)
                        # Decide whether to skip header based on if headers were provided
                        # For now, assume the file DOES NOT contain headers,
                        # as headers are passed in the constructor.
                        # If file *does* contain headers, uncomment next line:
                        # file_headers = next(reader, None)
                        data = list(reader)
                        print(f"Loaded {self.filename} with encoding {encoding}")
                        file_loaded = True
                        break # Stop trying encodings if successful
                except UnicodeDecodeError:
                    print(f"Failed to decode {self.filename} with {encoding}, trying next...")
                    continue # Try next encoding
                except Exception as e:
                    # Catch other potential file reading errors
                    print(f"Error reading CSV file {self.filename} with {encoding}: {e}")
                    QMessageBox.critical(self, "Load Error", f"Could not read file:\n{self.filename}\n\nError: {e}")
                    return # Stop loading if a non-decoding error occurs

            if not file_loaded:
                print(f"Error: Could not load or decode {self.filename} with any supported encoding.")
                QMessageBox.critical(self, "Load Error", f"Could not load or decode file:\n{self.filename}")
                return

            # Populate the table
            self.table.setRowCount(len(data))
            for r, row in enumerate(data):
                # Ensure we don't try to write more columns than we have headers for
                max_cols = min(len(row), len(self.headers))
                for c in range(max_cols):
                    item = QTableWidgetItem(str(row[c])) # Ensure data is string
                    self.table.setItem(r, c, item)

            # Resize columns to fit content
            self.table.resizeColumnsToContents()
            # Optional: Enable sorting
            self.table.setSortingEnabled(True)
            self._show_status_message(f"Loaded {os.path.basename(self.filename)}", 3000)

        except Exception as e:
            print(f"Error in load_csv method: {e}")
            QMessageBox.critical(self, "Load Error", f"An unexpected error occurred loading:\n{self.filename}\n\nError: {e}")
            # Ensure table is cleared on error
            self.table.setRowCount(0)
            self.table.setColumnCount(len(self.headers))
            self.table.setHorizontalHeaderLabels(self.headers)

    def add_row(self):
        """Adds a new, empty row to the end of the table."""
        current_row_count = self.table.rowCount()
        self.table.insertRow(current_row_count)
        # Optionally populate with empty items
        for col in range(self.table.columnCount()):
            self.table.setItem(current_row_count, col, QTableWidgetItem(""))
        self._mark_changed(None) # Mark data as changed
        self._show_status_message("Row added. Remember to save.", 3000)
        # Scroll to the newly added row
        self.table.scrollToBottom()


    def delete_row(self):
        """Deletes the currently selected row after confirmation."""
        current_row = self.table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "Delete Row", "Please select a row to delete.")
            return

        # Get data from the first column (or a key column) for the confirmation message
        item = self.table.item(current_row, 0)
        row_identifier = f"row {current_row + 1}" + (f" (starting with '{item.text()}')" if item else "")

        reply = QMessageBox.question(self, 'Confirm Delete',
                                       f"Are you sure you want to delete {row_identifier}?",
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.table.removeRow(current_row)
            self._mark_changed(None) # Mark data as changed
            self._show_status_message("Row deleted. Remember to save.", 3000)

    def save_csv(self, save_empty=False):
        """Saves the current table data back to the CSV file."""
        if not self._data_changed and not save_empty:
            self._show_status_message("No changes to save.", 2000)
            # return # Allow saving even if no changes detected, to overwrite/create

        row_count = self.table.rowCount()
        col_count = self.table.columnCount()

        if row_count == 0 and not save_empty:
             print("Table is empty, nothing to save.")
             # Optionally ask user if they want to save an empty file
             # reply = QMessageBox.question(...)
             # if reply == QMessageBox.No: return
             pass # Allow saving an empty file if desired, but maybe warn

        print(f"Attempting to save {row_count} rows to {self.filename}...")
        try:
            with open(self.filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, quoting=csv.QUOTE_ALL) # Quote all fields for safety
                # Write headers first
                writer.writerow(self.headers)
                # Write data rows
                for row in range(row_count):
                    row_data = []
                    for col in range(col_count):
                        item = self.table.item(row, col)
                        row_data.append(item.text() if item else "") # Get text or empty string
                    writer.writerow(row_data)

            self._data_changed = False # Reset change flag after successful save
            # self.save_btn.setEnabled(False) # Optionally disable save button
            self._show_status_message(f"Changes saved successfully to {os.path.basename(self.filename)}", 3000)
            print("Save successful.")

        except IOError as e:
             print(f"Error saving CSV file {self.filename}: {e}")
             QMessageBox.critical(self, "Save Error", f"Could not write to file:\n{self.filename}\n\nError: {e}\n\nCheck file permissions or if it's open elsewhere.")
             self._show_status_message(f"Error saving file: {e}", 5000)
        except Exception as e:
             print(f"An unexpected error occurred during save: {e}")
             import traceback
             traceback.print_exc()
             QMessageBox.critical(self, "Save Error", f"An unexpected error occurred during save:\n\nError: {e}")
             self._show_status_message(f"Unexpected save error: {e}", 5000)

    def closeEvent(self, event):
        """Handle the close event to check for unsaved changes."""
        if self._data_changed:
            reply = QMessageBox.question(self, 'Unsaved Changes',
                                           "There are unsaved changes. Do you want to save before closing?",
                                           QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                                           QMessageBox.Cancel)

            if reply == QMessageBox.Save:
                self.save_csv()
                # Check if save was successful? For now, assume it was or user got an error.
                event.accept()
            elif reply == QMessageBox.Discard:
                event.accept()
            else: # Cancel
                event.ignore()
        else:
            event.accept()


# --- Example Usage (for testing CSVEditor independently) ---
if __name__ == '__main__':
    app = QApplication(sys.argv)

    # --- Create a dummy CSV for testing ---
    test_filename = "test_editor.csv"
    test_headers = ["ID", "Name", "Value", "Notes"]
    test_data = [
        ["1", "Apple", "1.50", "Red fruit"],
        ["2", "Banana", "0.75", "Yellow, curved"],
        ["3", "Orange", "1.25", "Citrus, contains comma"],
    ]
    try:
        with open(test_filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            # writer.writerow(test_headers) # Write headers to file if needed
            writer.writerows(test_data)
        print(f"Created dummy file: {test_filename}")
    except Exception as e:
        print(f"Could not create dummy file: {e}")
        sys.exit(1)
    # --- End dummy CSV creation ---


    # Create and show the editor
    # Pass the headers explicitly
    editor = CSVEditor(filename=test_filename, headers=test_headers)
    editor.setWindowTitle(f"CSV Editor - {os.path.basename(test_filename)}")
    editor.resize(600, 400) # Set a default size
    editor.show()

    sys.exit(app.exec_())
