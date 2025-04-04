import sys # Added for QApplication potentially needed in run_receiving
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QTextEdit, QPushButton, QApplication # Added QApplication
from PyQt5.QtCore import Qt
# Remove the incorrect imports from here if you added them:
# from ReceivingModule import ReceivingModule # <- REMOVE
# from TrafficAuto import run_automation # <- REMOVE

class ReceivingModule(QWidget):
    # This __init__ correctly accepts main_window and traffic
    def __init__(self, main_window=None, traffic=None):
        super().__init__()
        self.main_window = main_window
        # Store the passed-in traffic function/object
        self.traffic = traffic

        # --- UI Setup ---
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)

        self.title = QLabel("📦 Receiving Module")
        self.title.setAlignment(Qt.AlignCenter)
        self.title.setStyleSheet("font-size: 24px; font-weight: bold; color: #2a5d24;")
        layout.addWidget(self.title)

        self.stock_input = QTextEdit()
        self.stock_input.setPlaceholderText("Paste stock numbers here (one per line)...")
        self.stock_input.setStyleSheet("font-size: 16px;")
        layout.addWidget(self.stock_input)

        self.run_btn = QPushButton("🚚 Run Receiving")
        self.run_btn.setStyleSheet("font-size: 18px; padding: 8px 16px;")
        self.run_btn.clicked.connect(self.run_receiving)
        layout.addWidget(self.run_btn)

        self.output = QTextEdit()
        self.output.setReadOnly(True)
        self.output.setStyleSheet("font-size: 15px; background: #f4f4f4; border: 1px solid #ccc;")
        layout.addWidget(self.output)
        # --- End UI Setup ---

    def run_receiving(self):
        stock_data = self.stock_input.toPlainText().strip()
        if not stock_data:
            self.output.setText("⚠️ Please enter stock numbers.")
            return

        # Check if self.traffic is actually callable before trying to call it
        if not callable(self.traffic):
            error_msg = "❌ ERROR: Traffic automation function not properly configured."
            print(error_msg) # Also print to console for debugging
            self.output.setText(error_msg)
            # Optionally show a QMessageBox
            # from PyQt5.QtWidgets import QMessageBox # Import if using
            # QMessageBox.critical(self, "Configuration Error", "The traffic automation function was not passed correctly during setup.")
            return

        stock_list = [s.strip() for s in stock_data.splitlines() if s.strip()]
        self.output.setText("🚚 Processing the following stock numbers:\n" + "\n".join(stock_list))
        QApplication.processEvents() # Allow UI to update before potentially long process

        results = []
        for stock in stock_list:
            self.output.append(f"\nProcessing {stock}...")
            QApplication.processEvents() # Update UI during loop
            try:
                # Call the traffic function that was passed in during __init__
                # Ensure TrafficAuto.run_automation returns something meaningful or handles its own errors well
                result = self.traffic(stock) # This calls run_automation if passed correctly
                # Check what run_automation returns (currently None in the pasted code)
                status = result if result is not None else "Completed (No return value)"
                results.append(f"✅ {stock}: {status}")
                self.output.append(f" -> {status}")
            except Exception as e:
                # Catch errors during the traffic call
                error_str = str(e)
                results.append(f"❌ {stock}: {error_str}")
                self.output.append(f" -> ERROR: {error_str}")
                # Log full traceback for debugging
                import traceback
                print(f"Error processing stock {stock}:")
                traceback.print_exc()

            QApplication.processEvents() # Update UI after each item

        self.output.append("\n\n📋 Processing complete.")

