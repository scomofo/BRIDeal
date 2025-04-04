from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel

class JDQuoteModule(QWidget):
    def __init__(self, main_window=None, sales_module=None, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("🔐 JDQuote Automation (Placeholder)"))
