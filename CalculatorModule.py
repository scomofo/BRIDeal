# core/calculator.py
from PyQt5.QtWidgets import (
    QWidget, QLineEdit, QLabel, QVBoxLayout,
    QHBoxLayout, QPushButton, QGridLayout
)
from PyQt5.QtCore import Qt

class CalculatorModule(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.layout = QVBoxLayout(self)
        self.form_layout = QGridLayout()
        self.layout.addLayout(self.form_layout)

        self.usd_cost = self._create_input("Enter USD Cost")
        self.exchange_rate = self._create_input("Enter USD-CAD Exchange Rate", "1")
        self.cad_cost = self._create_input("Enter CAD Cost")
        self.markup = self._create_input("Enter Markup (%)")
        self.margin = self._create_input("Enter Margin (%)")
        self.revenue = self._create_input("Enter Revenue (CAD $)")

        entries = [
            ("USD Cost ($)", self.usd_cost),
            ("Exchange Rate", self.exchange_rate),
            ("CAD Cost ($)", self.cad_cost),
            ("Markup (%)", self.markup),
            ("Margin (%)", self.margin),
            ("Revenue (CAD $)", self.revenue),
        ]

        for i, (label, field) in enumerate(entries):
            lbl = QLabel(label)
            lbl.setStyleSheet("font-size: 16px; color: #2a5d24; font-family: 'Segoe UI', Arial;")
            self.form_layout.addWidget(lbl, i, 0)
            self.form_layout.addWidget(field, i, 1)

        for field in [self.usd_cost, self.exchange_rate, self.cad_cost, self.markup, self.margin, self.revenue]:
            field.textChanged.connect(self.calculate)

        self._add_clear_button()

    def _create_input(self, placeholder, default_text=""):
        box = QLineEdit()
        box.setPlaceholderText(placeholder)
        box.setText(default_text)
        box.setStyleSheet("""
            border: 1px solid #d1d5db;
            border-radius: 8px;
            padding: 8px 10px;
            font-size: 15px;
            font-family: 'Segoe UI', Arial;
            background: #ffffff;
        """)
        return box

    def _add_clear_button(self):
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        clear_btn = QPushButton("Clear")
        clear_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #FFDE00, stop:1 #ffe633);
                color: #2a5d24;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-size: 16px;
                font-weight: bold;
                font-family: 'Segoe UI', Arial;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #ffe633, stop:1 #FFDE00);
            }
            QPushButton:pressed {
                background: #e6c700;
            }
        """)
        clear_btn.clicked.connect(self.clear_fields)
        button_layout.addWidget(clear_btn)

        self.layout.addLayout(button_layout)
        self.layout.addStretch()

    def clear_fields(self):
        self.usd_cost.clear()
        self.exchange_rate.setText("1")
        self.cad_cost.clear()
        self.markup.clear()
        self.margin.clear()
        self.revenue.clear()

    def calculate(self):
        def get_float(text):
            try:
                return float(text.strip()) if text.strip() else None
            except ValueError:
                return None

        def format_number(value):
            if value is None:
                return ""
            return str(int(value)) if value == int(value) else f"{value:.2f}"

        usd = get_float(self.usd_cost.text())
        rate = get_float(self.exchange_rate.text())
        cad = get_float(self.cad_cost.text())
        markup = get_float(self.markup.text())
        margin = get_float(self.margin.text())
        revenue = get_float(self.revenue.text())

        sender = self.sender()
        sender_name = {
            self.usd_cost: "usd",
            self.exchange_rate: "rate",
            self.cad_cost: "cad",
            self.markup: "markup",
            self.margin: "margin",
            self.revenue: "revenue"
        }.get(sender)

        if usd is not None and rate is not None and sender_name != "cad":
            cad = usd * rate
            self.cad_cost.setText(format_number(cad))
        elif cad is not None and rate and sender_name != "usd" and rate != 0:
            usd = cad / rate
            self.usd_cost.setText(format_number(usd))
        elif usd and cad and sender_name != "rate" and usd != 0:
            rate = cad / usd
            self.exchange_rate.setText(format_number(rate))

        if cad and markup and sender_name != "revenue":
            revenue = cad * (1 + markup / 100)
            self.revenue.setText(format_number(revenue))
        elif cad and revenue and sender_name != "markup" and cad != 0:
            markup = ((revenue / cad) - 1) * 100
            self.markup.setText(format_number(markup))
        elif revenue and markup and sender_name != "cad" and (1 + markup / 100) != 0:
            cad = revenue / (1 + markup / 100)
            self.cad_cost.setText(format_number(cad))

        if markup is not None and sender_name != "margin":
            margin = (markup / (100 + markup)) * 100 if markup != 0 else 0.0
            self.margin.setText(format_number(margin))
        elif margin is not None and sender_name != "markup" and margin != 100:
            markup = (margin / (100 - margin)) * 100
            self.markup.setText(format_number(markup))

        if usd and rate and sender_name in ["markup", "margin", "revenue"]:
            cad = usd * rate
            self.cad_cost.setText(format_number(cad))
