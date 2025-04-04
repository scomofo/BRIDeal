from PyQt5.QtWidgets import QWidget, QLabel, QVBoxLayout

class CalendarModule(QWidget):
    def __init__(self, main_window=None):
        super().__init__()

        self.main_window = main_window  # Optional, reference to the main window if needed

        layout = QVBoxLayout(self)

        title = QLabel("Outlook Calendar - Placeholder")
        title.setStyleSheet("""
            font-size: 28px;
            font-weight: bold;
            color: #2a5d24;
            font-family: 'Segoe UI', Arial;
        """)
        layout.addWidget(title)

        text = QLabel("Calendar integration requires Microsoft Graph API setup.")
        text.setStyleSheet("""
            font-size: 16px;
            color: #555;
            font-family: 'Segoe UI', Arial;
        """)
        layout.addWidget(text)

        # Optional calendar widget for display
        # from PyQt5.QtWidgets import QCalendarWidget
        # calendar = QCalendarWidget(self)
        # layout.addWidget(calendar)
