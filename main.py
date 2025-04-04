import sys
import os
import webbrowser # Import webbrowser for opening links
from PyQt5.QtWidgets import (QApplication, QMainWindow, QStackedWidget,
                             QListWidget, QListWidgetItem, QLabel, QWidget,
                             QVBoxLayout, QHBoxLayout, QMessageBox, QPushButton,
                             QSizePolicy, QScrollArea, QDockWidget,
                             QTextEdit, QStatusBar) # <-- Added QTextEdit and QStatusBar
from PyQt5 import QtGui
from PyQt5.QtCore import Qt, QSize, QUrl # Added QUrl
# --- Check for PyQtWebEngine ---
try:
    from PyQt5.QtWebEngineWidgets import QWebEngineView # Added QWebEngineView
except ImportError:
     print("ERROR: PyQtWebEngine is not installed.")
     print("Please install it using: pip install PyQtWebEngine")
     # Need a temporary app object to show message box if import fails early
     app_temp = QApplication(sys.argv)
     QMessageBox.critical(None, "Missing Dependency",
                          "Required package PyQtWebEngine is not installed.\nPlease install it using:\npip install PyQtWebEngine")
     sys.exit(1)
# --- End Check ---
import logging
import asyncio # Import asyncio
# Import asyncqt event loop
try:
    from asyncqt import QEventLoop
except ImportError:
    print("ERROR: asyncqt not installed. Run 'pip install asyncqt'. Async features will fail.")
    QEventLoop = None # Set to None if import fails

# Import module classes
from AMSDealForm import AMSDealForm
from HomeModule import HomeModule
from JDQuoteModule import JDQuoteModule
from CalendarModule import CalendarModule
from CalculatorModule import CalculatorModule
from ReceivingModule import ReceivingModule
from CSVEditor import CSVEditor
from RecentDealsModule import RecentDealsModule
from PriceBookModule import PriceBookModule
from UsedInventoryModule import UsedInventoryModule
from TrafficAuto import run_automation # Import the automation function
# Import SharePointManager (assuming it exists now)
try:
    from SharePointManager import SharePointExcelManager
except ImportError:
    SharePointExcelManager = None
    print("CRITICAL ERROR: SharePointManager.py not found or import failed.")
    # Define dummy if needed, although checks for None should suffice
    class SharePointExcelManager:
        def __init__(self): print("ERROR: Using Dummy SharePointExcelManager")
        def update_excel_data(self, data): return False
        def send_html_email(self, r, s, b): return False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Initialize paths
        self.base_path = os.path.dirname(os.path.abspath(__file__))
        self.data_path = os.path.join(self.base_path, "data")
        self.logs_path = os.path.join(self.base_path, "logs")
        os.makedirs(self.data_path, exist_ok=True)
        os.makedirs(self.logs_path, exist_ok=True)

        # Setup logging
        self.setup_logging()
        self.logger.info("Application starting...")

        # --- Instantiate SharePoint Manager ONCE ---
        self.sp_manager = None # Initialize as None
        if SharePointExcelManager: # Only instantiate if import succeeded
            try:
                self.sp_manager = SharePointExcelManager()
            except Exception as e:
                 self.logger.critical(f"Failed to initialize SharePoint Manager: {e}", exc_info=True)
                 QMessageBox.critical(self, "Startup Error", f"Failed to initialize SharePoint connection:\n{e}\n\nSharePoint features will be disabled.")
                 self.sp_manager = None # Ensure it's None if init fails
        else:
             QMessageBox.critical(self, "Startup Error", "SharePointManager module not found.\nSharePoint features will be disabled.")
        # ---

        # Initialize data structures (required files for editors)
        self.required_files = {
            'products.csv': ["ProductCode", "ProductName", "Price", "JDQName"],
            'parts.csv': ["Part Number", "Part Name"],
            'customers.csv': ["Name", "CustomerNumber"],
            'salesmen.csv': ["Name", "Email"]
        }

        # Load logo
        self.logo_path = os.path.join(self.base_path, "logo.png")
        self.logo_pixmap = QtGui.QPixmap(self.logo_path) if os.path.exists(self.logo_path) else QtGui.QPixmap()
        if self.logo_pixmap.isNull():
             self.logger.warning(f"Logo file not found or failed to load: {self.logo_path}")

        # --- Button Styling ---
        self.active_button_style = """
            QPushButton {
                color: #2a5d24; /* Dark Green text */
                background-color: #FFDE00; /* John Deere Yellow */
                border: none;
                padding: 10px 15px;
                text-align: left;
                font-size: 14px;
                font-weight: bold;
            }"""
        self.inactive_button_style = """
            QPushButton {
                color: white;
                background-color: #367C2B; /* Slightly lighter green */
                border: none;
                padding: 10px 15px;
                text-align: left;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #4a9c3e; /* Lighter green on hover */
            }
            QPushButton:pressed {
                background-color: #2a5d24; /* Darker green when pressed */
            }"""
        self.current_active_module_button = None # Track the active button

        # Set window properties
        self.setWindowTitle("AMSDeal Application")
        self.resize(1200, 800)
        icon_path = os.path.join(self.base_path, "BRIapp.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QtGui.QIcon(icon_path))
        else:
            self.logger.warning(f"Window icon file not found: {icon_path}")

        # --- Create Dock Widgets and Web Views ---
        # JD Portal Dock
        self.jd_portal_dock = QDockWidget("John Deere Portal", self)
        self.jd_portal_dock.setAllowedAreas(Qt.RightDockWidgetArea | Qt.LeftDockWidgetArea | Qt.BottomDockWidgetArea | Qt.TopDockWidgetArea)
        self.jd_portal_view = QWebEngineView()
        self.jd_portal_dock.setWidget(self.jd_portal_view)
        self.jd_portal_dock.setVisible(False)
        self.addDockWidget(Qt.RightDockWidgetArea, self.jd_portal_dock)

        # Email Preview Dock
        self.email_preview_dock = QDockWidget("Email Body Preview", self)
        self.email_preview_dock.setAllowedAreas(Qt.RightDockWidgetArea | Qt.LeftDockWidgetArea | Qt.BottomDockWidgetArea | Qt.TopDockWidgetArea)
        self.email_preview_view = QTextEdit() # Use QTextEdit for simple HTML display & copy
        self.email_preview_view.setReadOnly(True)
        self.email_preview_view.setStyleSheet("font-family: sans-serif; font-size: 10pt;") # Basic styling
        self.email_preview_dock.setWidget(self.email_preview_view)
        self.email_preview_dock.setVisible(False) # Start hidden
        self.addDockWidget(Qt.BottomDockWidgetArea, self.email_preview_dock)
        # --- End Dock Creation ---

        # Setup UI components
        self.setup_ui() # Creates sidebar, stack, etc.

        # *** Add Debugging before init_modules ***
        try:
            current_loop = asyncio.get_event_loop()
            self.logger.debug(f"Event loop BEFORE init_modules: {current_loop} (Running: {current_loop.is_running()})")
        except RuntimeError as loop_err:
             self.logger.error(f"Error getting event loop BEFORE init_modules: {loop_err}")
        # *** End Debugging ***

        # Initialize and load modules into the stack
        self.init_modules() # Instantiates modules and adds to stack

        # --- Connect Signals AFTER modules are instantiated ---
        if "RecentDeals" in self.modules and self.modules["RecentDeals"] is not None:
             # Ensure the signal actually exists on the module before connecting
             if hasattr(self.modules["RecentDeals"], 'reload_deal_requested'):
                  self.modules["RecentDeals"].reload_deal_requested.connect(self.handle_reload_deal)
                  self.logger.info("Connected RecentDeals signal to handler.")
             else:
                  self.logger.error("RecentDeals module instance does not have 'reload_deal_requested' signal.")
        else:
             self.logger.error("Could not connect RecentDeals signal - module instance not found.")
        # --- End Signal Connection ---


        # Set default view AFTER modules are initialized and added
        # Activate the "Home" button and set the initial stack widget
        home_button = self.module_buttons.get("Home")
        home_widget = self.modules.get("Home")
        if home_button and home_widget:
             home_button.setStyleSheet(self.active_button_style)
             self.current_active_module_button = home_button
             home_widget_index = self.stack.indexOf(home_widget)
             if home_widget_index >= 0:
                 self.stack.setCurrentIndex(home_widget_index) # Use index
                 self.logger.info("Set initial widget to Home.")
             else:
                 self.logger.error("Home widget not found in stack during initial setup.")
        else:
             self.logger.error("Could not set initial Home view (button or widget missing).")

        # Add Status Bar
        self.setStatusBar(QStatusBar(self))
        self.statusBar().showMessage("Ready", 5000) # Initial status message

        self.logger.info("MainWindow initialization complete.")

    def setup_logging(self):
        """Configure application logging"""
        log_file = os.path.join(self.logs_path, 'app.log')
        log_level = logging.DEBUG # Use DEBUG for more detail if needed
        logging.basicConfig(
            filename=log_file,
            level=log_level,
            format='%(asctime)s - %(levelname)s - [%(funcName)s] - %(message)s', # Added function name
            filemode='a' # Append mode
        )
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(log_level) # Match level
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(funcName)s] - %(message)s') # Added function name
        console_handler.setFormatter(formatter)

        self.logger = logging.getLogger('AMSDeal')
        # Prevent adding handlers multiple times if init is called again somehow
        if not self.logger.handlers:
             self.logger.addHandler(logging.FileHandler(log_file))
             self.logger.addHandler(console_handler)
        self.logger.setLevel(log_level) # Set overall logger level


    def get_data_status(self):
        """Return dictionary of file existence status in the data directory"""
        status = {}
        for filename_key, config in self.required_files.items():
             full_path = os.path.join(self.data_path, filename_key)
             status[filename_key] = os.path.exists(full_path)
        self.logger.debug(f"Data file status: {status}")
        return status

    def create_sidebar_button(self, text, icon_path=None):
        """Helper function to create and style sidebar buttons."""
        button = QPushButton(text)
        button.setStyleSheet(self.inactive_button_style) # Start inactive
        button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        button.setMinimumHeight(40) # Ensure buttons have a decent height
        if icon_path and os.path.exists(icon_path):
             button.setIcon(QtGui.QIcon(icon_path))
             button.setIconSize(QSize(20, 20))
        return button

    def setup_ui(self):
        """Initialize main window UI components"""
        self.logger.info("Setting up UI...")
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # --- Sidebar Setup ---
        sidebar_widget = QWidget()
        sidebar_widget.setFixedWidth(250)
        sidebar_widget.setStyleSheet("background-color: #2a5d24;") # John Deere Green
        sidebar_outer_layout = QVBoxLayout(sidebar_widget) # Use outer layout for margins
        sidebar_outer_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_outer_layout.setSpacing(0)

        # Logo
        logo_label = QLabel()
        if not self.logo_pixmap.isNull():
            logo_label.setPixmap(self.logo_pixmap.scaled(200, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else:
             logo_label.setText("Logo Missing")
             logo_label.setStyleSheet("color: white; font-size: 16px; padding: 20px 0;")
        logo_label.setAlignment(Qt.AlignCenter)
        logo_label.setStyleSheet("padding: 15px 0; border-bottom: 1px solid #367C2B;")
        sidebar_outer_layout.addWidget(logo_label)

        # --- Scroll Area for Buttons ---
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("QScrollArea { border: none; }") # Remove scroll area border
        scroll_area.verticalScrollBar().setStyleSheet("QScrollBar { width: 0px; }") # Hide scrollbar visually
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff) # No horizontal scrollbar

        # Container widget inside scroll area
        scroll_content_widget = QWidget()
        sidebar_layout = QVBoxLayout(scroll_content_widget) # Layout for buttons INSIDE scroll area
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(0) # No spacing between buttons
        sidebar_layout.setAlignment(Qt.AlignTop) # Align buttons to the top

        # --- Module Buttons ---
        self.module_buttons = {} # Dictionary to hold module buttons
        # *** Added Used Inventory ***
        self.module_definitions = [
            ("🏠 Dashboard", "Home"),
            ("📝 Deal Form", "DealForm"),
            ("🕒 Recent Deals", "RecentDeals"),
            ("📖 Price Book", "PriceBook"),
            ("🚜 Used Inventory", "UsedInventory"), # <-- Added Here
            ("🧮 Calculator", "Calculator"),
            ("📅 Calendar", "Calendar"),
            ("📊 Products Editor", "ProductsEditor"),
            ("🔩 Parts Editor", "PartsEditor"),
            ("👥 Customers Editor", "CustomersEditor"),
            ("👔 Salesmen Editor", "SalesmenEditor"),
            ("🔐 JD Quotes", "JDQuotes"),
            ("📦 Receiving", "Receiving")
        ]

        # Add module buttons to the layout and dictionary
        for display_text, module_key in self.module_definitions:
            button = self.create_sidebar_button(display_text)
            button.clicked.connect(lambda checked, key=module_key: self.sidebar_button_clicked(key))
            sidebar_layout.addWidget(button)
            self.module_buttons[module_key] = button # Store button reference

        # --- Separator ---
        separator = QWidget()
        separator.setFixedHeight(1)
        separator.setStyleSheet("background-color: #367C2B;")
        sidebar_layout.addWidget(separator)
        sidebar_layout.addSpacing(10) # Add some space

        # --- Action Buttons ---
        self.btn_jd_portal = self.create_sidebar_button("🔑 JD Portal")
        # Removed Google SSO button
        self.btn_build_price = self.create_sidebar_button("💲 Build & Price")
        self.btn_whats_new = self.create_sidebar_button("✨ What's New")
        self.btn_ccms = self.create_sidebar_button("⚙️ CCMS") # Will connect differently

        sidebar_layout.addWidget(self.btn_jd_portal)
        # sidebar_layout.addWidget(self.btn_sso_google) # Removed
        sidebar_layout.addWidget(self.btn_build_price)
        sidebar_layout.addWidget(self.btn_whats_new)
        sidebar_layout.addWidget(self.btn_ccms)

        # Connect Action Button Signals
        self.btn_jd_portal.clicked.connect(self.open_jd_portal) # Renamed handler
        # Keep Build & Price and What's New opening externally for now
        self.btn_build_price.clicked.connect(lambda: webbrowser.open("https://salescenter.deere.com/#/build-price", new=2))
        self.btn_whats_new.clicked.connect(lambda: webbrowser.open("https://shorturl.at/3ArwX", new=2))
        # Change CCMS to open internally
        self.btn_ccms.clicked.connect(self.open_ccms_internal) # New handler

        sidebar_layout.addStretch() # Push version label to bottom

        # Version info
        version_label = QLabel("v1.0.1") # Example version
        version_label.setAlignment(Qt.AlignCenter)
        version_label.setStyleSheet("color: #cccccc; padding: 10px; font-size: 12px;")
        sidebar_layout.addWidget(version_label)

        # --- Final Sidebar Assembly ---
        scroll_area.setWidget(scroll_content_widget) # Put button container in scroll area
        sidebar_outer_layout.addWidget(scroll_area) # Add scroll area below logo
        main_layout.addWidget(sidebar_widget) # Add the whole sidebar to main layout

        # Stacked widget for module content
        self.stack = QStackedWidget()
        self.stack.setStyleSheet("background-color: #ffffff;") # White background for content area
        main_layout.addWidget(self.stack)

        self.logger.info("UI setup complete.")

    def init_modules(self):
        """Initialize all application modules and add them to the stack"""
        self.logger.info("Initializing modules...")
        # Instantiate modules, passing necessary arguments
        # *** Pass self.sp_manager (created in __init__) to modules that need it ***
        self.modules = {
            "Home": HomeModule(self),
            "DealForm": AMSDealForm(main_window=self, sharepoint_manager=self.sp_manager), # Pass manager
            "RecentDeals": RecentDealsModule(main_window=self, data_path=self.data_path, sharepoint_manager=self.sp_manager), # Pass manager
            "PriceBook": PriceBookModule(main_window=self, sharepoint_manager=self.sp_manager),
            "UsedInventory": UsedInventoryModule(main_window=self, sharepoint_manager=self.sp_manager), # <-- Instantiate & Pass manager
            "Calculator": CalculatorModule(self),
            "Calendar": CalendarModule(self),
            "ProductsEditor": None, # Lazy loaded
            "PartsEditor": None, # Lazy loaded
            "CustomersEditor": None, # Lazy loaded
            "SalesmenEditor": None, # Lazy loaded
            "JDQuotes": JDQuoteModule(self),
            # Pass run_automation for traffic logic
            "Receiving": ReceivingModule(main_window=self, traffic=run_automation)
        }

        # Add ALL non-None modules directly from self.modules dict
        self.logger.info("Adding instantiated modules to stack...")
        for module_key, module_instance in self.modules.items():
             if module_instance is not None:
                 # Check if widget already added (e.g., during lazy loading error handling)
                 if self.stack.indexOf(module_instance) == -1:
                     self.stack.addWidget(module_instance)
                     self.logger.info(f"Added module to stack: {module_key} ({module_instance})")
                 else:
                      self.logger.warning(f"Module {module_key} instance already in stack during init?")

        # --- Connect RecentDeals signal ---
        if "RecentDeals" in self.modules and self.modules["RecentDeals"] is not None:
            # Ensure the signal actually exists on the module before connecting
            if hasattr(self.modules["RecentDeals"], 'reload_deal_requested'):
                 self.modules["RecentDeals"].reload_deal_requested.connect(self.handle_reload_deal)
                 self.logger.info("Connected RecentDeals signal to handler.")
            else:
                 self.logger.error("RecentDeals module instance does not have 'reload_deal_requested' signal.")
        else:
            self.logger.error("Could not connect RecentDeals signal - module instance not found.")
        # --- End Connect ---

        self.logger.info("Modules initialized and initial widgets added to stack.")

    def sidebar_button_clicked(self, module_key):
        """Handles clicks on module buttons in the sidebar."""
        self.logger.info(f"Sidebar button clicked, attempting to switch to module: {module_key}")

        # Deactivate previously active button
        if self.current_active_module_button:
            self.current_active_module_button.setStyleSheet(self.inactive_button_style)

        # Activate the new button
        button = self.module_buttons.get(module_key)
        if button:
            button.setStyleSheet(self.active_button_style)
            self.current_active_module_button = button
        else:
             self.logger.warning(f"Could not find button for module key: {module_key}")

        # Lazy-load CSV editors if selected and not already loaded
        widget_to_display = self.modules.get(module_key)
        if widget_to_display is None and module_key.endswith("Editor"):
            self.logger.info(f"Lazy loading editor: {module_key}")
            try:
                filename, headers = self.get_editor_config(module_key)
                if filename and headers:
                    full_path = os.path.join(self.data_path, filename) # Use data_path
                    if os.path.exists(full_path):
                        # *** Pass sp_manager also to CSVEditor if it needs saving capability ***
                        editor = CSVEditor(full_path, headers, self) # Pass main_window ref
                        self.modules[module_key] = editor # Store the loaded editor
                        self.stack.addWidget(editor) # Add it to the stack
                        widget_to_display = editor # Use the newly created editor
                        self.logger.info(f"Successfully initialized and added editor to stack: {module_key}")
                    else:
                        raise FileNotFoundError(f"Required data file not found: {full_path}")
                else:
                    raise ValueError(f"Configuration missing for editor: {module_key}")

            except Exception as e:
                self.logger.error(f"Failed to load {module_key}: {str(e)}", exc_info=True)
                QMessageBox.warning(self, "Error", f"Could not load {module_key}:\n{str(e)}")
                # Reset button style if load fails
                if button: button.setStyleSheet(self.inactive_button_style)
                self.current_active_module_button = None # No button is active
                return # Prevent switching if loading failed

        # Switch view in the QStackedWidget
        if widget_to_display is not None:
            # Use index-based switching
            widget_index = self.stack.indexOf(widget_to_display)
            self.logger.debug(f"Widget for '{module_key}': {widget_to_display}")
            self.logger.debug(f"Attempting to find index for widget: {widget_index}")

            if widget_index != -1: # Check if widget was found in stack
                self.stack.setCurrentIndex(widget_index)
                self.logger.info(f"Switched stack to index {widget_index} for {module_key}")
            else:
                 # If it still fails, log more details about the stack
                 self.logger.error(f"Widget for {module_key} NOT FOUND in stack (index is -1). Cannot switch.")
                 self.logger.error(f"Stack count: {self.stack.count()}")
                 for i in range(self.stack.count()):
                     self.logger.error(f"  Stack index {i}: {self.stack.widget(i)}")
                 # Reset button style if switch fails
                 if button: button.setStyleSheet(self.inactive_button_style)
                 self.current_active_module_button = None
        else:
            # This case should only happen if an editor failed to load and returned early
            # Or if PriceBook/UsedInventory failed init but wasn't handled yet
            self.logger.error(f"No widget instance found or loaded for module key: {module_key}")
            # Reset button style if no widget
            if button: button.setStyleSheet(self.inactive_button_style)
            self.current_active_module_button = None

    def get_editor_config(self, editor_key):
        """Get CSV filename and headers config for editors"""
        config = {
            "ProductsEditor": ("products.csv", ["ProductCode", "ProductName", "Price", "JDQName"]),
            "PartsEditor": ("parts.csv", ["Part Number", "Part Name"]),
            "CustomersEditor": ("customers.csv", ["Name", "CustomerNumber"]),
            "SalesmenEditor": ("salesmen.csv", ["Name", "Email"])
        }
        return config.get(editor_key, (None, None))

    # --- Slot to handle reloading deal data ---
    def handle_reload_deal(self, deal_data):
        """Receives deal data from RecentDealsModule and populates Deal Form."""
        self.logger.info(f"Received request to reload deal for: {deal_data.get('customer_name')}")
        deal_form_widget = self.modules.get("DealForm")
        # Ensure DealForm is actually instantiated
        if deal_form_widget and isinstance(deal_form_widget, AMSDealForm):
             try:
                 # Populate the form
                 deal_form_widget.populate_form(deal_data)
                 # Switch view to the Deal Form by simulating button click logic
                 self.sidebar_button_clicked("DealForm") # Call handler to switch view and style button
             except Exception as e:
                 self.logger.error(f"Error populating or switching to Deal Form: {e}", exc_info=True)
                 QMessageBox.critical(self, "Error", f"Failed to reload deal into form:\n{e}")
        else:
             self.logger.error("DealForm widget instance not found or not of expected type.")
             QMessageBox.critical(self, "Error", "Deal Form module is not loaded correctly.")


    # --- Methods for Action Buttons ---
    def open_jd_portal(self):
        """Loads initial JD SSO/Portal URL and shows the dock widget."""
        self.logger.info("JD Portal button clicked.")
        # *** REPLACE WITH YOUR BEST STARTING URL (Portal/Dashboard/Login/App) ***
        jd_portal_url = "https://sso.johndeere.com/" # Example: Main portal
        # jd_portal_url = "https://sso.johndeere.com/app/johndeere_dealerpath_2" # Example: Specific App
        self.logger.info(f"Loading JD Portal URL: {jd_portal_url}")
        self.jd_portal_view.setUrl(QUrl(jd_portal_url))
        self.jd_portal_dock.setVisible(True) # Show the dock
        self.jd_portal_dock.raise_() # Bring it to the front

    def open_ccms_internal(self):
        """Loads CCMS URL into the JD Portal dock widget."""
        self.logger.info("CCMS button clicked.")
        ccms_url = "https://ccms.deere.com/"
        self.logger.info(f"Loading CCMS URL into JD Portal view: {ccms_url}")
        self.jd_portal_view.setUrl(QUrl(ccms_url)) # Load into the SAME view
        self.jd_portal_dock.setVisible(True) # Ensure dock is visible
        self.jd_portal_dock.raise_() # Bring it to the front

    # --- REMOVED open_sso_google_dialog ---

# Main execution block
if __name__ == "__main__":
    # --- Check for PyQtWebEngine ---
    try:
        from PyQt5.QtWebEngineWidgets import QWebEngineView
    except ImportError:
         print("ERROR: PyQtWebEngine is not installed.")
         print("Please install it using: pip install PyQtWebEngine")
         app_temp = QApplication(sys.argv)
         QMessageBox.critical(None, "Missing Dependency", "Required package PyQtWebEngine is not installed.\nPlease install it using:\npip install PyQtWebEngine")
         sys.exit(1)
    # --- End Check ---

    # --- High DPI Scaling ---
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    # --- Setup App and Async Event Loop ---
    app = QApplication(sys.argv)
    loop = None
    if QEventLoop:
        loop = QEventLoop(app)
        asyncio.set_event_loop(loop) # Set the loop for asyncio
        print("DEBUG: asyncqt event loop set.")
    else:
        print("ERROR: asyncqt not found, async features in modules may fail.")

    app.setStyle("Fusion")

    # --- Create and Show Main Window ---
    # Must be created AFTER setting the event loop for asyncqt integration
    window = MainWindow()
    window.show()

    # --- Run Event Loop ---
    if loop:
        print("DEBUG: Starting asyncqt event loop...")
        with loop: # Use context manager to properly handle loop lifecycle
            sys.exit(loop.run_forever()) # Run the loop until app closes
    else:
        print("DEBUG: Starting standard Qt event loop...")
        sys.exit(app.exec_()) # Fallback standard Qt loop

