import sys
import os
import webbrowser  # Import webbrowser for opening links
import logging
import traceback # Add traceback import at the top for use in multiple blocks
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QStackedWidget, QListWidget, QListWidgetItem,
    QLabel, QWidget, QVBoxLayout, QHBoxLayout, QMessageBox, QPushButton,
    QSizePolicy, QScrollArea, QDockWidget, QTextEdit, QStatusBar, QFrame 
)
from PyQt5 import QtGui
from PyQt5.QtCore import Qt, QSize, QUrl, QTimer

# --- Check for PyQtWebEngine ---
# (Using the corrected import - only QWebEngineView)
try:
    from PyQt5.QtWebEngineWidgets import QWebEngineView
    WEBENGINE_AVAILABLE = True
    print("PyQtWebEngine found.")
except ImportError as e:
    WEBENGINE_AVAILABLE = False
    # Optional: Keep detailed logging during debugging
    # print(f"ERROR: PyQtWebEngine failed to import. Details: {e}")
    # traceback.print_exc()
    print("ERROR: PyQtWebEngine features could not be enabled (Import failed).") # Simpler message
    print("Please ensure it is installed correctly: pip install PyQtWebEngine")
    # Defer the critical message until QApplication is properly initialized
# --- End Check ---


# --- Import Application Modules ---
# Adding detailed error catching to all module imports (as done before)

# HomeModule
try:
    from HomeModule import HomeModule
    print("Successfully imported HomeModule")
except ImportError as e:
    print(f"--- FAILED TO IMPORT HomeModule ---")
    traceback.print_exc()
    print(f"--- END HomeModule IMPORT ERROR ---")
    HomeModule = None
except Exception as e:
    print(f"--- UNEXPECTED ERROR IMPORTING HomeModule ---")
    traceback.print_exc()
    print(f"--- END HomeModule UNEXPECTED ERROR ---")
    HomeModule = None

# AMSDealForm
try:
    from AMSDealForm import AMSDealForm
    print("Successfully imported AMSDealForm")
except ImportError as e:
    print(f"--- FAILED TO IMPORT AMSDealForm ---")
    traceback.print_exc()
    print(f"--- END AMSDealForm IMPORT ERROR ---")
    AMSDealForm = None
except Exception as e:
    print(f"--- UNEXPECTED ERROR IMPORTING AMSDealForm ---")
    traceback.print_exc()
    print(f"--- END AMSDealForm UNEXPECTED ERROR ---")
    AMSDealForm = None

# JDQuoteModule
try:
    from JDQuoteModule import JDQuoteModule
    print("Successfully imported JDQuoteModule")
except ImportError as e:
    print(f"--- FAILED TO IMPORT JDQuoteModule ---")
    traceback.print_exc()
    print(f"--- END JDQuoteModule IMPORT ERROR ---")
    JDQuoteModule = None
except Exception as e:
    print(f"--- UNEXPECTED ERROR IMPORTING JDQuoteModule ---")
    traceback.print_exc()
    print(f"--- END JDQuoteModule UNEXPECTED ERROR ---")
    JDQuoteModule = None

# CalendarModule
try:
    from CalendarModule import CalendarModule
    print("Successfully imported CalendarModule")
except ImportError as e:
    print(f"--- FAILED TO IMPORT CalendarModule ---")
    traceback.print_exc()
    print(f"--- END CalendarModule IMPORT ERROR ---")
    CalendarModule = None
except Exception as e:
    print(f"--- UNEXPECTED ERROR IMPORTING CalendarModule ---")
    traceback.print_exc()
    print(f"--- END CalendarModule UNEXPECTED ERROR ---")
    CalendarModule = None

# CalculatorModule
try:
    from CalculatorModule import CalculatorModule
    print("Successfully imported CalculatorModule")
except ImportError as e:
    print(f"--- FAILED TO IMPORT CalculatorModule ---")
    traceback.print_exc()
    print(f"--- END CalculatorModule IMPORT ERROR ---")
    CalculatorModule = None
except Exception as e:
    print(f"--- UNEXPECTED ERROR IMPORTING CalculatorModule ---")
    traceback.print_exc()
    print(f"--- END CalculatorModule UNEXPECTED ERROR ---")
    CalculatorModule = None

# ReceivingModule and TrafficAuto
try:
    from ReceivingModule import ReceivingModule
    run_automation = None  # Default to None
    if ReceivingModule:
        try:
            from TrafficAuto import run_automation  # Import the automation function
            print("Successfully imported TrafficAuto.run_automation")
        except ImportError:
            print("--- FAILED TO IMPORT TrafficAuto ---")
            traceback.print_exc()
            print("--- END TrafficAuto IMPORT ERROR ---")
            # run_automation remains None
        except Exception:
            print("--- UNEXPECTED ERROR IMPORTING TrafficAuto ---")
            traceback.print_exc()
            print("--- END TrafficAuto UNEXPECTED ERROR ---")
            # run_automation remains None
    print(
        f"Successfully imported ReceivingModule "
        f"(run_automation is {'set' if run_automation else 'None'})"
    )

except ImportError:
    print("--- FAILED TO IMPORT ReceivingModule ---")
    traceback.print_exc()
    print("--- END ReceivingModule IMPORT ERROR ---")
    ReceivingModule = None
    run_automation = None  # Ensure None if ReceivingModule fails
except Exception:
    print("--- UNEXPECTED ERROR IMPORTING ReceivingModule ---")
    traceback.print_exc()
    print("--- END ReceivingModule UNEXPECTED ERROR ---")
    ReceivingModule = None
    run_automation = None  # Ensure None on other errors

# CSVEditor
try:
    from CSVEditor import CSVEditor
    print("Successfully imported CSVEditor")
except ImportError as e:
    print(f"--- FAILED TO IMPORT CSVEditor ---")
    traceback.print_exc()
    print(f"--- END CSVEditor IMPORT ERROR ---")
    CSVEditor = None
except Exception as e:
    print(f"--- UNEXPECTED ERROR IMPORTING CSVEditor ---")
    traceback.print_exc()
    print(f"--- END CSVEditor UNEXPECTED ERROR ---")
    CSVEditor = None

# RecentDealsModule
try:
    from RecentDealsModule import RecentDealsModule
    print("Successfully imported RecentDealsModule")
except ImportError as e:
    print(f"--- FAILED TO IMPORT RecentDealsModule ---")
    traceback.print_exc()
    print(f"--- END RecentDealsModule IMPORT ERROR ---")
    RecentDealsModule = None
except Exception as e:
    print(f"--- UNEXPECTED ERROR IMPORTING RecentDealsModule ---")
    traceback.print_exc()
    print(f"--- END RecentDealsModule UNEXPECTED ERROR ---")
    RecentDealsModule = None

# PriceBookModule
try:
    from PriceBookModule import PriceBookModule
    print("Successfully imported PriceBookModule")
except ImportError as e:
    print(f"--- FAILED TO IMPORT PriceBookModule ---")
    traceback.print_exc()
    print(f"--- END PriceBookModule IMPORT ERROR ---")
    PriceBookModule = None
except Exception as e:
    print(f"--- UNEXPECTED ERROR IMPORTING PriceBookModule ---")
    traceback.print_exc()
    print(f"--- END PriceBookModule UNEXPECTED ERROR ---")
    PriceBookModule = None

# UsedInventoryModule
try:
    from UsedInventoryModule import UsedInventoryModule
    print("Successfully imported UsedInventoryModule")
except ImportError as e:
    print(f"--- FAILED TO IMPORT UsedInventoryModule ---")
    traceback.print_exc()
    print(f"--- END UsedInventoryModule IMPORT ERROR ---")
    UsedInventoryModule = None
except Exception as e:
    print(f"--- UNEXPECTED ERROR IMPORTING UsedInventoryModule ---")
    traceback.print_exc()
    print(f"--- END UsedInventoryModule UNEXPECTED ERROR ---")
    UsedInventoryModule = None

# --- Import SharePointManager (optional) ---
try:
    from SharePointManager import SharePointExcelManager
    SHAREPOINT_AVAILABLE = True
    print("SharePointManager found.")
except ImportError:
    SHAREPOINT_AVAILABLE = False
    SharePointExcelManager = None # Ensure it's None if import fails
    print("WARNING: SharePointManager.py not found or import failed. SharePoint features disabled.")
    # Define a dummy class if needed for type hinting or basic checks
    class SharePointExcelManager:
        def __init__(self, *args, **kwargs): print("ERROR: Using Dummy SharePointExcelManager") # Add *args, **kwargs
        def update_excel_data(self, data): return False
        def send_html_email(self, r, s, b): return False
# --- End Imports ---

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # --- Initialize paths ---
        self.base_path = os.path.dirname(os.path.abspath(__file__))
        # Determine if running as a bundled app
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            # Path for bundled resources (often same as executable dir or _MEIPASS)
            self.app_data_path = os.path.dirname(sys.executable)
            # Data path might be relative to executable
            self.data_path = os.path.join(self.app_data_path, "data")
        else:
            # Development path
            self.app_data_path = self.base_path # Resources relative to script
            self.data_path = os.path.join(self.base_path, "data") # Data relative to script

        # Ensure data path exists
        os.makedirs(self.data_path, exist_ok=True)

        # Logs path (usually relative to script/executable base)
        self.logs_path = os.path.join(self.base_path, "logs")
        os.makedirs(self.logs_path, exist_ok=True)

        # --- Setup logging ---
        self.setup_logging()
        self.logger.info(f"Application starting...")
        self.logger.info(f"Base Path: {self.base_path}")
        self.logger.info(f"App Data Path: {self.app_data_path}")
        self.logger.info(f"Data Path: {self.data_path}")
        self.logger.info(f"Logs Path: {self.logs_path}")

        # --- Instantiate SharePoint Manager ---
        self.sp_manager = None # Initialize as None
        if SHAREPOINT_AVAILABLE and SharePointExcelManager:
            try:
                # Assuming SharePointExcelManager takes no arguments, adjust if needed
                self.sp_manager = SharePointExcelManager()
                self.logger.info("SharePoint Manager Initialized.")
            except Exception as e:
                self.logger.critical(f"Failed to initialize SharePoint Manager: {e}", exc_info=True)
                # Use QTimer to show message after window is potentially shown
                QTimer.singleShot(100, lambda: QMessageBox.critical(self, "Startup Error", f"Failed to initialize SharePoint connection:\n{e}\n\nSharePoint features will be disabled."))
                self.sp_manager = None # Ensure it's None on failure
        else:
            self.logger.warning("SharePoint Manager not available or import failed.")
            # Optional: Use QTimer to show a non-critical warning later
            # QTimer.singleShot(100, lambda: QMessageBox.warning(self, "Startup Info", "SharePointManager module not found.\nSharePoint features will be disabled."))

        # --- Define required data files and headers ---
        self.required_files = {
            'products.csv': ["ProductCode", "ProductName", "Price", "JDQName"],
            'parts.csv': ["Part Number", "Part Name"],
            'customers.csv': ["Name", "CustomerNumber"],
            'salesmen.csv': ["Name", "Email"]
        }
        # Check/Create required files
        self.check_and_create_data_files()

        # --- Load logo ---
        # Look for logo relative to where resources are expected (app_data_path)
        self.logo_path = os.path.join(self.app_data_path, "logo.png")
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

        # --- Set window properties ---
        self.setWindowTitle("AMSDeal Application")
        self.resize(1200, 800)
        icon_path = os.path.join(self.app_data_path, "BRIapp.ico") # Look for icon relative to app path
        if os.path.exists(icon_path):
            self.setWindowIcon(QtGui.QIcon(icon_path))
        else:
            self.logger.warning(f"Window icon file not found: {icon_path}")

        # --- Create Dock Widgets (conditional on WebEngine) ---
        self.jd_portal_dock = None
        self.jd_portal_view = None
        if WEBENGINE_AVAILABLE:
            self.logger.info("Creating JD Portal dock widget.")
            self.jd_portal_dock = QDockWidget("John Deere Portal", self)
            self.jd_portal_dock.setAllowedAreas(Qt.RightDockWidgetArea | Qt.LeftDockWidgetArea | Qt.BottomDockWidgetArea | Qt.TopDockWidgetArea)
            self.jd_portal_view = QWebEngineView()
            self.jd_portal_dock.setWidget(self.jd_portal_view)
            self.jd_portal_dock.setVisible(False) # Start hidden
            self.addDockWidget(Qt.RightDockWidgetArea, self.jd_portal_dock)
        else:
            self.logger.warning("WebEngine not available, JD Portal dock will not be created.")

        # Email Preview Dock (always available)
        self.logger.info("Creating Email Preview dock widget.")
        self.email_preview_dock = QDockWidget("Email Body Preview", self)
        self.email_preview_dock.setAllowedAreas(Qt.RightDockWidgetArea | Qt.LeftDockWidgetArea | Qt.BottomDockWidgetArea | Qt.TopDockWidgetArea)
        self.email_preview_view = QTextEdit()
        self.email_preview_view.setReadOnly(True)
        self.email_preview_view.setStyleSheet("font-family: sans-serif; font-size: 10pt;")
        self.email_preview_dock.setWidget(self.email_preview_view)
        self.email_preview_dock.setVisible(False)
        self.addDockWidget(Qt.BottomDockWidgetArea, self.email_preview_dock)
        # --- End Dock Creation ---

        # Setup core UI (sidebar, stack)
        self.setup_ui()

        # Initialize modules and add to stack
        self.init_modules()

        # Connect signals (e.g., for reloading deals)
        self.connect_module_signals()

        # Set default view
        self.set_initial_view()

        # Add Status Bar
        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready", 5000)

        self.logger.info("MainWindow initialization complete.")

    def setup_logging(self):
        """Configure application logging"""
        log_file = os.path.join(self.logs_path, 'app.log')
        # Use DEBUG level for more detailed logs during troubleshooting
        log_level = logging.DEBUG # CHANGED TO DEBUG
        # log_level = logging.INFO # Use INFO for production

        # Configure root logger
        logging.basicConfig(
            level=log_level,
            format='%(asctime)s - %(levelname)s - [%(name)s:%(funcName)s:%(lineno)d] - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S', # Added date format
            handlers=[
                logging.FileHandler(log_file, mode='a', encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        # Get application-specific logger
        self.logger = logging.getLogger('AMSDeal')
        self.logger.setLevel(log_level) # Ensure app logger respects the level
        self.logger.info("Logging configured.")
        # Optionally suppress verbose library logs
        logging.getLogger("yfinance").setLevel(logging.WARNING)
        # Suppress MSAL info logs unless debugging auth issues
        msal_log_level = logging.WARNING if log_level <= logging.INFO else logging.INFO
        logging.getLogger("msal").setLevel(msal_log_level)


    def check_and_create_data_files(self):
        """Checks for required CSV files and creates them with headers if missing."""
        self.logger.info("Checking required data files...")
        for filename, headers in self.required_files.items():
            full_path = os.path.join(self.data_path, filename)
            if not os.path.exists(full_path):
                self.logger.warning(f"Data file not found: {filename}. Creating at {full_path}")
                try:
                    with open(full_path, 'w', newline='', encoding='utf-8') as f:
                        # Write header row (quoting headers is safer)
                        f.write(','.join(f'"{h}"' for h in headers) + '\n')
                    self.logger.info(f"Created data file with headers: {filename}")
                except Exception as e:
                    self.logger.error(f"Failed to create data file {filename}: {e}", exc_info=True)
                    # Show error but allow app to continue if possible
                    QMessageBox.critical(self, "File Creation Error", f"Could not create required data file:\n{filename}\n\nError: {e}")
            else:
                self.logger.debug(f"Data file exists: {filename}")


    def get_data_status(self):
        """Return dictionary of file existence status in the data directory"""
        status = {}
        for filename_key in self.required_files.keys():
            full_path = os.path.join(self.data_path, filename_key)
            status[filename_key] = os.path.exists(full_path)
        self.logger.debug(f"Data file status: {status}")
        return status

    def create_sidebar_button(self, text, icon_name=None):
        """Helper function to create and style sidebar buttons."""
        button = QPushButton(text)
        button.setStyleSheet(self.inactive_button_style)
        button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        button.setMinimumHeight(40)
        if icon_name:
            # Prefer icons from resources folder if it exists, fallback to app_data_path/icons
            resource_icon_path = os.path.join(self.base_path, "resources", "icons", f"{icon_name}.png")
            app_data_icon_path = os.path.join(self.app_data_path, "icons", f"{icon_name}.png")

            final_icon_path = None
            if os.path.exists(resource_icon_path):
                 final_icon_path = resource_icon_path
            elif os.path.exists(app_data_icon_path):
                 final_icon_path = app_data_icon_path

            if final_icon_path:
                button.setIcon(QtGui.QIcon(final_icon_path))
                button.setIconSize(QSize(20, 20)) # Adjust icon size if needed
                self.logger.debug(f"Loaded sidebar icon for '{icon_name}' from {final_icon_path}")
            else:
                self.logger.warning(f"Sidebar icon not found in resources or app data: icons/{icon_name}.png")
        return button

    def setup_ui(self):
        """Initialize main window UI components (sidebar, stack)."""
        self.logger.info("Setting up UI...")
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # --- Sidebar ---
        sidebar_widget = QWidget()
        sidebar_widget.setFixedWidth(250)
        sidebar_widget.setStyleSheet("background-color: #2a5d24;")
        sidebar_outer_layout = QVBoxLayout(sidebar_widget)
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

        # Scroll Area for Buttons
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("QScrollArea { border: none; background-color: #2a5d24; }")
        scroll_area.verticalScrollBar().setStyleSheet("""
            QScrollBar:vertical { border: none; background: #2a5d24; width: 8px; margin: 0; }
            QScrollBar::handle:vertical { background: #FFDE00; min-height: 20px; border-radius: 4px; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; border: none; background: none; }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: none; }""")
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        scroll_content_widget = QWidget()
        scroll_content_widget.setStyleSheet("background-color: #2a5d24;")
        sidebar_layout = QVBoxLayout(scroll_content_widget)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(0)
        sidebar_layout.setAlignment(Qt.AlignTop)

        # --- CORRECTED BUTTON CREATION AND ADDING LOGIC ---
        # Create ALL button instances first
        self.btn_home = self.create_sidebar_button("Dashboard", "home")
        self.btn_deal_form = self.create_sidebar_button("Deal Form", "form")
        self.btn_recent_deals = self.create_sidebar_button("Recent Deals", "recent")
        self.btn_price_book = self.create_sidebar_button("Price Book", "book")
        self.btn_used_inventory = self.create_sidebar_button("Used Inventory", "inventory")
        self.btn_calculator = self.create_sidebar_button("Calculator", "calculator")
        self.btn_calendar = self.create_sidebar_button("Calendar", "calendar")
        self.btn_jd_quotes = self.create_sidebar_button("JD Quotes", "quote")
        self.btn_receiving = self.create_sidebar_button("Receiving", "receiving")
        # Create editor buttons too
        self.btn_products_editor = self.create_sidebar_button("Products Editor", "edit")
        self.btn_parts_editor = self.create_sidebar_button("Parts Editor", "edit")
        self.btn_customers_editor = self.create_sidebar_button("Customers Editor", "edit")
        self.btn_salesmen_editor = self.create_sidebar_button("Salesmen Editor", "edit")

        # --- Conditionally add buttons to layout and connect signals ---
        self.module_buttons = {} # Reset dictionary
        # Map the key used in sidebar_button_clicked to the button variable and the class variable
        self.button_map = {
            "Home": (self.btn_home, HomeModule),
            "DealForm": (self.btn_deal_form, AMSDealForm),
            "RecentDeals": (self.btn_recent_deals, RecentDealsModule),
            "PriceBook": (self.btn_price_book, PriceBookModule),
            "UsedInventory": (self.btn_used_inventory, UsedInventoryModule),
            "Calculator": (self.btn_calculator, CalculatorModule),
            "Calendar": (self.btn_calendar, CalendarModule),
            "JDQuotes": (self.btn_jd_quotes, JDQuoteModule),
            "Receiving": (self.btn_receiving, ReceivingModule),
            # Editors use CSVEditor class
            "ProductsEditor": (self.btn_products_editor, CSVEditor),
            "PartsEditor": (self.btn_parts_editor, CSVEditor),
            "CustomersEditor": (self.btn_customers_editor, CSVEditor),
            "SalesmenEditor": (self.btn_salesmen_editor, CSVEditor),
        }

        # Define the desired order
        self.module_definitions_config = [
            ("Dashboard", "Home", "home"),
            ("Deal Form", "DealForm", "form"),
            ("Recent Deals", "RecentDeals", "recent"),
            ("Price Book", "PriceBook", "book"),
            ("Used Inventory", "UsedInventory", "inventory"),
            ("Calculator", "Calculator", "calculator"),
            ("Calendar", "Calendar", "calendar"),
            ("Products Editor", "ProductsEditor", "edit"),
            ("Parts Editor", "PartsEditor", "edit"),
            ("Customers Editor", "CustomersEditor", "edit"),
            ("Salesmen Editor", "SalesmenEditor", "edit"),
            ("JD Quotes", "JDQuotes", "quote"),
            ("Receiving", "Receiving", "receiving")
        ]
        self.active_module_definitions = [] # Reset this list

        # Iterate through the desired order, check availability, and add
# Iterate through the desired order, check availability, and add
        for display_text, module_key, icon_name in self.module_definitions_config:
            button_tuple = self.button_map.get(module_key)
            if button_tuple:
                button_instance, module_class_ref = button_tuple
                self.logger.debug(
                    f"Checking for module key: '{module_key}'"
                )  # Log key being checked
                if module_class_ref:  # Check if the CLASS variable is not None (import succeeded)
                    self.logger.info(
                        f"Check for '{module_key}': Found class reference directly: {module_class_ref}"
                    )
                    self.logger.info(f"Creating sidebar button for '{module_key}'.")
                    self.active_module_definitions.append(
                        (display_text, module_key, icon_name)
                    )  # Add to active list
                    sidebar_layout.addWidget(button_instance)
                    # Use lambda with default argument to capture correct key
                    button_instance.clicked.connect(
                        lambda checked, key=module_key: self.sidebar_button_clicked(
                            key
                        )
                )
                self.module_buttons[
                    module_key
                ] = button_instance  # Store for later use (styling etc)
            else:
                # Log why it's None
                if module_key in globals() and globals()[module_key] is None:
                    self.logger.warning(
                        f"Check for '{module_key}': Class variable exists in globals() but is None (Import likely failed). Sidebar button skipped."
                    )
                elif module_key not in self.module_buttons:  # Check against the instance variable
                    self.logger.error(
                        f"Check for '{module_key}': This key from config was not found in button_map!"
                    )
                else:  # Variable doesn't exist or wasn't handled? Unlikely.
                    self.logger.warning(
                        f"Check for '{module_key}': Class reference resolved to None. Sidebar button skipped."
                    )
        else:
            self.logger.error(
                f"Configuration error: Module key '{module_key}' from config list not found in button_map."
            )
            # --- Continue with Separator, Action Buttons etc. ---
            sidebar_layout.addStretch(1)  # Pushes subsequent items down

            separator = QFrame()  # Make sure QFrame is imported from QtWidgets
            separator.setFrameShape(QFrame.HLine)
            separator.setFrameShadow(QFrame.Sunken)
            separator.setStyleSheet(
                "QFrame { border: 1px solid #367C2B; margin: 10px 5px; }"
            )
            sidebar_layout.addWidget(separator)

            # Action Buttons
            self.btn_jd_portal = self.create_sidebar_button("JD Portal", "key")
            self.btn_build_price = self.create_sidebar_button(
                "Build & Price", "build"
            )
            self.btn_whats_new = self.create_sidebar_button("What's New", "info")
            self.btn_ccms = self.create_sidebar_button("CCMS", "ccms")

            sidebar_layout.addWidget(self.btn_jd_portal)
            sidebar_layout.addWidget(self.btn_build_price)
            sidebar_layout.addWidget(self.btn_whats_new)
            sidebar_layout.addWidget(self.btn_ccms)

            # Disable web buttons if needed
            if not WEBENGINE_AVAILABLE:
                self.btn_jd_portal.setEnabled(False)
                self.btn_jd_portal.setToolTip(
                    "JD Portal requires PyQtWebEngine (not installed)"
                )
                self.btn_ccms.setEnabled(False)
                self.btn_ccms.setToolTip(
                    "CCMS requires PyQtWebEngine (not installed)"
                )

            # Connect Action Button Signals
            if WEBENGINE_AVAILABLE:
                self.btn_jd_portal.clicked.connect(self.open_jd_portal)
                self.btn_ccms.clicked.connect(self.open_ccms_internal)
            self.btn_build_price.clicked.connect(
                lambda: webbrowser.open(
                    "https://salescenter.deere.com/#/build-price", new=2
                )
            )
            self.btn_whats_new.clicked.connect(
                lambda: webbrowser.open("https://shorturl.at/3ArwX", new=2)
            )  # Shortened URL

            sidebar_layout.addStretch(2)  # Add more stretch before version

            # Version info
            # TODO: Load version dynamically (e.g., from a file or variable)
            try:
                # Example: Load from a VERSION file
                version_file = os.path.join(self.base_path, "VERSION")
                if os.path.exists(version_file):
                    with open(version_file, "r") as vf:
                        app_version = vf.read().strip()
                else:
                    app_version = "v1.0.2"  # Fallback
            except Exception:
                app_version = "v1.0.2"  # Fallback on error
            version_label = QLabel(app_version)
            version_label.setAlignment(Qt.AlignCenter)
            version_label.setStyleSheet(
                "color: #cccccc; padding: 10px; font-size: 11px;"
            )
            sidebar_layout.addWidget(version_label)

            # Final Sidebar Assembly
            scroll_area.setWidget(scroll_content_widget)
            sidebar_outer_layout.addWidget(scroll_area)
            main_layout.addWidget(sidebar_widget)

            # Stacked widget for content
            self.stack = QStackedWidget()
            self.stack.setStyleSheet(
                "background-color: #F0F0F0;"
            )  # Light gray content background
            main_layout.addWidget(self.stack, 1)  # Give stack stretch factor

            self.logger.info("UI setup complete.")


    def init_modules(self):
        """Initialize available application modules and add them to the stack."""
        self.logger.info("Initializing modules...")
        self.modules = {} # Dictionary to store module instances {module_key: instance}

        # Define modules to initialize, check if class exists before trying
        # Use the global variables which might be None if import failed
        module_init_map = {
            "Home": HomeModule,
            "DealForm": AMSDealForm,
            "RecentDeals": RecentDealsModule,
            "PriceBook": PriceBookModule,
            "UsedInventory": UsedInventoryModule,
            "Calculator": CalculatorModule,
            "Calendar": CalendarModule,
            "JDQuotes": JDQuoteModule,
            "Receiving": ReceivingModule
        }

        # Iterate through module_init_map to initialize modules
        for key, module_class_ref in module_init_map.items():
            if module_class_ref:  # Check if class variable is not None (import succeeded)
                self.logger.info(f"Instantiating module: {key}")
                try:
                    # Prepare arguments dynamically, starting empty
                    args = {}
                    # Conditionally add main_window if not Calculator
                    if key != "Calculator":
                        args['main_window'] = self

                    # Add other arguments based on the key
                    if key in ["DealForm", "RecentDeals", "PriceBook", "UsedInventory"]:
                        # Only add sp_manager if it exists and SHAREPOINT_AVAILABLE is True
                        if SHAREPOINT_AVAILABLE and self.sp_manager:
                            args['sharepoint_manager'] = self.sp_manager
                        elif key in ["DealForm", "RecentDeals"]: # Log warning only if SP needed but unavailable
                            self.logger.warning(f"SharePoint Manager not available for {key}. Module may have reduced functionality.")

                    if key == "RecentDeals":
                        args['data_path'] = self.data_path # Pass data path to RecentDeals

                    if key == "Receiving":
                        if run_automation: # Check if run_automation was imported
                            args['traffic'] = run_automation
                        else:
                            self.logger.warning(f"Traffic automation function not available for {key}. Module may have reduced functionality.")

                    # Instantiate the module
                    self.logger.debug(f"Instantiating {key} with args: {list(args.keys())}")
                    module_instance = module_class_ref(**args) # Use the class reference
                    self.modules[key] = module_instance
                    # Add to stack immediately
                    if self.stack.indexOf(module_instance) == -1:
                        self.stack.addWidget(module_instance)
                        self.logger.info(f"Added module to stack: {key}")
                except Exception as e:
                    self.logger.error(f"Failed to instantiate module {key}: {e}", exc_info=True)
                    self.modules[key] = None  # Mark as failed
                    # Use QTimer to show message box non-blockingly if needed, or just log
                    QTimer.singleShot(100, lambda k=key, err=e: QMessageBox.critical(self, "Module Load Error", f"Could not load essential module {k}:\n{err}"))

            else:
                # This log now accurately reflects that the import failed earlier
                self.logger.warning(f"Skipping initialization for {key} - Class reference is None (import likely failed).")
                self.modules[key] = None  # Mark as unavailable

        # Add placeholders for lazy-loaded editors using CSVEditor class reference
        if CSVEditor: # Only add placeholders if CSVEditor itself was imported
             # Iterate through the actual module keys we attempted to use
             editor_keys = [key for key in self.module_definitions_config if key[1].endswith("Editor")]
             for _, key, _ in editor_keys:
                 self.modules[key] = None  # Mark as not loaded yet
        else:
            self.logger.error("CSVEditor class failed to import. Editor modules will not be available.")


        self.logger.info("Modules initialized and widgets added to stack.")


    def connect_module_signals(self):
        """Connect signals between modules after they are initialized."""
        self.logger.info("Connecting module signals...")
        # RecentDeals -> MainWindow -> DealForm
        recent_deals_module = self.modules.get("RecentDeals")
        # Check if module instance exists AND is the correct type (or has the signal)
        # Check the class variable first (if import failed, RecentDealsModule will be None)
        if RecentDealsModule and recent_deals_module and isinstance(recent_deals_module, RecentDealsModule):
            if hasattr(recent_deals_module, 'reload_deal_requested'):
                try:
                    recent_deals_module.reload_deal_requested.connect(self.handle_reload_deal)
                    self.logger.info("Connected RecentDeals.reload_deal_requested -> MainWindow.handle_reload_deal")
                except TypeError as e:
                    self.logger.error(f"TypeError connecting RecentDeals signal: {e}.")
                except Exception as e:
                     self.logger.error(f"Unknown error connecting RecentDeals signal: {e}")
            else:
                 self.logger.warning("Could not connect RecentDeals signal - signal 'reload_deal_requested' missing from instance.")
        elif "RecentDeals" in self.modules and not recent_deals_module: # Log warning if module was expected but instance is None
             self.logger.warning("Could not connect RecentDeals signal - module instance is None (initialization likely failed).")
        elif RecentDealsModule is None: # Log if import failed
             self.logger.warning("Could not connect RecentDeals signal - module failed to import.")


        # Add other inter-module signal connections here if needed


    def set_initial_view(self):
        """Sets the initial view to the Home module if available."""
        self.logger.info("Setting initial view...")
        home_button = self.module_buttons.get("Home")
        home_widget = self.modules.get("Home")

        # Check if HomeModule class itself was imported successfully
        if HomeModule is None:
             self.logger.critical("Home module failed to import. Application cannot start correctly.")
             # Display critical error centered in the window
             error_label = QLabel("Critical Error: Home module failed to load.\nApplication cannot continue.")
             error_label.setAlignment(Qt.AlignCenter)
             error_label.setStyleSheet("font-size: 16px; color: red;")
             # Try setting central widget - might fail if UI setup also failed badly
             try: self.setCentralWidget(error_label)
             except Exception as set_widget_err: print(f"Error setting error label: {set_widget_err}")
             return # Stop further processing

        # Now check if the instance exists and was added to stack
        if home_widget and self.stack.indexOf(home_widget) != -1:
            if home_button:
                home_button.setStyleSheet(self.active_button_style)
                self.current_active_module_button = home_button
            else:
                # This warning should not happen if the button logic in setup_ui is correct
                self.logger.error("Home button not found in sidebar buttons dict, despite HomeModule being available.")

            self.stack.setCurrentWidget(home_widget) # Use setCurrentWidget for clarity
            self.logger.info(f"Set initial widget to Home ({home_widget}).")

        elif self.stack.count() > 0:
             # Home widget instance doesn't exist or wasn't added to stack, but other widgets are there
             self.logger.warning("Home module failed to initialize or add to stack. Defaulting to first available module.")
             self.stack.setCurrentIndex(0)
             first_widget = self.stack.widget(0)
             # Try to find the key and button for the first widget
             found_first = False
             for key, widget in self.modules.items():
                 if widget == first_widget and key in self.module_buttons:
                      button = self.module_buttons[key]
                      button.setStyleSheet(self.active_button_style)
                      self.current_active_module_button = button
                      self.logger.info(f"Defaulted initial view to: {key}")
                      found_first = True
                      break
             if not found_first: self.logger.error("Could not find button corresponding to the first widget in stack.")
        else:
            # No widgets in the stack at all
            self.logger.critical("No modules loaded into the stack!")
            # Display critical error centered in the window
            error_label = QLabel("Critical Error: No modules loaded successfully.\nApplication cannot continue.")
            error_label.setAlignment(Qt.AlignCenter)
            error_label.setStyleSheet("font-size: 16px; color: red;")
            try: self.setCentralWidget(error_label)
            except Exception as set_widget_err: print(f"Error setting error label: {set_widget_err}")


    def sidebar_button_clicked(self, module_key):
        """Handles clicks on module buttons in the sidebar."""
        self.logger.info(f"Sidebar button clicked, attempting to switch to module: {module_key}")

        # Deactivate previously active button
        if self.current_active_module_button:
            self.current_active_module_button.setStyleSheet(self.inactive_button_style)
            self.current_active_module_button = None # Reset tracker

        # Get the newly clicked button
        button = self.module_buttons.get(module_key)

        # Get the widget instance associated with the key
        widget_to_display = self.modules.get(module_key)

        # --- Lazy-load CSV editors if selected and not already loaded ---
        if widget_to_display is None and module_key.endswith("Editor"):
            self.logger.info(f"Lazy loading editor: {module_key}")
            if CSVEditor is None: # Check if CSVEditor class itself loaded
                self.handle_module_load_error(module_key, ImportError("CSVEditor class failed to import earlier."))
                if button: button.setStyleSheet(self.inactive_button_style) # Ensure button inactive
                return

            try:
                filename, headers = self.get_editor_config(module_key)
                if not filename or not headers:
                    raise ValueError(f"Configuration missing for editor: {module_key}")

                full_path = os.path.join(self.data_path, filename)
                if not os.path.exists(full_path):
                    self.logger.warning(f"Required editor file missing: {full_path}. Attempting creation.")
                    self.check_and_create_data_files() # Try creating it
                    if not os.path.exists(full_path):
                        raise FileNotFoundError(f"Required data file still not found after check/create: {full_path}")

                # Ensure CSVEditor class is valid before instantiating
                if not CSVEditor or not callable(CSVEditor):
                     raise TypeError("CSVEditor is not a valid callable class.")

                editor = CSVEditor(full_path, headers, main_window=self) # Pass main_window ref
                self.modules[module_key] = editor
                self.stack.addWidget(editor)
                widget_to_display = editor
                self.logger.info(f"Successfully lazy-loaded editor: {module_key}")

            except Exception as e:
                self.handle_module_load_error(module_key, e) # Handles logging, message box, button style
                return # Stop processing if load failed

        # --- Switch view in the QStackedWidget ---
        if widget_to_display is not None:
            widget_index = self.stack.indexOf(widget_to_display)
            self.logger.debug(f"Attempting switch to {module_key} ({widget_to_display}) at index {widget_index}.")

            if widget_index != -1:
                self.stack.setCurrentIndex(widget_index)
                if button:
                    button.setStyleSheet(self.active_button_style)
                    self.current_active_module_button = button # Track active button
                self.logger.info(f"Switched stack view to {module_key}")
                # Use replace to make key more readable in status bar
                display_key = module_key.replace("Module", "").replace("Editor", " Editor")
                self.status_bar.showMessage(f"Switched to {display_key}", 3000)
            else:
                self.logger.error(f"Widget for {module_key} exists but not found in stack. This indicates an issue adding widgets.")
                if button: button.setStyleSheet(self.inactive_button_style) # Ensure button inactive
                QMessageBox.critical(self, "UI Error", f"Could not switch display to {module_key} (widget not in stack).")
        else:
            # This case means self.modules[module_key] is None (and not lazy-loaded)
            # Check if the class failed import earlier for a better message
            module_class_ref = None
            map_entry = self.button_map.get(module_key) # Use button_map defined in setup_ui
            if map_entry: module_class_ref = map_entry[1]

            if module_class_ref is None:
                 error_msg = f"The module '{module_key}' failed to import earlier. Check startup logs for details."
            else:
                 error_msg = f"The module '{module_key}' failed to initialize correctly. Check startup logs for details."

            self.logger.error(f"No widget instance available for module key: {module_key}. Cannot switch view. Reason: {error_msg}")
            QMessageBox.warning(self, "Module Not Available", error_msg)
            if button: button.setStyleSheet(self.inactive_button_style) # Ensure button inactive


    def handle_module_load_error(self, module_key, error):
        """Handles errors during module loading (specifically for lazy loading)."""
        self.logger.error(f"Failed to load module '{module_key}': {error}", exc_info=True)
        QMessageBox.warning(self, "Module Load Error", f"Could not load {module_key}:\n{error}")
        # Reset the button style for the failed module
        button = self.module_buttons.get(module_key)
        if button:
            button.setStyleSheet(self.inactive_button_style)
        # Ensure no button appears active if loading failed
        if self.current_active_module_button == button:
            self.current_active_module_button = None

    def get_editor_config(self, editor_key):
        """Get CSV filename and headers config for editors"""
        # Assumes editor_key format like "ProductsEditor"
        base_name = editor_key.replace("Editor", "").lower()
        filename = f"{base_name}.csv"
        headers = self.required_files.get(filename)
        if headers:
            return filename, headers
        else:
            self.logger.error(f"No header configuration found in self.required_files for filename '{filename}' (derived from key '{editor_key}')")
            return None, None


    def handle_reload_deal(self, deal_data):
        """Receives deal data from RecentDealsModule and populates Deal Form."""
        if not deal_data:
            self.logger.warning("handle_reload_deal received empty data.")
            return

        customer = deal_data.get('customer_name', 'N/A')
        self.logger.info(f"Received request to reload deal for customer: {customer}")
        deal_form_widget = self.modules.get("DealForm")

        # Check if the class itself was imported successfully before checking instance
        if AMSDealForm is None:
            self.logger.error("Cannot reload deal: AMSDealForm module failed to import earlier.")
            QMessageBox.critical(self, "Module Error", "Deal Form module is not available (failed to import).")
            return

        if deal_form_widget and isinstance(deal_form_widget, AMSDealForm):
            try:
                deal_form_widget.populate_form(deal_data) # Call populate method on the instance
                self.sidebar_button_clicked("DealForm") # Switch view and style button
                self.status_bar.showMessage(f"Loaded deal for {customer}", 5000)
            except Exception as e:
                self.logger.error(f"Error populating or switching to Deal Form: {e}", exc_info=True)
                QMessageBox.critical(self, "Reload Error", f"Failed to reload deal into form:\n{e}")
        else:
             # If AMSDealForm is not None, but deal_form_widget is None or wrong type
             self.logger.error("DealForm widget instance not found or is None/wrong type. Check initialization logs for 'DealForm'.")
             QMessageBox.critical(self, "Module Error", "Deal Form module did not initialize correctly. Check logs.")


    # --- Methods for Action Buttons ---
    def open_jd_portal(self):
        """Loads JD Portal URL into the integrated browser dock."""
        if not WEBENGINE_AVAILABLE or not self.jd_portal_dock or not self.jd_portal_view:
            self.logger.error("Cannot open JD Portal: WebEngine not available or dock/view not initialized.")
            QMessageBox.warning(self, "Feature Unavailable", "The integrated JD Portal requires PyQtWebEngine, which is not installed.")
            return

        self.logger.info("JD Portal button clicked.")
        jd_portal_url = "https://dealerportal.deere.com/" # Configurable start URL?
        self.logger.info(f"Loading JD Portal URL: {jd_portal_url}")
        self.jd_portal_view.setUrl(QUrl(jd_portal_url))
        self.jd_portal_dock.setVisible(True)
        self.jd_portal_dock.raise_() # Bring dock to front
        self.status_bar.showMessage("Loading JD Portal...", 3000)

    def open_ccms_internal(self):
        """Loads CCMS URL into the integrated browser dock."""
        if not WEBENGINE_AVAILABLE or not self.jd_portal_dock or not self.jd_portal_view:
            self.logger.error("Cannot open CCMS: WebEngine not available or dock/view not initialized.")
            QMessageBox.warning(self, "Feature Unavailable", "The integrated CCMS view requires PyQtWebEngine, which is not installed.")
            return

        self.logger.info("CCMS button clicked.")
        ccms_url = "https://ccms.deere.com/"
        self.logger.info(f"Loading CCMS URL into JD Portal view: {ccms_url}")
        self.jd_portal_view.setUrl(QUrl(ccms_url))
        self.jd_portal_dock.setVisible(True)
        self.jd_portal_dock.raise_() # Bring dock to front
        self.status_bar.showMessage("Loading CCMS...", 3000)


    # Inside class MainWindow(QMainWindow): in main.py

    def closeEvent(self, event):
        """Handle the main window closing."""

        self.logger.info(
            "MainWindow closeEvent triggered. Cleaning up modules..."
        )
        all_threads_stopped = True

        # Iterate through instantiated modules
        for module_key, module_instance in self.modules.items():
            if module_instance is not None:
                # Check if the underlying C++ object still exists (safer check)
                try:
                    # Attempting to access a property like objectName will raise
                    # the RuntimeError if the underlying object is deleted.
                    _ = module_instance.objectName()  # Check existence
                except RuntimeError:
                    self.logger.warning(
                        f"Module {module_key}'s underlying object already "
                        f"deleted. Skipping close call."
                    )
                    continue  # Skip to the next module

                close_method = None
                if hasattr(module_instance, "closeEvent") and callable(
                    module_instance.closeEvent
                ):
                    close_method = module_instance.closeEvent
                elif hasattr(module_instance, "close") and callable(
                    module_instance.close
                ):
                    close_method = module_instance.close

                if close_method:
                    try:
                        self.logger.info(
                            f"Calling close method ({close_method.__name__}) "
                            f"for module: {module_key}"
                        )
                        try:
                            close_method(event)  # Try calling with event
                        except TypeError:
                            close_method()  # Fallback: call without event
                    except RuntimeError as e:
                        self.logger.error(
                            f"Error calling close method for module "
                            f"{module_key} (object likely deleted): {e}"
                        )
                    except Exception as e:
                        self.logger.error(
                            f"Unexpected error calling close method for module "
                            f"{module_key}: {e}",
                            exc_info=True,
                        )
                else:
                    self.logger.debug(
                        f"Module {module_key} ({type(module_instance).__name__}) "
                        f"has no standard close method."
                    )

                # Explicitly check for known threads (adjust if needed)
                if (
                    module_key == "Home"
                    and HomeModule
                    and isinstance(module_instance, HomeModule)
                ):
                    if (
                        hasattr(module_instance, "scheduled_refresh")
                        and module_instance.scheduled_refresh.isRunning()
                    ):
                        self.logger.warning(
                            f"Scheduled refresh in {module_key} still running."
                        )
                        all_threads_stopped = False
                    if (
                        hasattr(module_instance, "fetcher_thread")
                        and module_instance.fetcher_thread
                        and module_instance.fetcher_thread.isRunning()
                    ):
                        self.logger.warning(
                            f"Data fetcher in {module_key} still running."
                        )
                        all_threads_stopped = False

            else:
                self.logger.debug(
                    f"Module {module_key} was not loaded or already cleaned up."
                )

        if not all_threads_stopped:
            self.logger.warning(
                "Some background threads may not have stopped cleanly."
            )

        self.logger.info("MainWindow closeEvent finished.")
        logging.shutdown()  # Ensure logs are flushed before exit
        event.accept()  # Accept the close event

# --- Main execution block ---
if __name__ == "__main__":
    # Set High DPI attributes BEFORE creating QApplication
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        print("High DPI Scaling Enabled")
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
        print("High DPI Pixmaps Enabled")

    app = QApplication(sys.argv)
    app.setStyle("Fusion") # Or "Windows", "macOS" etc.

    # Show WebEngine Missing Message AFTER app exists
    if not WEBENGINE_AVAILABLE:
        # Use a QTimer to delay the message box slightly, ensures main window is visible
        QTimer.singleShot(100, lambda: QMessageBox.warning(None, "Missing Dependency",
                            "Required package PyQtWebEngine is not installed.\n"
                            "Features like JD Portal and CCMS integration will be disabled.\n\n"
                            "Recommended install:\npip install PyQtWebEngine"))

    # Create and Show Main Window, catching critical init errors
    window = None # Ensure window is defined outside try
    try:
        window = MainWindow()
        window.show()
    except Exception as main_init_error:
        # Use basic print/logging as full logger might not be up
        print(f"CRITICAL ERROR during MainWindow initialization: {main_init_error}")
        # Use root logger as self.logger might not exist yet
        logging.critical(f"CRITICAL ERROR during MainWindow initialization: {main_init_error}", exc_info=True)
        QMessageBox.critical(None, "Fatal Initialization Error",
                             f"Application could not start:\n{main_init_error}\n\n"
                             "Check the logs/app.log file for details.")
        logging.shutdown()
        sys.exit(1)

    # Run Event Loop
    exit_code = 0
    print("Starting Qt event loop...")
    # Define traceback import here for the final except block
    import traceback
    try:
        exit_code = app.exec_()
    except Exception as e:
        # Log critical error before exiting
        logger = logging.getLogger('AMSDeal') # Try getting logger again
        if window and hasattr(window, 'logger'):
             window.logger.critical(f"Unhandled exception during application execution: {e}", exc_info=True)
        else: # Fallback print and root logging
             print(f"CRITICAL ERROR during execution: {e}")
             traceback.print_exc()
             logging.critical(f"Unhandled exception during application execution: {e}", exc_info=True)
        exit_code = 1
    finally:
        print(f"Application exiting with code: {exit_code}")
        logger = logging.getLogger('AMSDeal')
        if window and hasattr(window, 'logger'):
             window.logger.info(f"Application exiting with code: {exit_code}")
        else:
             logging.info(f"Application exiting with code: {exit_code}") # Use root logger if main logger unavailable
        logging.shutdown() # Final log flush

    sys.exit(exit_code)