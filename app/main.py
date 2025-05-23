# app/main.py
import sys
import os
import logging # Base logging
from typing import Optional, List, Dict, Any # Ensure all typing imports are present

# Ensure the project root is in sys.path
project_root_main = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if project_root_main not in sys.path:
    sys.path.insert(0, project_root_main)

from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget,
                             QLabel, QStackedWidget, QListWidget, QHBoxLayout,
                             QMessageBox, QListWidgetItem, QSizePolicy, QListView, QDialog)
from PyQt5.QtCore import Qt, pyqtSignal, QSize, QThreadPool
from PyQt5.QtGui import QFont, QIcon

# Core application components
from app.core.config import Config
from app.core.logger_config import setup_logging
from app.core.app_auth_service import AppAuthService
from app.utils.theme_manager import ThemeManager
from app.utils.cache_handler import CacheHandler
from app.utils.general_utils import set_app_user_model_id, get_resource_path

# Service Imports
from app.services.integrations.token_handler import TokenHandler
try:
    from app.services.integrations.sharepoint_manager import SharePointExcelManager as SharePointManagerService
except ImportError:
    # This logger will be used before full setup_logging if main.py is entry point
    logging.getLogger(__name__).error(
        "Failed to import SharePointExcelManager. SharePoint features will be severely affected.", exc_info=True
    )
    SharePointManagerService = None # Fallback


from app.services.api_clients.quote_builder import QuoteBuilder
from app.services.integrations.jd_auth_manager import JDAuthManager
from app.services.api_clients.jd_quote_client import JDQuoteApiClient
from app.services.api_clients.maintain_quotes_api import MaintainQuotesAPI
from app.services.integrations.jd_quote_integration_service import JDQuoteIntegrationService

# View Module Imports
from app.views.modules.deal_form_view import DealFormView
from app.views.modules.recent_deals_view import RecentDealsView
from app.views.modules.price_book_view import PriceBookView
from app.views.modules.used_inventory_view import UsedInventoryView
from app.views.modules.receiving_view import ReceivingView
from app.views.modules.csv_editors_manager_view import CsvEditorsManagerView
from app.views.modules.calculator_view import CalculatorView
from app.views.modules.jd_external_quote_view import JDExternalQuoteView
from app.views.modules.invoice_module_view import InvoiceModuleView
from app.utils.resource_checker import check_resources
# Main Window and Splash Screen
from app.views.main_window.splash_screen_view import SplashScreenView

# Logger for this module, configured properly after setup_logging()
logger = logging.getLogger(__name__)


class MainWindow(QMainWindow):
    def __init__(self, config: Config,
                 cache_handler: CacheHandler,
                 token_handler: 'TokenHandler',
                 sharepoint_manager: Optional[SharePointManagerService] = None, # Type hint uses the alias
                 jd_auth_manager: Optional['JDAuthManager'] = None,
                 jd_quote_integration_service: Optional['JDQuoteIntegrationService'] = None,
                 parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.config = config
        self.cache_handler = cache_handler
        self.token_handler = token_handler
        self.sharepoint_manager_service = sharepoint_manager # Store the passed-in object
        self.jd_auth_manager_service = jd_auth_manager
        self.jd_quote_integration_service = jd_quote_integration_service

        # Correctly initialize self.logger for the instance
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
        self.logger.info(f"DEBUG: MainWindow received sharepoint_manager of type: {type(self.sharepoint_manager_service)}")

        self.app_name = self.config.get("APP_NAME", "BRIDeal")
        self.app_version = self.config.get("APP_VERSION", "1.5.23")

        self.theme_manager = ThemeManager(config=self.config, resource_path_base="resources")

        # --- INITIALIZE THREAD POOL and NOTIFICATION MANAGER ---
        self.thread_pool = QThreadPool()
        self.logger.info(f"Global QThreadPool initialized with max threads: {self.thread_pool.maxThreadCount()}")

        self.notification_manager = None # Placeholder if you add a dedicated manager later
        self.logger.info(f"Notification manager set to: {self.notification_manager}")
        # --- END INITIALIZATION ---

        self._init_ui()
        self._load_modules()

        self.logger.info(f"{self.app_name} v{self.app_version} MainWindow initialized.")
        self.show_status_message(f"Welcome to {self.app_name}!", "info")

    def _init_ui(self):
        self.setWindowTitle(f"{self.app_name} - v{self.app_version}")
        self.setGeometry(100, 100,
                         self.config.get("window_width", 1200),
                         self.config.get("window_height", 800))

        self.theme_manager.apply_theme(self.config.get("CURRENT_THEME", "default_light.qss"))
        self.theme_manager.apply_system_style(self.config.get("QT_STYLE_OVERRIDE", "Fusion"))

        app_icon_config_key = "APP_ICON_PATH"
        default_app_icon_rel_to_resources = "icons/app_icon.png"
        app_icon_path_from_config = self.config.get(app_icon_config_key, default_app_icon_rel_to_resources)
        app_icon_abs_path = get_resource_path(app_icon_path_from_config, self.config)

        if app_icon_abs_path and os.path.exists(app_icon_abs_path):
            self.setWindowIcon(QIcon(app_icon_abs_path))
            self.logger.info(f"Application icon set from: {app_icon_abs_path}")
        else:
            self.logger.warning(
                f"Application icon not found. Config key: '{app_icon_config_key}', "
                f"Value from config/default: '{app_icon_path_from_config}', "
                f"Resolved absolute path: {app_icon_abs_path}"
            )

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        nav_panel = QWidget()
        nav_panel.setFixedWidth(280) # Adjusted width for better text fit
        nav_layout = QVBoxLayout(nav_panel)
        nav_layout.setAlignment(Qt.AlignTop) # Keep items starting from the top

        app_title_label = QLabel(self.app_name)
        app_title_font = QFont("Arial", 16, QFont.Bold)
        app_title_label.setFont(app_title_font)
        app_title_label.setAlignment(Qt.AlignCenter)
        app_title_label.setWordWrap(True) # Ensure title wraps if too long
        nav_layout.addWidget(app_title_label)

        self.nav_list = QListWidget()
        self.nav_list.itemClicked.connect(self._on_nav_item_selected)

        self.nav_list.setViewMode(QListView.ListMode) # Changed to ListMode for better readability
        self.nav_list.setIconSize(QSize(24, 24)) # Adjusted icon size for ListMode
        self.nav_list.setSpacing(5) # Reduced spacing for ListMode
        # self.nav_list.setWordWrap(True) # Already default for ListMode items
        # self.nav_list.setUniformItemSizes(False) # Already default
        # self.nav_list.setFlow(QListView.TopToBottom) # Already default

        nav_layout.addWidget(self.nav_list)
        main_layout.addWidget(nav_panel)

        self.stacked_widget = QStackedWidget()
        main_layout.addWidget(self.stacked_widget, 1) # Add stretch factor

        self.statusBar().showMessage("Ready")

    def _add_module_to_stack(self, name: str, widget: QWidget, icon_name: Optional[str] = None):
        item = QListWidgetItem(name)

        actual_icon_name = icon_name
        if hasattr(widget, 'get_icon_name') and callable(widget.get_icon_name):
            mod_icon_name = widget.get_icon_name()
            if mod_icon_name:  # Ensure module provides a non-empty name
                actual_icon_name = mod_icon_name

        icon_path = None
        if actual_icon_name:
            # Try pre-resolved paths first
            icon_path = self.module_icons_paths.get(actual_icon_name)

            # Fallback to filesystem lookup
            if not icon_path:
                icons_dir = get_resource_path("icons", self.config)
                icon_path = os.path.join(icons_dir, actual_icon_name)
                if not os.path.exists(icon_path):
                    self.logger.warning(f"Icon not found at: {icon_path}")
                    icon_path = None

        if icon_path:
            item.setIcon(QIcon(icon_path))
            self.logger.debug(f"Icon for module '{name}' (icon file: {actual_icon_name}) loaded from: {icon_path}")
        elif actual_icon_name:  # Log warning if icon name was provided but file not found
            self.logger.warning(f"Icon file '{actual_icon_name}' for module '{name}' not found or path not resolved. Searched path: {icon_path}")
        else:  # Log warning if no icon name at all
            self.logger.warning(f"No icon name provided or obtainable for module: '{name}'")

        self.nav_list.addItem(item)
        self.stacked_widget.addWidget(widget)
        self.modules[name] = widget
        self.logger.debug(f"Module '{name}' added to UI.")

    def _load_modules(self):
        self.modules: Dict[str, QWidget] = {}
        self.module_icons_paths: Dict[str, str] = {} # To store pre-resolved icon paths keyed by icon filename

        # Pre-resolve paths for known icon filenames (optional, but good for early check)
        # Modules should preferably provide their icon names via get_icon_name()
        # This dictionary is more of a fallback or for icons not tied to a specific module's class logic
        known_icon_files = [
            "new_deal_icon.png", "recent_deals_icon.png", "price_book_icon.png",
            "used_inventory_icon.png", "receiving_icon.png", "data_editors_icon.png",
            "calculator_icon.png", "jd_quote_icon.png", "invoice_icon.png"
        ]
        for icon_file in known_icon_files:
            path = get_resource_path(os.path.join("icons", icon_file), self.config)
            if path and os.path.exists(path):
                self.module_icons_paths[icon_file] = path
            else:
                self.logger.warning(f"Pre-resolved icon check: Icon '{icon_file}' not found at expected path: {path}")

        # Instantiate modules
        deal_form_view = DealFormView(
            config=self.config, main_window=self,
            jd_quote_service=self.jd_quote_integration_service,
            sharepoint_manager=self.sharepoint_manager_service,
            logger_instance=logging.getLogger("DealFormViewLogger")
        )
        self._add_module_to_stack(getattr(deal_form_view, 'MODULE_DISPLAY_NAME', "New Deal"), deal_form_view)

        recent_deals_view = RecentDealsView(
            config=self.config, main_window=self,
            logger_instance=logging.getLogger("RecentDealsViewLogger")
        )
        self._add_module_to_stack(getattr(recent_deals_view, 'MODULE_DISPLAY_NAME', "Recent Deals"), recent_deals_view)

        price_book_view = PriceBookView(
            config=self.config, main_window=self,
            sharepoint_manager=self.sharepoint_manager_service,
            logger_instance=logging.getLogger("PriceBookViewLogger")
        )
        self._add_module_to_stack(getattr(price_book_view, 'MODULE_DISPLAY_NAME', "Price Book"), price_book_view)

        used_inventory_view = UsedInventoryView(
            config=self.config, main_window=self,
            sharepoint_manager=self.sharepoint_manager_service,
            logger_instance=logging.getLogger("UsedInventoryViewLogger")
        )
        self._add_module_to_stack(getattr(used_inventory_view, 'MODULE_DISPLAY_NAME', "Used Inventory"), used_inventory_view)

        receiving_view = ReceivingView(
            config=self.config,
            logger_instance=logging.getLogger("ReceivingViewLogger"),
            thread_pool=self.thread_pool,
            notification_manager=self.notification_manager,
            main_window=self
        )
        self._add_module_to_stack(getattr(receiving_view, 'MODULE_DISPLAY_NAME', "Receiving"), receiving_view)

        csv_editors_manager = CsvEditorsManagerView(
            config=self.config, main_window=self,
            logger_instance=logging.getLogger("CsvEditorsManagerLogger")
        )
        self._add_module_to_stack(getattr(csv_editors_manager, 'MODULE_DISPLAY_NAME', "Data Editors"), csv_editors_manager)

        calculator_view = CalculatorView(
            config=self.config, main_window=self,
            logger_instance=logging.getLogger("CalculatorViewLogger")
        )
        self._add_module_to_stack(getattr(calculator_view, 'MODULE_DISPLAY_NAME', "Calculator"), calculator_view)

        jd_ext_quote_view = JDExternalQuoteView(
            config=self.config, main_window=self,
            jd_quote_integration_service=self.jd_quote_integration_service,
            logger_instance=logging.getLogger("JDExternalQuoteViewLogger")
        )
        self._add_module_to_stack(getattr(jd_ext_quote_view, 'MODULE_DISPLAY_NAME', "JD External Quote"), jd_ext_quote_view)

        invoice_module_view = InvoiceModuleView(
            config=self.config,
            main_window=self,
            jd_quote_integration_service=self.jd_quote_integration_service,
            logger_instance=logging.getLogger("InvoiceModuleViewLogger")
        )
        self._add_module_to_stack(getattr(invoice_module_view, 'MODULE_DISPLAY_NAME', "Invoice"), invoice_module_view)

        self.logger.info("--- NAV LIST CHECK ---")
        for i in range(self.nav_list.count()):
            self.logger.info(f"Nav item {i}: {self.nav_list.item(i).text()}")
        self.logger.info(f"Total nav items: {self.nav_list.count()}")
        self.logger.info(f"Modules in self.modules dict: {list(self.modules.keys())}")

        # Set default view
        if self.nav_list.count() > 0:
            default_module_title = getattr(deal_form_view, 'MODULE_DISPLAY_NAME', "New Deal")
            items = self.nav_list.findItems(default_module_title, Qt.MatchExactly)
            if items:
                self.nav_list.setCurrentItem(items[0]) # Triggers _on_nav_item_selected
            else: # Fallback to first item if specific default not found
                self.nav_list.setCurrentRow(0) # Triggers _on_nav_item_selected
                self.logger.warning(f"Default module '{default_module_title}' not found in nav list. Falling back to first item.")
        else:
            self.logger.warning("No modules loaded into the navigation list.")

    def _on_nav_item_selected(self, item: QListWidgetItem):
        module_name = item.text()
        if module_name in self.modules:
            current_module_widget = self.modules[module_name]
            self.stacked_widget.setCurrentWidget(current_module_widget)
            self.logger.debug(f"Switched to module: {module_name}")
            self.show_status_message(f"Viewing: {module_name}", "info")
            if hasattr(current_module_widget, 'load_module_data') and callable(current_module_widget.load_module_data):
                try:
                    current_module_widget.load_module_data()
                except Exception as e:
                    self.logger.error(f"Error calling load_module_data for {module_name}: {e}", exc_info=True)

    def show_status_message(self, message: str, level: str = "info", duration: int = 5000):
        self.statusBar().showMessage(message, duration if duration > 0 else 0) # 0 for persistent if duration <=0
        log_func = getattr(self.logger, level, self.logger.info)
        log_func(f"Status: {message}")

    def closeEvent(self, event):
        self.logger.info("MainWindow closing. Performing cleanup...")
        if self.thread_pool:
             self.thread_pool.clear() # Clear pending tasks
             self.thread_pool.waitForDone(-1) # Wait indefinitely for active tasks
        event.accept()

    def navigate_to_invoice(self, quote_id, dealer_account_no):
        self.logger.info(f"Navigating to invoice view for quote ID: {quote_id}")
        invoice_module_key = None
        for name, module_instance in self.modules.items():
            if isinstance(module_instance, InvoiceModuleView):
                invoice_module_key = name
                break
        
        if invoice_module_key and invoice_module_key in self.modules:
            invoice_module = self.modules[invoice_module_key]
            items = self.nav_list.findItems(invoice_module_key, Qt.MatchExactly)
            if items:
                self.nav_list.setCurrentItem(items[0]) # This will trigger _on_nav_item_selected
                # _on_nav_item_selected handles setCurrentWidget and load_module_data
            else: # Fallback if item text doesn't match key for some reason
                self.stacked_widget.setCurrentWidget(invoice_module)
                if hasattr(invoice_module, 'load_module_data'): invoice_module.load_module_data()

            if hasattr(invoice_module, 'initiate_invoice_from_quote'):
                invoice_module.initiate_invoice_from_quote(quote_id, dealer_account_no)
            self.show_status_message(f"Viewing invoice for Quote ID: {quote_id}", "info")
        else:
            self.logger.error(f"Invoice module not found. Cannot navigate for quote ID: {quote_id}")
            
    def check_jd_authentication(self):
        """
        Checks John Deere API authentication status and prompts for auth if needed
        """
        if not hasattr(self, 'jd_auth_manager_service') or not self.jd_auth_manager_service:
            self.logger.warning("JD Auth Manager not available for authentication check")
            return False
        
        if not self.jd_auth_manager_service.is_operational:
            self.logger.warning("JD Auth Manager is not operational")
            return False
        
        # Check if we have a valid token
        token = self.jd_auth_manager_service.get_access_token()
        if token:
            self.logger.info("JD API authentication token is available and valid")
            return True
        
        # If we're here, we need to authenticate
        self.logger.info("No valid JD API token found, showing authentication dialog")
        
        # Ask the user if they want to authenticate
        result = QMessageBox.question(
            self,
            "John Deere API Authentication",
            "John Deere API authentication is required for quoting features. "
            "Would you like to authenticate now?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if result == QMessageBox.Yes:
            from app.views.dialogs.jd_auth_dialog import JDAuthDialog
            auth_dialog = JDAuthDialog(self.jd_auth_manager_service, self)
            auth_dialog.auth_completed.connect(self.on_jd_auth_completed)
            return auth_dialog.exec_() == QDialog.Accepted
        
        return False

    def on_jd_auth_completed(self, success, message):
        """
        Handle the JD API authentication completion
        
        Args:
            success (bool): Whether authentication was successful
            message (str): Message from the authentication process
        """
        if success:
            self.logger.info("JD API authentication completed successfully")
            self.show_status_message("John Deere API connection established")
            
            # Update any components that depend on JD API
            if hasattr(self, 'jd_quote_integration_service'):
                self.jd_quote_integration_service.is_operational = True
        else:
            self.logger.warning(f"JD API authentication failed: {message}")
            self.show_status_message("John Deere API authentication failed")

def run_application():
    env_file_path = os.path.join(project_root_main, ".env")
    app_config = Config(env_path=env_file_path, config_file_path="config.json")

    # Setup logging as early as possible
    setup_logging(config=app_config)
    # Now logger = logging.getLogger(__name__) at the module level will use the configured setup
    logger.info(f"Starting {app_config.get('APP_NAME', 'BRIDeal Application')}...")
    logger.info(f"Version: {app_config.get('APP_VERSION', 'N/A')}")

    # Resource checks - FIXED
    logger.info("Starting resource checks...")
    check_resources(app_config)  # Pass app_config to the function

    app_id_key = "APP_USER_MODEL_ID"
    app_name_for_id = app_config.get("APP_NAME", "BRIDeal").replace(" ", "")
    default_app_id = f"YourCompany.{app_name_for_id}.MainApp.1"
    app_id = app_config.get(app_id_key, default_app_id)
    set_app_user_model_id(app_id)

    app_auth_service = AppAuthService(config=app_config)
    logger.info(f"Current application user: {app_auth_service.get_current_user_name()}")

    cache_handler = CacheHandler(config=app_config)
    token_handler = TokenHandler(config=app_config, cache_handler=cache_handler)

    # --- SharePoint Service Instantiation with Debugging ---
    sharepoint_service_instance = None
    logger.info(f"DEBUG: Initial type of SharePointManagerService (before if): {type(SharePointManagerService)}")
    logger.info(f"DEBUG: SharePointManagerService is None (before if): {SharePointManagerService is None}")

    if SharePointManagerService: # This is the class if import was successful
        logger.info(f"DEBUG: SharePointManagerService IS NOT None. Type: {type(SharePointManagerService)}")
        try:
            logger.info("DEBUG: Attempting to instantiate SharePointManagerService with () call.")
            sharepoint_service_instance = SharePointManagerService() # CORRECTED: Instantiate the class
            logger.info(f"DEBUG: Type of sharepoint_service_instance after instantiation: {type(sharepoint_service_instance)}")
            logger.info(f"DEBUG: sharepoint_service_instance is SharePointManagerService class: {sharepoint_service_instance is SharePointManagerService}")

            if hasattr(sharepoint_service_instance, 'is_operational'):
                logger.info(f"DEBUG: sharepoint_service_instance.is_operational = {sharepoint_service_instance.is_operational}")
                if not sharepoint_service_instance.is_operational:
                    logger.warning(f"SharePointExcelManager initialized (instance: {type(sharepoint_service_instance)}) but is not operational.")
                else:
                    logger.info(f"SharePointExcelManager (instance: {type(sharepoint_service_instance)}) appears operational.")
            else: # This case should ideally not happen if __init__ defines is_operational
                logger.warning(f"DEBUG: sharepoint_service_instance LACKS is_operational attribute after init. Type: {type(sharepoint_service_instance)}")
        except Exception as sp_init_e:
            logger.error(f"DEBUG: Exception during SharePointManagerService instantiation: {sp_init_e}", exc_info=True)
            logger.info(f"DEBUG: Type of sharepoint_service_instance after EXCEPTION during init: {type(sharepoint_service_instance)}")
    else:
        logger.error("SharePointManagerService class not available (was None after import attempt).")
        logger.info(f"DEBUG: sharepoint_service_instance remains None because SharePointManagerService was None.")
    # --- End SharePoint Service Instantiation ---

    quote_builder = QuoteBuilder(config=app_config)
    jd_auth_manager = JDAuthManager(config=app_config, token_handler=token_handler)
    jd_quote_api_client = JDQuoteApiClient(config=app_config, auth_manager=jd_auth_manager)
    maintain_quotes_api = MaintainQuotesAPI(config=app_config, jd_quote_api_client=jd_quote_api_client)
    jd_quote_integration_service = JDQuoteIntegrationService(
        config=app_config,
        maintain_quotes_api=maintain_quotes_api,
        quote_builder=quote_builder
    )

    # Fix potential issues with the config
    try:
        from app.services.integrations.jd_auth_manager_improvements import check_and_fix_redirect_uri, ensure_auth_persistence

        # Fix potential issues with the config
        redirect_uri_fixed = check_and_fix_redirect_uri(app_config)
        persistence_enabled = ensure_auth_persistence(app_config)

        if redirect_uri_fixed or persistence_enabled:
            logger.info("Fixed John Deere API authentication configuration issues")

        # Check authentication status but don't prompt yet
        if jd_auth_manager and jd_auth_manager.is_operational:
            logger.info("Checking JD API authentication status on startup")
            token = jd_auth_manager.get_access_token()
            if not token:
                logger.info("No valid JD API token found at startup, authentication will be required")
    except Exception as auth_fix_error:
        logger.warning(f"Error while checking/fixing authentication configuration: {auth_fix_error}", exc_info=True)

    if hasattr(jd_quote_integration_service, 'is_operational') and not jd_quote_integration_service.is_operational:
        logger.warning("JDQuoteIntegrationService is not operational.")
    elif not hasattr(jd_quote_integration_service, 'is_operational'): # Check if attribute exists
        logger.info("JDQuoteIntegrationService status unknown (no 'is_operational' attribute).")
    else: # Attribute exists and is True
        logger.info("JDQuoteIntegrationService appears operational.")

    # Create QApplication instance
    qt_app = QApplication.instance()
    if qt_app is None:
        qt_app = QApplication(sys.argv)

    qt_app.setApplicationName(app_config.get("APP_NAME", "BRIDeal"))
    qt_app.setOrganizationName(app_config.get("ORGANIZATION_NAME", "YourCompany"))

    qapp_icon_config_key = "APP_ICON_PATH"
    qapp_default_icon_rel_to_res = "icons/app_icon.png"
    qapp_icon_path = app_config.get(qapp_icon_config_key, qapp_default_icon_rel_to_res)
    qapp_icon_abs_path = get_resource_path(qapp_icon_path, app_config)

    if qapp_icon_abs_path and os.path.exists(qapp_icon_abs_path):
        qt_app.setWindowIcon(QIcon(qapp_icon_abs_path))
        logger.info(f"QApplication icon set from: {qapp_icon_abs_path}")
    else:
        logger.warning(
            f"QApplication icon not found. Config key: '{qapp_icon_config_key}', "
            f"Value from config/default: '{qapp_icon_path}', "
            f"Resolved absolute path: {qapp_icon_abs_path}"
        )

    splash_config_key = "SPLASH_IMAGE_MAIN"
    splash_default_rel_to_res = "images/splash_main.png"
    splash_image_rel_path = app_config.get(splash_config_key, splash_default_rel_to_res)
    splash_image_abs_path = get_resource_path(splash_image_rel_path, app_config)

    if not (splash_image_abs_path and os.path.exists(splash_image_abs_path)):
        logger.warning(
            f"Splash image not found. Config key: '{splash_config_key}', "
            f"Value from config/default: '{splash_image_rel_path}', "
            f"Resolved absolute path: {splash_image_abs_path}"
        )
        splash_image_abs_path = None # Use None if not found

    splash_screen = SplashScreenView(
        image_path=splash_image_abs_path,
        app_name=app_config.get("APP_NAME", "BRIDeal"),
        version_text=f"v{app_config.get('APP_VERSION', 'N/A')}",
        config=app_config # Pass config for splash screen to use if needed
    )
    splash_screen.show_message("Initializing application core...")
    qt_app.processEvents() # Allow UI to update

    try:
        splash_screen.show_message("Loading main interface...")
        qt_app.processEvents()

        # Correctly pass the instantiated sharepoint_service_instance (or None)
        logger.info(f"DEBUG: Final type of sharepoint_service_instance being passed to MainWindow: {type(sharepoint_service_instance)}")
        main_window = MainWindow(
            config=app_config,
            cache_handler=cache_handler,
            token_handler=token_handler,
            sharepoint_manager=sharepoint_service_instance, # Use the (potentially None) instance
            jd_auth_manager=jd_auth_manager,
            jd_quote_integration_service=jd_quote_integration_service
        )

        main_window.show()
        splash_screen.finish(main_window) # Pass main_window to splash screen
    except Exception as e:
        logger.critical(f"Failed to initialize and show MainWindow: {e}", exc_info=True)
        splash_screen.close() # Ensure splash is closed on error
        QMessageBox.critical(None, "Application Initialization Error",
                                f"A critical error occurred while loading the main application window:\n{e}\n\n"
                                "The application will now exit. Please check the logs for more details.")
        sys.exit(1) # Exit if main window fails

    sys.exit(qt_app.exec_())


if __name__ == '__main__':
    # BasicConfig for pre-config logging. setup_logging will reconfigure.
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - PRE_CONFIG - %(message)s')
    try:
        run_application()
    except Exception as e:
        # This is a last resort catch for unhandled exceptions during launch.
        # It's better if errors are caught within run_application or by Qt's hook.
        critical_logger = logging.getLogger("critical_launch_error") # A specific logger
        critical_logger.critical(f"Unhandled exception during application launch: {e}", exc_info=True)
        try:
            # Attempt to show a GUI message box even in critical failure
            app_temp = QApplication.instance() # Check if an app instance exists
            if not app_temp: 
                app_temp = QApplication(sys.argv) # Create if not
            QMessageBox.critical(None, "Critical Application Failure",
                                    f"A critical unhandled error occurred that prevented the application from starting properly:\n{e}\n\n"
                                    "Please check logs or console output for details.")
        except Exception as qmb_error:
            # If even the QMessageBox fails (e.g., display not available), print to stderr
            print(f"CRITICAL LAUNCH ERROR: {e}", file=sys.stderr)
            print(f"GUI DIALOG FOR CRITICAL ERROR FAILED: {qmb_error}", file=sys.stderr)
        sys.exit(1)