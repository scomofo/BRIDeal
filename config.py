import os
from dotenv import load_dotenv
# Removed unused imports
import logging

# Set up basic logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# --- Load .env File ---
script_dir = os.path.dirname(os.path.abspath(__file__))
dotenv_path = os.path.join(os.path.dirname(script_dir), '.env') # Assume .env in project root (parent of config.py dir)

logger.info(f"Attempting to load .env file from: {dotenv_path}")
loaded = load_dotenv(dotenv_path=dotenv_path, verbose=True, override=True)

if not loaded:
    logger.warning(f".env file not found or failed to load from {dotenv_path}. Trying default load.")
    loaded = load_dotenv(verbose=True, override=True)
    if not loaded:
        logger.warning("Default .env load also failed. Environment variables might not be set.")
    else:
        logger.info("Default .env loaded successfully (using default search path).")
else:
    logger.info(f".env file loaded successfully from {dotenv_path}.")


# --- Configuration Classes ---

class APIConfig:
    """Holds configuration for external APIs."""
    # --- REMOVED ---
    # OPENWEATHER_API_KEY = os.getenv("OPENWEATHER_API_KEY")
    # ALPHAVANTAGE_API_KEY = os.getenv("ALPHAVANTAGE_API_KEY")
    # OPENWEATHER_BASE_URL = "https://api.openweathermap.org/data/2.5/weather"
    # ALPHAVANTAGE_BASE_URL = "https://www.alphavantage.co/query"
    # --- ADDED ---
    FINNHUB_API_KEY = os.getenv("FINNHUB_API_KEY") # Key for Finnhub

    # --- KEPT ---
    MS_GRAPH_BASE_URL = os.getenv('MS_GRAPH_BASE_URL', 'https://graph.microsoft.com/v1.0')

    # Commodity symbols - update with appropriate symbols for Finnhub/YFinance if needed
    # Using example Finnhub/common symbols. Adjust if needed based on chosen provider.
    COMMODITIES = {
        "WHEAT_SYMBOL": os.getenv("WHEAT_SYMBOL", "CBOT:ZW1!"), # Example: CBOT Wheat Front Month
        "BITCOIN_SYMBOL": os.getenv("BITCOIN_SYMBOL", "BINANCE:BTCUSDT"), # Example: Binance BTC/USDT
        "CANOLA_SYMBOL": os.getenv("CANOLA_SYMBOL", "ICECANOLA:RS1!"), # Example: ICE Canola Front Month
        "FX_SYMBOL_USDCAD": os.getenv("FX_SYMBOL_USDCAD", "OANDA:USD_CAD") # Example: OANDA USD/CAD
    }

class AzureConfig:
    """Holds configuration for Azure Active Directory authentication."""
    CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
    CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
    TENANT_ID = os.getenv("AZURE_TENANT_ID")
    AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}" if TENANT_ID else None
    SCOPES = [f"{APIConfig.MS_GRAPH_BASE_URL}/.default"]

class SharePointConfig:
    """Holds configuration for SharePoint access."""
    SITE_ID = os.getenv("SHAREPOINT_SITE_ID")
    SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME")
    BASE_URL = f"https://briltd.sharepoint.com/sites/{SITE_NAME}" if SITE_NAME else None
    FILE_PATH = os.getenv("FILE_PATH")
    DRIVE_NAME = os.getenv("SHAREPOINT_DRIVE_NAME", "Documents")
    SHEET_NAME = os.getenv("SHAREPOINT_SHEET_NAME", "App")
    SENDER_EMAIL = os.getenv("SHAREPOINT_SENDER_EMAIL")

class PathConfig:
    """Manages application file paths."""
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    # Define directories relative to BASE_DIR, allowing overrides from .env
    DATA_DIR = os.path.join(BASE_DIR, os.getenv("DATA_DIR_PATH", "data"))
    LOGS_DIR = os.path.join(BASE_DIR, os.getenv("LOGS_DIR_PATH", "logs"))
    ASSETS_DIR = os.path.join(BASE_DIR, os.getenv("ASSETS_DIR_PATH", "assets"))
    CACHE_DIR = os.path.join(BASE_DIR, os.getenv("CACHE_DIR_PATH", "cache"))
    RESOURCES_DIR = os.path.join(BASE_DIR, os.getenv("RESOURCES_DIR_PATH", "resources"))
    EXPORTS_DIR = os.path.join(BASE_DIR, os.getenv("EXPORTS_DIR_PATH", "modules/exports"))

    # Construct full paths using the directory variables and env overrides/defaults
    PRODUCTS_CSV = os.path.join(DATA_DIR, os.getenv("PRODUCTS_CSV", "products.csv"))
    CUSTOMERS_CSV = os.path.join(DATA_DIR, os.getenv("CUSTOMERS_CSV", "customers.csv"))
    SALESMEN_CSV = os.path.join(DATA_DIR, os.getenv("SALESMEN_CSV", "salesmen.csv"))
    PARTS_CSV = os.path.join(DATA_DIR, os.getenv("PARTS_CSV", "parts.csv"))
    DEAL_FORM_CSV = os.path.join(DATA_DIR, os.getenv("DEAL_FORM_CSV", "deal_form.csv"))
    STOCKS_CSV = os.path.join(DATA_DIR, os.getenv("STOCKS_CSV", "stocks.csv"))
    LOG_FILE = os.path.join(LOGS_DIR, os.getenv("LOG_FILE", "amsdeal.log"))
    ICON_PATH = os.path.join(RESOURCES_DIR, os.getenv("ICON_FILE", "BRIapp.ico"))

    DEAL_DRAFT_JSON = os.path.join(DATA_DIR, os.getenv('DEAL_DRAFT_JSON', 'deal_draft.json'))
    RECENT_DEALS_JSON = os.path.join(DATA_DIR, os.getenv('RECENT_DEALS_JSON', 'recent_deals.json'))
    WEATHER_CACHE_JSON = os.path.join(CACHE_DIR, os.getenv('WEATHER_CACHE_JSON', 'weather_cache.json'))
    EXCHANGE_CACHE_JSON = os.path.join(CACHE_DIR, os.getenv('EXCHANGE_CACHE_JSON', 'exchange_cache.json'))
    COMMODITIES_CACHE_JSON = os.path.join(CACHE_DIR, os.getenv('COMMODITIES_CACHE_JSON', 'commodities_cache.json'))
    PRICEBOOK_SETTINGS_JSON = os.path.join(DATA_DIR, os.getenv('PRICEBOOK_SETTINGS_JSON', 'pricebook_settings.json'))

    @staticmethod
    def ensure_directories_exist():
        """Creates required directories if they don't exist."""
        dirs_to_create = [
            PathConfig.DATA_DIR, PathConfig.LOGS_DIR, PathConfig.ASSETS_DIR,
            PathConfig.CACHE_DIR, PathConfig.RESOURCES_DIR, PathConfig.EXPORTS_DIR
        ]
        for dir_path in dirs_to_create:
            try:
                abs_dir_path = os.path.abspath(dir_path)
                os.makedirs(abs_dir_path, exist_ok=True)
                logger.info(f"Ensured directory exists: {abs_dir_path}")
            except OSError as e:
                logger.error(f"Failed to create directory {abs_dir_path}: {e}", exc_info=True)

class AppSettings:
    """Holds general application settings."""
    WEATHER_REFRESH_INTERVAL = int(os.getenv("WEATHER_REFRESH_INTERVAL_HRS", 1)) # Refresh interval in HOURS
    EXCHANGE_REFRESH_INTERVAL = int(os.getenv("EXCHANGE_REFRESH_INTERVAL_HRS", 6)) # Refresh interval in HOURS
    COMMODITIES_REFRESH_INTERVAL = int(os.getenv("COMMODITIES_REFRESH_INTERVAL_HRS", 4)) # Refresh interval in HOURS
    API_TIMEOUT = int(os.getenv("API_TIMEOUT_SECS", 15)) # API timeout in SECONDS
    AUTO_REFRESH_INTERVAL = int(os.getenv("AUTO_REFRESH_INTERVAL", 60000)) # General UI refresh (unused?)
    UI_IMAGE_SEARCH_CONFIDENCE = float(os.getenv("UI_IMAGE_SEARCH_CONFIDENCE", 0.8))
    APP_NAME = os.getenv('APP_NAME', 'AMS Deal App')

    CHART_STYLE = { # Keep styling
        "figure.facecolor": "#f8f9fa", "axes.facecolor": "#f0f4f8",
        "axes.labelcolor": "#2a5d24", "axes.titlesize": 16,
        "axes.titleweight": "bold", "axes.titlecolor": "#2a5d24",
        "axes.grid": True, "grid.linestyle": "--", "grid.alpha": 0.7,
        "grid.color": "#d1d5db", "xtick.color": "#555", "ytick.color": "#555"
    }


# --- Validation ---

class Validation:
    """Performs validation checks on loaded configuration."""
    @staticmethod
    def check_required_env_vars() -> bool:
        """Checks if essential environment variables are loaded."""
        required_vars = {
            # --- REMOVED CHECKS ---
            # "OPENWEATHER_API_KEY": APIConfig.OPENWEATHER_API_KEY,
            # "ALPHAVANTAGE_API_KEY": APIConfig.ALPHAVANTAGE_API_KEY,
            # --- ADDED CHECK ---
            "FINNHUB_API_KEY": APIConfig.FINNHUB_API_KEY,
            # --- KEPT CHECKS ---
            "AZURE_CLIENT_ID": AzureConfig.CLIENT_ID,
            "AZURE_CLIENT_SECRET": AzureConfig.CLIENT_SECRET,
            "AZURE_TENANT_ID": AzureConfig.TENANT_ID,
            "SHAREPOINT_SITE_ID": SharePointConfig.SITE_ID,
        }
        missing = [name for name, value in required_vars.items() if not value]
        if missing:
            logger.warning(f"Missing required environment variables: {', '.join(missing)}")
            logger.warning("Application functionality may be limited or fail.")
            return False
        logger.info("All checked required environment variables are present.")
        return True


# --- Initialization ---
apis = APIConfig()
azure = AzureConfig()
sharepoint = SharePointConfig()
paths = PathConfig()
settings = AppSettings()

try:
    PathConfig.ensure_directories_exist()
except Exception as e:
    logger.error(f"Critical error during directory creation: {e}", exc_info=True)


# --- Masking Function for Sensitive Data ---
def mask_sensitive(value: str, visible_chars: int = 4) -> str:
    """Masks a string, showing only the last few characters."""
    if not isinstance(value, str) or len(value) <= visible_chars:
        return "****"
    return '*' * (len(value) - visible_chars) + value[-visible_chars:]


# --- Log Configuration Summary (Optional) ---
logger.info("--- Configuration Summary ---")
logger.info(f"Base Directory: {os.path.abspath(paths.BASE_DIR)}")
logger.info(f"Data Directory: {os.path.abspath(paths.DATA_DIR)}")
logger.info(f"Log File Path: {os.path.abspath(paths.LOG_FILE)}")
logger.info(f"Cache Directory: {os.path.abspath(paths.CACHE_DIR)}")
# --- REMOVED LOGS ---
# logger.info(f"OpenWeather API Key (from env): {mask_sensitive(apis.OPENWEATHER_API_KEY)}")
# logger.info(f"AlphaVantage API Key (from env): {mask_sensitive(apis.ALPHAVANTAGE_API_KEY)}")
# --- ADDED LOG ---
logger.info(f"Finnhub API Key (from env): {mask_sensitive(apis.FINNHUB_API_KEY)}")
# --- KEPT LOGS ---
logger.info(f"Azure Client ID (from env): {azure.CLIENT_ID}")
logger.info(f"Azure Client Secret (from env): {mask_sensitive(azure.CLIENT_SECRET)}")
logger.info(f"Azure Tenant ID (from env): {azure.TENANT_ID}")
logger.info(f"SharePoint Site ID (from env): {sharepoint.SITE_ID}")
logger.info(f"SharePoint Site Name (from env): {sharepoint.SITE_NAME}")
logger.info(f"SharePoint File Path (from env): {sharepoint.FILE_PATH}")
logger.info(f"Weather Refresh Interval (Hours): {settings.WEATHER_REFRESH_INTERVAL}") # Log hours
logger.info(f"Exchange Refresh Interval (Hours): {settings.EXCHANGE_REFRESH_INTERVAL}") # Log hours
logger.info(f"Commodities Refresh Interval (Hours): {settings.COMMODITIES_REFRESH_INTERVAL}") # Log hours
logger.info(f"API Timeout (Seconds): {settings.API_TIMEOUT}") # Log seconds
logger.info("--- End Configuration Summary ---")

# Run validation check
validation_passed = Validation.check_required_env_vars()
# (Optional: Add handling if validation fails, e.g., sys.exit)

# Make configuration objects easily importable
__all__ = ['apis', 'azure', 'sharepoint', 'paths', 'settings']