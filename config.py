import os
from dotenv import load_dotenv
from typing import Dict, Any

# Load .env
dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
print(f"Attempting to load .env file from: {dotenv_path}")
loaded = load_dotenv(dotenv_path=dotenv_path, verbose=True)

if not loaded:
    print(f"Warning: .env file not found or failed to load from {dotenv_path}")
else:
    print(".env file loaded successfully.")

# Load again to ensure environment is populated
load_dotenv()

# --- Config Classes ---

class APIConfig:
    OPENWEATHER = {
        "API_KEY": os.getenv("OPENWEATHER_API_KEY", "711ac00142aa78e1807ce84a8bf1582b"),
        "BASE_URL": os.getenv("OPENWEATHER_BASE_URL", "https://api.openweathermap.org/data/2.5/weather")
    }
    ALPHAVANTAGE = {
        "API_KEY": os.getenv("ALPHAVANTAGE_API_KEY", "PHNW69I8KX24I5PT"),
        "BASE_URL": os.getenv("ALPHAVANTAGE_BASE_URL", "https://www.alphavantage.co/query")
    }
    COMMODITIES = {
        "WHEAT_SYMBOL": "ZW=F",
        "CANOLA_SYMBOL": "RS=F"
    }

class AzureConfig:
    CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
    CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
    TENANT_ID = os.getenv("AZURE_TENANT_ID")
    SCOPES = ["https://graph.microsoft.com/.default"]
    AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

class SharePointConfig:
    SITE_ID = os.getenv("SHAREPOINT_SITE_ID")
    SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME")
    BASE_URL = f"https://briltd.sharepoint.com/sites/{os.getenv('SHAREPOINT_SITE_NAME')}"
    FILE_PATH = os.getenv("SHAREPOINT_FILE_PATH")
    DRIVE_NAME = "Documents"
    SHEET_NAME = "App"

class PathConfig:
    DATA_DIR = "data"
    PRODUCTS_CSV = os.path.join(DATA_DIR, "products.csv")
    CUSTOMERS_CSV = os.path.join(DATA_DIR, "customers.csv")
    SALESMEN_CSV = os.path.join(DATA_DIR, "salesmen.csv")
    PARTS_CSV = os.path.join(DATA_DIR, "parts.csv")
    DEAL_FORM_CSV = "deal_form.csv"
    STOCKS_CSV = "stocks.csv"
    LOG_FILE = "amsdeal.log"
    ICON_PATH = os.path.join("resources", "BRIapp.ico")

class AppSettings:
    WEATHER_REFRESH_INTERVAL = int(os.getenv("WEATHER_REFRESH_INTERVAL", 900000))  # 15 minutes
    STOCK_REFRESH_INTERVAL = int(os.getenv("STOCK_REFRESH_INTERVAL", 300000))      # 5 minutes
    API_TIMEOUT = 10
    AUTO_REFRESH_INTERVAL = 60000
    UI_IMAGE_SEARCH_CONFIDENCE = 0.8
    CHART_STYLE = {
        "figure.facecolor": "#f8f9fa",
        "axes.facecolor": "#f0f4f8",
        "axes.labelcolor": "#2a5d24",
        "axes.titlesize": 16,
        "axes.titleweight": "bold",
        "axes.titlecolor": "#2a5d24",
        "axes.grid": True,
        "grid.linestyle": "--",
        "grid.alpha": 0.7,
        "grid.color": "#d1d5db",
        "xtick.color": "#555",
        "ytick.color": "#555"
    }

# --- Validation ---

class Validation:
    @staticmethod
    def check_required_env_vars() -> None:
        required_vars = {
            "OPENWEATHER_API_KEY": APIConfig.OPENWEATHER["API_KEY"],
            "ALPHAVANTAGE_API_KEY": APIConfig.ALPHAVANTAGE["API_KEY"],
            "AZURE_CLIENT_ID": AzureConfig.CLIENT_ID,
            "AZURE_CLIENT_SECRET": AzureConfig.CLIENT_SECRET,
            "AZURE_TENANT_ID": AzureConfig.TENANT_ID,
            "SHAREPOINT_SITE_ID": SharePointConfig.SITE_ID
        }
        missing = [name for name, value in required_vars.items() if not value]
        if missing:
            raise EnvironmentError(f"Missing required environment variables: {', '.join(missing)}")

# Run validation
Validation.check_required_env_vars()
