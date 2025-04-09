import csv
import os
import logging
import json
import time
import requests # Kept in case other functions use it
import sys # Added for get_resource_path
from typing import Optional, Tuple, Dict, Any, List

# Import config objects
# Ensure config.py is in the python path or adjust import as needed
try:
    # Note: config.py itself has been modified, ensure imports reflect that if needed
    from config import apis, paths, settings
except ImportError:
    logging.critical("Failed to import configuration from config.py. Utils functions will likely fail.")
    # Define dummy objects to prevent NameErrors, but functions will not work
    class DummyConfig:
        # Removed API keys/URLs related to removed services
        COMMODITIES = {}
        CACHE_JSON = "dummy_cache.json" # Keep generic cache names if needed elsewhere
        API_TIMEOUT = 25
        # Define dummy paths if needed by remaining utils functions
        DATA_DIR = "data"
        LOGS_DIR = "logs"
        ASSETS_DIR = "assets"
        CACHE_DIR = "cache"
        RESOURCES_DIR = "resources"
        EXPORTS_DIR = "exports"
        PRODUCTS_CSV = os.path.join(DATA_DIR, "products.csv")
        # ... add other paths if used by remaining functions ...
        LOG_FILE = os.path.join(LOGS_DIR, "amsdeal.log")
        ICON_PATH = os.path.join(RESOURCES_DIR, "BRIapp.ico")
        WEATHER_CACHE_JSON = os.path.join(CACHE_DIR, 'weather_cache.json')
        EXCHANGE_CACHE_JSON = os.path.join(CACHE_DIR, 'exchange_cache.json')
        COMMODITIES_CACHE_JSON = os.path.join(CACHE_DIR, 'commodities_cache.json')


    apis = DummyConfig()
    paths = DummyConfig()
    settings = DummyConfig()


# Get a logger for this module
logger = logging.getLogger(__name__)
# Ensure logging is configured in your main application entry point (e.g., main.py)
# Example basic config (if not done elsewhere):
# logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')


# ==============================================
# Caching Utilities
# ==============================================

def save_cache(cache_path: str, data: Any) -> None:
    """Saves data to a JSON cache file."""
    try:
        # Ensure cache directory exists (config.py should handle this, but double-check)
        cache_dir = os.path.dirname(cache_path)
        if not os.path.exists(cache_dir):
             os.makedirs(cache_dir, exist_ok=True)
             logger.info(f"Created cache directory: {cache_dir}")

        # Add timestamp to the cached data
        cache_content = {
            'timestamp': time.time(),
            'data': data
        }
        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump(cache_content, f, indent=4)
        logger.info(f"Data successfully saved to cache: {cache_path}")
    except IOError as e:
        logger.error(f"Failed to save cache file {cache_path}: {e}", exc_info=True)
    except TypeError as e:
         logger.error(f"Failed to serialize data for cache file {cache_path}: {e}", exc_info=True)
    except Exception as e:
        logger.error(f"An unexpected error occurred saving cache to {cache_path}: {e}", exc_info=True)


def load_cache(cache_path: str) -> Tuple[Optional[Any], float]:
    """
    Loads data from a JSON cache file.

    Returns:
        Tuple[Optional[Any], float]: A tuple containing the cached data (or None if not found/invalid)
                                     and the timestamp of the cache (or 0 if not found/invalid).
    """
    if not os.path.exists(cache_path):
        logger.info(f"Cache file not found: {cache_path}")
        return None, 0

    try:
        with open(cache_path, 'r', encoding='utf-8') as f:
            cache_content = json.load(f)
        timestamp = cache_content.get('timestamp', 0)
        data = cache_content.get('data')
        logger.info(f"Cache loaded successfully from: {cache_path}")
        return data, timestamp
    except (IOError, json.JSONDecodeError, KeyError) as e:
        logger.error(f"Failed to load or parse cache file {cache_path}: {e}", exc_info=True)
        # Optionally delete corrupted cache file
        # try:
        #     os.remove(cache_path)
        #     logger.warning(f"Removed corrupted cache file: {cache_path}")
        # except OSError as remove_err:
        #     logger.error(f"Failed to remove corrupted cache file {cache_path}: {remove_err}")
        return None, 0
    except Exception as e:
         logger.error(f"An unexpected error occurred loading cache from {cache_path}: {e}", exc_info=True)
         return None, 0


# ==============================================
# API Fetching Utilities (REMOVED)
# ==============================================

# fetch_weather_data function removed

# fetch_exchange_rates function removed

# fetch_commodity_data function removed


# ==============================================
# Formatting and Path Utilities
# ==============================================

def format_currency(value: Any, default: str = "N/A") -> str:
    """Formats a numeric value as currency (e.g., $1,234.56)."""
    try:
        # Attempt to convert to float if it's a string representation of a number
        if isinstance(value, str):
            value = float(value.replace(',', '')) # Handle potential commas
        if isinstance(value, (int, float)):
            return f"${value:,.2f}"
        else:
            logger.debug(f"format_currency: Input value '{value}' is not numeric.")
            return default
    except (ValueError, TypeError) as e:
        logger.warning(f"Could not format value '{value}' as currency: {e}")
        return default

def get_local_time_from_utc(utc_timestamp: float, time_format: str = '%Y-%m-%d %H:%M') -> str:
    """Converts a UTC timestamp (seconds since epoch) to a local time string."""
    if not utc_timestamp:
        return "N/A"
    try:
        local_time = time.localtime(utc_timestamp)
        return time.strftime(time_format, local_time)
    except Exception as e:
        logger.error(f"Failed to convert UTC timestamp {utc_timestamp} to local time: {e}")
        return "Invalid Date"

def get_resource_path(relative_path: str) -> str:
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        # Using getattr to safely check for _MEIPASS attribute
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        if base_path == os.path.dirname(os.path.abspath(__file__)):
             logger.debug(f"Running from source, base path: {base_path}")
             # Optional: Adjust if utils.py is in a subfolder relative to project root
             # base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        else:
             logger.debug(f"Running in PyInstaller bundle, base path: {base_path}")

    except Exception as e:
         logger.error(f"Error determining base path in get_resource_path: {e}. Falling back to script directory.")
         # Fallback to the directory of the current file if getattr fails unexpectedly
         base_path = os.path.dirname(os.path.abspath(__file__))


    # Construct path relative to the determined base path
    resource_path = os.path.join(base_path, relative_path)
    logger.debug(f"Resolved resource path for '{relative_path}' to: {resource_path}")
    return resource_path


# ==============================================
# CSV Loading Utility (Corrected Version)
# ==============================================

class CSVLoader:
    """
    A utility class to load data from CSV files with robust error handling
    and encoding detection.
    """
    def __init__(self, full_file_path: str):
        """
        Initializes the CSVLoader with the full path to the CSV file.

        Args:
            full_file_path (str): The absolute path to the CSV file,
                                  preferably obtained from config.paths.
        """
        # It's generally better if the caller ensures the path is absolute
        # if not os.path.isabs(full_file_path):
        #      logger.warning(f"CSVLoader initialized with non-absolute path: {full_file_path}. This might lead to issues.")
        self.file_path = full_file_path
        self.filename = os.path.basename(full_file_path)
        logger.debug(f"CSVLoader initialized for file: {self.file_path}")

    def load(self, skip_header: bool = False) -> list[list[str]]:
        """
        Load CSV data into a list of rows.

        Args:
            skip_header (bool): If True, the first row (header) will be skipped.
                                Defaults to False.

        Returns:
            list[list[str]]: A list of lists, where each inner list represents a row.
                             Returns an empty list if the file is not found or cannot be read.
        """
        data = []
        encodings = ['utf-8', 'latin1', 'windows-1252']
        file_found = False

        for encoding in encodings:
            try:
                logger.debug(f"Attempting to read {self.filename} with encoding '{encoding}'...")
                with open(self.file_path, mode='r', newline='', encoding=encoding) as csvfile:
                    reader = csv.reader(csvfile)
                    file_found = True
                    if skip_header:
                        try:
                            next(reader)
                            logger.debug("Skipped header row.")
                        except StopIteration:
                            logger.warning(f"CSV file '{self.filename}' seems empty or has no header to skip.")
                            return []

                    row_count = 0
                    for row in reader:
                        if row: # Only append non-empty rows
                            data.append(row)
                            row_count += 1
                    logger.info(f"Successfully loaded {row_count} rows from '{self.filename}' using encoding '{encoding}'.")
                    return data

            except UnicodeDecodeError:
                logger.debug(f"Encoding '{encoding}' failed for '{self.filename}'. Trying next.")
                continue
            except FileNotFoundError:
                logger.error(f"CSV file not found at path: {self.file_path}")
                return []
            except Exception as e:
                logger.error(f"An unexpected error occurred while reading '{self.filename}' with encoding '{encoding}': {e}", exc_info=True)
                return [] # Stop trying encodings if another error occurs

        if file_found:
             logger.error(f"Could not decode '{self.filename}' with any attempted encodings: {encodings}")
        # If file wasn't found in the loop, the error was already logged.

        return []

    def load_dict(self, key_column: str, value_column: str = None) -> dict:
        """
        Load CSV into a dictionary with a specified key column.
        The dictionary values can be the entire row or a specific value column.

        Args:
            key_column (str): The name of the header column to use as dictionary keys.
            value_column (str, optional): The name of the header column to use as
                                          dictionary values. If None, the entire row
                                          (as a list) will be the value. Defaults to None.

        Returns:
            dict: A dictionary mapping keys from key_column to corresponding values.
                  Returns an empty dictionary if loading fails, file is empty,
                  or required columns are not found.
        """
        data = self.load(skip_header=False) # Load including header to find indices

        if not data or len(data) < 2: # Need at least header and one data row
            logger.warning(f"Cannot load '{self.filename}' into dict: File is empty or contains only a header.")
            return {}

        headers = [h.strip() for h in data[0]] # Trim whitespace from headers
        logger.debug(f"CSV Headers found for '{self.filename}': {headers}")

        try:
            key_index = headers.index(key_column)
            logger.debug(f"Found key column '{key_column}' at index {key_index}.")
            value_index = -1 # Initialize value_index
            if value_column:
                value_index = headers.index(value_column)
                logger.debug(f"Found value column '{value_column}' at index {value_index}.")

            result_dict = {}
            skipped_rows = 0
            for i, row in enumerate(data[1:], start=1): # Start from the first data row
                try:
                    # Ensure row has enough columns before accessing indices
                    if key_index >= len(row) or (value_column and value_index >= len(row)):
                         raise IndexError(f"Row {i+1} has {len(row)} cells, needs at least {max(key_index, value_index if value_column else key_index) + 1}.")

                    key = row[key_index].strip() # Trim whitespace from key
                    if not key: # Skip rows with empty keys
                         logger.warning(f"Skipping row {i+1} in '{self.filename}' due to empty key in column '{key_column}'.")
                         skipped_rows += 1
                         continue

                    if value_column is not None:
                        # Trim whitespace from value if it's a specific column
                        value = row[value_index].strip()
                    else:
                        # Trim whitespace from all cells if the value is the whole row
                        value = [cell.strip() for cell in row]

                    if key in result_dict:
                        logger.warning(f"Duplicate key '{key}' found in '{self.filename}' at row {i+1}. Previous value will be overwritten.")
                    result_dict[key] = value

                except IndexError as e:
                    logger.warning(f"Skipping malformed row {i+1} in '{self.filename}'. Error: {e}. Row content: {row}")
                    skipped_rows += 1
                except Exception as e: # Catch other potential errors during row processing
                    logger.error(f"Error processing row {i+1} in '{self.filename}': {e}. Row: {row}", exc_info=True)
                    skipped_rows += 1

            loaded_count = len(result_dict)
            logger.info(f"Successfully loaded {loaded_count} items into dict from '{self.filename}'. Skipped {skipped_rows} rows due to errors or empty keys.")
            return result_dict

        except ValueError: # Handles headers.index() failure
            missing_col = key_column if key_column not in headers else value_column
            logger.error(f"Required column '{missing_col}' not found in '{self.filename}'. Available headers: {headers}")
            return {}
        except Exception as e: # Catch unexpected errors during setup
             logger.error(f"An unexpected error occurred during dictionary creation for '{self.filename}': {e}", exc_info=True)
             return {}


# ==============================================
# Main execution block for testing (optional)
# ==============================================
if __name__ == '__main__':
    # This block runs only when utils.py is executed directly
    # Configure logging for direct testing
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    # --- Test API Functions (REMOVED) ---
    # print("\n--- Testing API Fetches ---")
    # Load .env if running directly
    # from dotenv import load_dotenv
    # Assume .env is in the parent directory of utils.py (project root)
    # dotenv_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '.env')
    # if os.path.exists(dotenv_path):
    #      load_dotenv(dotenv_path=dotenv_path, verbose=True)
    #      logger.info(f"Loaded .env for direct testing from {dotenv_path}")
         # Re-import config after loading .env if necessary, or ensure config loads it
    #      from config import apis as test_apis, paths as test_paths, settings as test_settings
    # else:
    #      logger.warning(f".env file not found at {dotenv_path}. API tests might fail if keys not in environment.")
         # Use potentially dummy config objects loaded earlier
    #      test_apis, test_paths, test_settings = apis, paths, settings

    # Test Weather (REMOVED)
    # print("\nTesting Weather Fetch...")
    # weather_result = fetch_weather_data(test_apis.OPENWEATHER_API_KEY, cache_path=test_paths.WEATHER_CACHE_JSON)
    # print(f"Weather Result: {weather_result}")

    # Test Exchange Rate (REMOVED)
    # print("\nTesting Exchange Rate Fetch...")
    # fx_result = fetch_exchange_rates(test_apis.ALPHAVANTAGE_API_KEY, cache_path=test_paths.EXCHANGE_CACHE_JSON)
    # print(f"Exchange Rate Result: {fx_result}")

    # Test Commodities (REMOVED)
    # print("\nTesting Commodity Fetch...")
    # comm_symbols = [s for s in [test_apis.COMMODITIES.get("WHEAT_SYMBOL"), test_apis.COMMODITIES.get("CANOLA_SYMBOL")] if s]
    # if comm_symbols:
    #     comm_result = fetch_commodity_data(test_apis.ALPHAVANTAGE_API_KEY, symbols=comm_symbols, cache_path=test_paths.COMMODITIES_CACHE_JSON)
    #     print(f"Commodities Result: {comm_result}")
    # else:
    #     print("No commodity symbols configured for testing.")


    # --- Test CSV Loader ---
    print("\n--- Testing CSVLoader ---")
    # Use paths potentially defined in the dummy config or loaded from .env
    try:
         test_csv_dir = os.path.dirname(paths.PRODUCTS_CSV) # Assuming CSVs are in DATA_DIR
         customers_csv_path = paths.CUSTOMERS_CSV # Use path from config
         products_csv_path = paths.PRODUCTS_CSV   # Use path from config
         logger.info(f"Using CSV paths from config: {customers_csv_path}, {products_csv_path}")
    except AttributeError:
         logger.error("Could not determine test CSV paths from config. Falling back to relative paths.")
         # Fallback if config/paths object is incomplete
         test_csv_dir = os.path.dirname(__file__) # Test files relative to utils.py
         customers_csv_path = os.path.join(test_csv_dir, 'test_customers.csv')
         products_csv_path = os.path.join(test_csv_dir, 'test_products.csv')


    # Ensure the directory exists before creating files
    os.makedirs(test_csv_dir, exist_ok=True)

    # Create dummy CSV files for testing if they don't exist
    if not os.path.exists(customers_csv_path):
        with open(customers_csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['CustomerID', 'Name', 'City'])
            writer.writerow(['CUST001', 'Alice', 'Calgary'])
            writer.writerow(['CUST002', 'Bob', 'Edmonton'])
            writer.writerow(['CUST003', 'Charlie', 'Camrose']) # Test duplicate key later
            writer.writerow(['CUST003', 'Chuck', 'Red Deer']) # Duplicate key
            writer.writerow(['CUST004', 'David']) # Malformed row (too short for dict test)
            writer.writerow(['', 'EmptyKey', 'Nowhere']) # Empty key test
        logger.info(f"Created dummy file: {customers_csv_path}")

    if not os.path.exists(products_csv_path):
         with open(products_csv_path, 'w', newline='', encoding='utf-8') as f:
             writer = csv.writer(f)
             writer.writerow([' ProductCode ', ' ProductName ', ' Price ']) # Headers with spaces
             writer.writerow([' P01 ', ' Widget ', ' 10.99 '])
             writer.writerow([' P02 ', ' Gadget ', ' 25.50 '])
         logger.info(f"Created dummy file: {products_csv_path}")


    print(f"\n--- Testing CSVLoader with: {customers_csv_path} ---")
    customer_loader = CSVLoader(customers_csv_path)
    customer_list = customer_loader.load(skip_header=True)
    print(f"Loaded list ({len(customer_list)} items):\n{customer_list[:2]}...") # Print first 2 items

    customer_dict = customer_loader.load_dict(key_column='CustomerID') # Assuming 'CustomerID' header exists
    print(f"\nLoaded dict by CustomerID ({len(customer_dict)} items):\n{list(customer_dict.items())[:2]}...")

    print(f"\n--- Testing CSVLoader with: {products_csv_path} ---")
    product_loader = CSVLoader(products_csv_path)
    # Assuming 'ProductCode' and 'ProductName' headers exist (after stripping)
    product_name_dict = product_loader.load_dict(key_column='ProductCode', value_column='ProductName')
    print(f"\nLoaded dict ProductCode->ProductName ({len(product_name_dict)} items):\n{list(product_name_dict.items())[:2]}...")
    # Test loading whole row as value
    product_row_dict = product_loader.load_dict(key_column='ProductCode')
    print(f"\nLoaded dict ProductCode->Row ({len(product_row_dict)} items):\n{list(product_row_dict.items())[:2]}...")


    # Test file not found
    print(f"\n--- Testing non-existent file ---")
    missing_loader = CSVLoader("non_existent_file.csv") # Use relative path intentionally for warning test
    missing_list = missing_loader.load()
    print(f"Loaded list from missing file: {missing_list}")
    missing_dict = missing_loader.load_dict(key_column='ID')
    print(f"Loaded dict from missing file: {missing_dict}")

    # Test missing column
    print(f"\n--- Testing missing column ---")
    missing_col_dict = customer_loader.load_dict(key_column='InvalidColumn')
    print(f"Loaded dict with missing key column: {missing_col_dict}")
    missing_val_col_dict = customer_loader.load_dict(key_column='CustomerID', value_column='InvalidValueCol')
    print(f"Loaded dict with missing value column: {missing_val_col_dict}")


    # --- Test Formatting Utils ---
    print("\n--- Testing Formatting Utils ---")
    print(f"Format 12345.67: {format_currency(12345.67)}")
    print(f"Format '12,345.67': {format_currency('12,345.67')}")
    print(f"Format 'abc': {format_currency('abc')}")
    print(f"Format None: {format_currency(None)}")
    ts = time.time()
    print(f"Timestamp {ts} to local: {get_local_time_from_utc(ts)}")
    print(f"Timestamp None to local: {get_local_time_from_utc(None)}")


    # --- Test Resource Path Util ---
    print("\n--- Testing Resource Path Util ---")
    # Assuming an 'assets' folder exists relative to utils.py for this test
    test_relative_path = os.path.join('assets', 'some_icon.png')
    resolved_path = get_resource_path(test_relative_path)
    print(f"Resolved path for '{test_relative_path}': {resolved_path}")