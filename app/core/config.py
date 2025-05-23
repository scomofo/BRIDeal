# app/core/config.py
import os
import json
import logging
from dotenv import load_dotenv
from typing import Any, Optional, Type, Union, List

# Initialize a logger for this module
# In a real app, the root logger would be configured by logger_config.py
# but we need a logger instance here for config loading messages.
logger = logging.getLogger(__name__)
# Basic config for this logger instance if no other config is set yet.
# This will be overridden if logger_config.setup_logging() is called later.
if not logger.hasHandlers():
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(name)s - %(message)s')
    logger.info("Basic logging configured for Config initialization.")


class Config:
    """
    Handles application configuration from environment variables and a JSON file.
    Environment variables take precedence over JSON file settings.
    """
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(Config, cls).__new__(cls)
        return cls._instance

    def __init__(self, env_path: Optional[str] = None, config_file_path: str = "config.json"):
        """
        Initializes the Config object. Loads .env file first, then JSON config file,
        then overrides with any environment variables already set.

        Args:
            env_path (Optional[str]): Path to the .env file. Defaults to project root .env.
            config_file_path (str): Path to the JSON configuration file. Defaults to 'config.json'.
        """
        if hasattr(self, '_initialized') and self._initialized:
            return # Already initialized (singleton pattern)

        self.settings = {}
        self.env_path = env_path
        self.config_file_path = config_file_path

        # 1. Load from .env file
        loaded_from_env_file = load_dotenv(dotenv_path=self.env_path, override=False) # override=False means os.environ takes precedence
        logger.info(f"Loaded environment variables from: {self.env_path if self.env_path else 'default .env location'} (Success: {loaded_from_env_file})")

        # 2. Load from JSON config file
        try:
            if os.path.exists(self.config_file_path):
                with open(self.config_file_path, 'r') as f:
                    json_config = json.load(f)
                    self.settings.update(json_config)
                logger.info(f"Loaded and applied configuration from {self.config_file_path}")
            else:
                logger.warning(f"JSON configuration file not found at {self.config_file_path}. Skipping.")
        except json.JSONDecodeError:
            logger.error(f"Error decoding JSON from {self.config_file_path}. Using defaults/env vars only.", exc_info=True)
        except Exception as e:
            logger.error(f"An unexpected error occurred while reading {self.config_file_path}: {e}", exc_info=True)


        # 3. Load/Override with actual environment variables (takes highest precedence)
        # This ensures that any environment variables set directly in the system
        # (or by a preceding .env load with override=True, or by load_dotenv with override=False if already set)
        # will take precedence.
        for key, value in os.environ.items():
            # Optionally, only load keys that are expected or have a certain prefix
            # For now, load all to mimic typical dotenv behavior where os.environ is the source of truth.
            # However, it's often better to explicitly map expected env vars.
            self.settings[key] = value
            logger.debug(f"Loaded from environment: {key} = '********'") # Mask sensitive values in logs

        logger.info("Environment variables processed and applied to configuration.")
        self._initialized = True


    def get(self, key: str, default: Any = None, var_type: Optional[Type] = None) -> Any:
        """
        Retrieves a configuration value.

        Args:
            key (str): The configuration key.
            default (Any, optional): The default value if the key is not found. Defaults to None.
            var_type (Optional[Type], optional): The expected type to cast the value to (e.g., int, bool, float).
                                                  For bool, "true", "1", "yes" (case-insensitive) are True.
                                                  Defaults to None (no type casting).

        Returns:
            Any: The configuration value, or the default if not found.
        """
        value = self.settings.get(key, default)

        if value is not None and var_type is not None:
            try:
                if var_type == bool:
                    if isinstance(value, str):
                        return value.lower() in ['true', '1', 'yes', 't', 'y']
                    return bool(value) # Fallback for non-string types
                elif var_type == list or var_type == dict:
                    # For list/dict, expect comma-separated string for list, or JSON string for dict
                    if isinstance(value, str):
                        if var_type == list:
                            return [item.strip() for item in value.split(',')]
                        elif var_type == dict:
                            try:
                                return json.loads(value)
                            except json.JSONDecodeError:
                                logger.warning(f"Could not parse JSON string for key '{key}': {value}. Returning as string.")
                                return value # Or raise error, or return default
                    return var_type(value) # Direct cast if not string or specific parsing failed
                return var_type(value)
            except ValueError as e:
                logger.warning(f"Could not cast value for key '{key}' ('{value}') to type {var_type}. Error: {e}. Returning default or original value.")
                # If casting fails, decide whether to return original 'value' or 'default'
                # Returning 'default' is safer if type is critical.
                # Returning 'value' might be okay if the caller can handle the original type.
                # For this implementation, let's return the original value if default is None, else default.
                return default if default is not None else value
            except Exception as e: # Catch any other casting errors
                logger.error(f"Unexpected error casting value for key '{key}': {e}. Returning default or original value.", exc_info=True)
                return default if default is not None else value
        return value

    def get_all_settings(self) -> dict:
        """Returns all current settings."""
        return self.settings.copy()

    # --- Methods for JD API Configuration Check ---
    def is_jd_api_fully_configured(self) -> bool:
        """
        Checks if all essential JD API configurations (including base URL and OAuth) are present.
        This is a more comprehensive check.
        """
        return self.is_jd_api_base_configured() and self.is_jd_oauth_configured()

    def is_jd_api_base_configured(self) -> bool:
        """Checks if the JD API base URL is configured."""
        if not self.get("JD_API_BASE_URL"):
            logger.debug("JD API configuration check: JD_API_BASE_URL is missing.")
            return False
        logger.debug("JD API configuration check: JD_API_BASE_URL is present.")
        return True

    def is_jd_oauth_configured(self) -> bool:
        """
        Checks if essential JD OAuth specific configurations are present.
        Adjust the list of required_oauth_keys based on what JDAuthManager
        strictly needs to be considered 'configurable' even if not fully operational
        for all OAuth flows.
        """
        required_oauth_keys = [
            "JD_CLIENT_ID",
            "DEERE_CLIENT_SECRET",
            # "JD_AUTH_URL",  # Uncomment if JDAuthManager cannot initialize meaningfully without them
            # "JD_TOKEN_URL"
        ]
        for key in required_oauth_keys:
            if not self.get(key):
                logger.debug(f"JD OAuth configuration check: {key} is missing.")
                return False
        logger.debug("JD OAuth configuration check: All specified keys are present.")
        return True

    def is_sharepoint_configured(self) -> bool:
        """Checks if essential SharePoint configurations are present."""
        # Based on the log: "SharePoint core credentials (client_id, client_secret, tenant_id) are not fully configured."
        # The application might be looking for generic names or specific SHAREPOINT_ prefixed names.
        # This check assumes the application internally maps/uses specific keys for SharePoint.
        # Let's assume it looks for these specific keys after loading from .env
        required_sp_keys = [
            # These are the names the SharePointManager log message implies it's looking for.
            # Your .env might have AZURE_CLIENT_ID etc. The Config class or SharePointManager
            # needs to handle the mapping if names differ.
            # For this check, we'll assume the Config.get() would resolve to the correct value
            # if .env has SHAREPOINT_CLIENT_ID or if AZURE_CLIENT_ID is mapped to it internally.
            "SHAREPOINT_CLIENT_ID", # Or the key it's mapped to, e.g., self.get('AZURE_CLIENT_ID') if mapped
            "SHAREPOINT_CLIENT_SECRET",
            "SHAREPOINT_TENANT_ID",
            "SHAREPOINT_SITE_ID", # From your .env
            "SHAREPOINT_DRIVE_ID" # From your .env
        ]
        for key in required_sp_keys:
            if not self.get(key):
                logger.debug(f"SharePoint configuration check: {key} is missing.")
                return False
        logger.debug("SharePoint configuration check: All specified keys are present.")
        return True


# Global instance (singleton pattern)
# This allows other modules to import and use `app_config` directly.
# app_config = Config()

# Example usage (typically in main.py or at the start of the app):
if __name__ == "__main__":
    # This part is for testing the Config class itself.
    # Create a dummy .env for testing
    with open(".env.test_config", "w") as f:
        f.write("APP_NAME=TestAppFromEnv\n")
        f.write("DEBUG_MODE=True\n")
        f.write("PORT=8080\n")
        f.write("API_KEYS=key1,key2,key3\n")
        f.write("JD_CLIENT_ID=jd_client_id_from_env\n")
        f.write("DEERE_CLIENT_SECRET=deere_secret_from_env\n")
        f.write("JD_API_BASE_URL=https://jd.api.example.com\n")
        f.write("SHAREPOINT_CLIENT_ID=sp_client_id_from_env\n")
        f.write("SHAREPOINT_CLIENT_SECRET=sp_secret_from_env\n")
        f.write("SHAREPOINT_TENANT_ID=sp_tenant_id_from_env\n")
        f.write("SHAREPOINT_SITE_ID=sp_site_id_from_env\n")
        f.write("SHAREPOINT_DRIVE_ID=sp_drive_id_from_env\n")


    # Create a dummy config.json for testing
    dummy_json_config = {
        "APP_NAME": "TestAppFromJson", # Will be overridden by .env if present
        "DEFAULT_TIMEOUT": 30,
        "FEATURE_FLAGS": {
            "new_dashboard": True,
            "experimental_reports": False
        },
        "PORT": "9090" # Will be overridden by .env
    }
    with open("config.test_config.json", "w") as f:
        json.dump(dummy_json_config, f, indent=4)

    print("--- Testing Config Initialization ---")
    # Initialize with test files
    test_config = Config(env_path=".env.test_config", config_file_path="config.test_config.json")

    print(f"\nApp Name (from .env): {test_config.get('APP_NAME')}")
    print(f"Debug Mode (from .env, as bool): {test_config.get('DEBUG_MODE', var_type=bool)}")
    print(f"Port (from .env, as int): {test_config.get('PORT', var_type=int)}")
    print(f"Default Timeout (from JSON): {test_config.get('DEFAULT_TIMEOUT', var_type=int)}")
    print(f"API Keys (from .env, as list): {test_config.get('API_KEYS', var_type=list)}")
    print(f"Feature Flags (from JSON, as dict): {test_config.get('FEATURE_FLAGS', var_type=dict)}")
    print(f"Non-existent key with default: {test_config.get('NON_EXISTENT_KEY', 'default_value')}")
    print(f"Non-existent key without default: {test_config.get('NON_EXISTENT_KEY_NO_DEFAULT')}")

    print("\n--- Testing JD API Configuration Checks ---")
    print(f"Is JD API Base Configured? {test_config.is_jd_api_base_configured()}")
    print(f"Is JD OAuth Configured? {test_config.is_jd_oauth_configured()}")
    print(f"Is JD API Fully Configured? {test_config.is_jd_api_fully_configured()}")

    print("\n--- Testing SharePoint Configuration Check ---")
    print(f"Is SharePoint Configured? {test_config.is_sharepoint_configured()}")


    print("\n--- Testing with missing JD values ---")
    # Simulate missing JD_API_BASE_URL by temporarily removing it from settings for this test instance
    original_base_url = test_config.settings.pop("JD_API_BASE_URL", None)
    print(f"Is JD API Base Configured (after removing BASE_URL)? {test_config.is_jd_api_base_configured()}")
    print(f"Is JD API Fully Configured (after removing BASE_URL)? {test_config.is_jd_api_fully_configured()}")
    if original_base_url: # Add it back if it was there
        test_config.settings["JD_API_BASE_URL"] = original_base_url


    # Clean up dummy files
    try:
        os.remove(".env.test_config")
        os.remove("config.test_config.json")
    except OSError as e:
        print(f"Error removing test files: {e}")

