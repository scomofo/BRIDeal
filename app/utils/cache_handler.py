# bridleal_refactored/app/utils/cache_handler.py
import json
import os
import time
import logging
import hashlib # For creating unique cache keys from complex inputs if needed

# Attempt to import constants for default paths/settings
try:
    from . import constants as app_constants
except ImportError:
    app_constants = None # Fallback if constants can't be imported

# Get a logger for this module
logger = logging.getLogger(__name__)

class CacheHandler:
    """
    Handles caching and retrieval of data to/from JSON files.
    Manages cache expiration.
    """
    def __init__(self, cache_dir=None, default_duration_seconds=None, config=None):
        """
        Initialize the CacheHandler.

        Args:
            cache_dir (str, optional): The directory to store cache files.
                                       Defaults to 'cache' in the app's data directory.
            default_duration_seconds (int, optional): Default cache duration in seconds.
                                                      Defaults to value from constants or 1 hour.
            config (Config, optional): Application configuration object.
        """
        self.config = config
        
        if cache_dir:
            self.cache_dir = cache_dir
        elif self.config and self.config.get('CACHE_DIR'):
            self.cache_dir = self.config.get('CACHE_DIR')
        else:
            # Fallback to a 'cache' subdirectory in the current working directory or a user-specific app data dir
            # For robustness, an application should define CACHE_DIR in its config.
            # Using os.path.expanduser("~/.bridleal_refactored/cache") or similar might be better for user-specific installs.
            self.cache_dir = os.path.join(os.getcwd(), "cache") 
            logger.warning(f"CACHE_DIR not specified in config, defaulting to: {self.cache_dir}")

        if default_duration_seconds is not None:
            self.default_duration_seconds = default_duration_seconds
        elif self.config and self.config.get('DEFAULT_CACHE_DURATION_SECONDS', var_type=int):
            self.default_duration_seconds = self.config.get('DEFAULT_CACHE_DURATION_SECONDS', var_type=int)
        elif app_constants and hasattr(app_constants, 'DEFAULT_CACHE_EXPIRATION_SECONDS'):
            self.default_duration_seconds = app_constants.DEFAULT_CACHE_EXPIRATION_SECONDS
        else:
            self.default_duration_seconds = 3600  # Default to 1 hour

        try:
            if not os.path.exists(self.cache_dir):
                os.makedirs(self.cache_dir, exist_ok=True)
                logger.info(f"Cache directory created: {self.cache_dir}")
        except OSError as e:
            logger.error(f"Error creating cache directory {self.cache_dir}: {e}")
            # Potentially raise an error or operate in a no-cache mode
            self.cache_dir = None # Disable caching if directory can't be created

        logger.info(f"CacheHandler initialized. Cache directory: {self.cache_dir}, Default duration: {self.default_duration_seconds}s")

    def _get_cache_filepath(self, cache_key, subfolder=None):
        """Constructs the full path for a given cache key and optional subfolder."""
        if not self.cache_dir:
            return None # Caching disabled

        # Sanitize cache_key to be a valid filename
        # Replace common problematic characters. A more robust solution might involve hashing.
        safe_cache_key = "".join(c if c.isalnum() or c in ('_', '-') else '_' for c in cache_key)
        if not safe_cache_key.endswith(".json"):
            safe_cache_key += ".json"
        
        current_cache_dir = self.cache_dir
        if subfolder:
            current_cache_dir = os.path.join(self.cache_dir, subfolder)
            if not os.path.exists(current_cache_dir):
                try:
                    os.makedirs(current_cache_dir, exist_ok=True)
                except OSError as e:
                    logger.error(f"Error creating cache subfolder {current_cache_dir}: {e}")
                    return None # Cannot create subfolder, cannot cache here
        
        return os.path.join(current_cache_dir, safe_cache_key)

    def get(self, cache_key, subfolder=None, duration_seconds=None):
        """
        Retrieve data from the cache.

        Args:
            cache_key (str): A unique key for the cached item.
            subfolder (str, optional): An optional subfolder within the cache directory.
            duration_seconds (int, optional): Specific cache duration for this item.
                                             Overrides default if provided.

        Returns:
            The cached data if found and not expired, otherwise None.
        """
        if not self.cache_dir:
            logger.warning("Cache directory not available. Cannot get cache item.")
            return None

        filepath = self._get_cache_filepath(cache_key, subfolder)
        if not filepath or not os.path.exists(filepath):
            logger.debug(f"Cache miss (file not found): {cache_key} in subfolder '{subfolder}'")
            return None

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                cache_entry = json.load(f)
            
            timestamp = cache_entry.get('timestamp')
            data = cache_entry.get('data')
            
            if timestamp is None or data is None:
                logger.warning(f"Invalid cache file format for {cache_key}. Removing.")
                self.remove(cache_key, subfolder)
                return None

            effective_duration = duration_seconds if duration_seconds is not None else self.default_duration_seconds
            
            if time.time() - timestamp > effective_duration:
                logger.info(f"Cache expired for {cache_key}. Removing.")
                self.remove(cache_key, subfolder) # Remove expired cache
                return None
            
            logger.info(f"Cache hit for {cache_key} in subfolder '{subfolder}'.")
            return data
        except json.JSONDecodeError:
            logger.error(f"Error decoding JSON from cache file {filepath}. Removing.")
            self.remove(cache_key, subfolder)
            return None
        except Exception as e:
            logger.error(f"Error reading cache file {filepath}: {e}")
            return None

    def set(self, cache_key, data, subfolder=None, duration_seconds=None):
        """
        Store data in the cache.

        Args:
            cache_key (str): A unique key for the cached item.
            data: The data to cache (must be JSON serializable).
            subfolder (str, optional): An optional subfolder within the cache directory.
            duration_seconds (int, optional): Specific cache duration for this item.
                                             Does not affect retrieval duration check,
                                             but could be stored if needed for advanced logic.
        """
        if not self.cache_dir:
            logger.warning("Cache directory not available. Cannot set cache item.")
            return

        filepath = self._get_cache_filepath(cache_key, subfolder)
        if not filepath:
            logger.error(f"Could not determine cache filepath for {cache_key}. Caching failed.")
            return

        cache_entry = {
            'timestamp': time.time(),
            'data': data
            # 'duration': duration_seconds or self.default_duration_seconds # Could store duration if needed
        }
        
        try:
            # Ensure the directory for the file exists (especially if subfolder was just created)
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(cache_entry, f, indent=4) # Use indent for readability
            logger.info(f"Cached data for {cache_key} in subfolder '{subfolder}' at {filepath}")
        except TypeError as e: # Data not JSON serializable
            logger.error(f"Data for cache key {cache_key} is not JSON serializable: {e}")
        except Exception as e:
            logger.error(f"Error writing cache file {filepath}: {e}")

    def remove(self, cache_key, subfolder=None):
        """Remove a specific item from the cache."""
        if not self.cache_dir:
            logger.warning("Cache directory not available. Cannot remove cache item.")
            return False

        filepath = self._get_cache_filepath(cache_key, subfolder)
        if filepath and os.path.exists(filepath):
            try:
                os.remove(filepath)
                logger.info(f"Removed cache file: {filepath}")
                return True
            except Exception as e:
                logger.error(f"Error removing cache file {filepath}: {e}")
                return False
        logger.debug(f"Cache file not found for removal: {filepath}")
        return False

    def clear_all(self, subfolder=None):
        """Clear all cache files in the specified subfolder or the entire cache directory if no subfolder."""
        if not self.cache_dir:
            logger.warning("Cache directory not available. Cannot clear cache.")
            return

        target_dir = os.path.join(self.cache_dir, subfolder) if subfolder else self.cache_dir
        
        if not os.path.exists(target_dir):
            logger.info(f"Cache directory/subfolder to clear does not exist: {target_dir}")
            return

        logger.info(f"Clearing cache in directory: {target_dir}")
        cleared_count = 0
        error_count = 0
        for filename in os.listdir(target_dir):
            if filename.endswith(".json"): # Only remove .json files (our cache files)
                filepath = os.path.join(target_dir, filename)
                try:
                    os.remove(filepath)
                    cleared_count += 1
                except Exception as e:
                    logger.error(f"Error removing cache file {filepath} during clear_all: {e}")
                    error_count +=1
        logger.info(f"Cache cleared in {target_dir}. Removed {cleared_count} files, {error_count} errors.")

    def generate_key_from_params(self, base_name, params_dict):
        """
        Generates a cache key from a base name and a dictionary of parameters.
        Useful for caching results of functions with multiple arguments.
        """
        if not isinstance(params_dict, dict):
            logger.warning("params_dict for generate_key_from_params should be a dictionary.")
            return base_name

        # Create a stable string representation of the parameters
        # Sort by key to ensure consistent order
        param_string = "_".join(f"{k}_{v}" for k, v in sorted(params_dict.items()))
        
        # Use hashlib for a more robust and shorter key if param_string gets too long or complex
        # For now, simple concatenation for readability if keys are not too complex
        # If param_string is potentially very long or has difficult characters:
        # return f"{base_name}_{hashlib.md5(param_string.encode('utf-8')).hexdigest()}"
        
        # Simple concatenation for now
        key = f"{base_name}_{param_string}"
        # Sanitize further if needed, _get_cache_filepath does basic sanitization
        return key


# Example Usage
if __name__ == '__main__':
    # Basic logger for testing
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(name)s - %(message)s')

    # Mock config for testing
    class MockConfig:
        def get(self, key, default=None, var_type=None):
            if key == 'CACHE_DIR':
                return 'test_cache_dir' # Use a 'test_cache_dir' subdirectory
            if key == 'DEFAULT_CACHE_DURATION_SECONDS':
                return 60 # 1 minute for testing
            return default

    test_config = MockConfig()
    
    # Ensure test_cache_dir exists and is clean for the example
    test_cache_path = test_config.get('CACHE_DIR')
    if os.path.exists(test_cache_path):
        import shutil
        shutil.rmtree(test_cache_path) # Remove old test cache
    os.makedirs(test_cache_path, exist_ok=True)
    os.makedirs(os.path.join(test_cache_path, "sub"), exist_ok=True)


    cache = CacheHandler(config=test_config)
    
    # Test basic set and get
    cache.set("my_data_key", {"name": "Test User", "value": 123})
    retrieved_data = cache.get("my_data_key")
    logger.info(f"Retrieved data: {retrieved_data}")
    assert retrieved_data == {"name": "Test User", "value": 123}

    # Test with subfolder
    cache.set("sub_data_key", {"info": "In subfolder"}, subfolder="sub")
    retrieved_sub_data = cache.get("sub_data_key", subfolder="sub")
    logger.info(f"Retrieved sub data: {retrieved_sub_data}")
    assert retrieved_sub_data == {"info": "In subfolder"}

    # Test expiration
    cache.set("expiring_key", "This will expire soon", duration_seconds=1) # Set with custom duration for get
    logger.info("Waiting for cache to expire (2 seconds)...")
    time.sleep(2)
    expired_data = cache.get("expiring_key", duration_seconds=1) # Check with the same duration
    logger.info(f"Expired data: {expired_data}")
    assert expired_data is None 

    # Test get with a different (longer) duration than default set for the item
    cache.set("short_lived", "data", duration_seconds=5) # This duration is not stored with item
    time.sleep(1)
    # If we try to get it with a shorter duration than its age, it should be considered expired
    retrieved_short_custom_duration = cache.get("short_lived", duration_seconds=0)
    assert retrieved_short_custom_duration is None, "Should be expired with 0s duration"
    # If we get it with a longer duration than its age, it should be fine
    retrieved_short_default_duration = cache.get("short_lived") # Uses default_duration_seconds (60s)
    assert retrieved_short_default_duration == "data", "Should not be expired with default duration"


    # Test remove
    cache.set("to_remove", "data to be removed")
    assert cache.get("to_remove") == "data to be removed"
    cache.remove("to_remove")
    assert cache.get("to_remove") is None

    # Test generate_key_from_params
    params1 = {"id": 1, "type": "user"}
    params2 = {"type": "user", "id": 1} # Same params, different order
    key1 = cache.generate_key_from_params("user_data", params1)
    key2 = cache.generate_key_from_params("user_data", params2)
    logger.info(f"Generated key 1: {key1}")
    logger.info(f"Generated key 2: {key2}")
    assert key1 == key2 # Keys should be the same due to sorting

    # Test clear_all
    cache.set("item1_in_root", "data1")
    cache.set("item2_in_sub", "data2", subfolder="sub")
    cache.clear_all(subfolder="sub")
    assert cache.get("item1_in_root") == "data1" # Root item should remain
    assert cache.get("item2_in_sub", subfolder="sub") is None # Subfolder item should be gone
    
    cache.clear_all() # Clear root
    assert cache.get("item1_in_root") is None

    logger.info("CacheHandler tests completed.")
    # Clean up test directory
    if os.path.exists(test_cache_path):
        shutil.rmtree(test_cache_path)
