# app/services/api_clients/jd_quote_client.py
import logging
from typing import Optional, Dict, Any
import requests # For making HTTP requests
import os # Added for cleanup in __main__

# Assuming Config class is in app.core.config
from app.core.config import Config
# Assuming JDAuthManager is for getting tokens
from app.services.integrations.jd_auth_manager import JDAuthManager

logger = logging.getLogger(__name__)

class JDQuoteApiClient:
    """
    API Client for interacting with the John Deere Quoting System API.
    Handles making HTTP requests to the API endpoints.
    """
    def __init__(self, config: Config, auth_manager: Optional[JDAuthManager] = None):
        """
        Initializes the JDQuoteApiClient.

        Args:
            config (Config): The application's configuration object.
            auth_manager (Optional[JDAuthManager]): The authentication manager for JD API.
        """
        self.config = config
        self.auth_manager = auth_manager
        self.base_url: Optional[str] = None
        self.is_operational: bool = False
        self.default_headers = {
            "Content-Type": "application/json",
            "Accept": "application/json"
        }

        if not self.config:
            logger.error("JDQuoteApiClient: Config object not provided. Client will be non-operational.")
            return # Cannot proceed without config

        # Check if the base API URL is configured
        # Using the specific check method from Config class
        if self.config.is_jd_api_base_configured():
            self.base_url = self.config.get("JD_API_BASE_URL")
            self.is_operational = True
            logger.info(f"JDQuoteApiClient initialized with base URL: {self.base_url}. Client is operational.")

            # Further check on auth_manager if it's strictly required for all operations
            if self.auth_manager:
                if not self.auth_manager.is_operational:
                    logger.warning("JDQuoteApiClient: JDAuthManager is provided but not operational. Authentication-dependent calls will fail.")
                    # Depending on API design, client might still be "operational" for public endpoints.
                    # For now, base_url makes it operational; individual methods will handle auth.
            else:
                logger.warning("JDQuoteApiClient: JDAuthManager not provided. Only unauthenticated API calls can be made.")
        else:
            logger.warning("JDQuoteApiClient: JD_API_BASE_URL is not configured. Client will be non-operational.")
            # self.is_operational remains False

    def _get_auth_header(self) -> Dict[str, str]:
        """Prepares the Authorization header with the access token."""
        if self.auth_manager:
            token = self.auth_manager.get_access_token()
            if token:
                return {"Authorization": f"Bearer {token}"}
            else:
                logger.warning("JDQuoteApiClient: No access token available from auth_manager for authenticated request.")
        return {}

    def _request(self, method: str, endpoint: str, data: Optional[Dict[str, Any]] = None,
                 params: Optional[Dict[str, Any]] = None, requires_auth: bool = True) -> Optional[Dict[str, Any]]:
        """
        Helper method to make HTTP requests to the JD API.

        Args:
            method (str): HTTP method (GET, POST, PUT, DELETE).
            endpoint (str): API endpoint path (e.g., '/quotes').
            data (Optional[Dict[str, Any]]): JSON payload for POST/PUT requests.
            params (Optional[Dict[str, Any]]): URL parameters for GET requests.
            requires_auth (bool): Whether the endpoint requires authentication.

        Returns:
            Optional[Dict[str, Any]]: The JSON response from the API, or None on failure.
        """
        if not self.is_operational:
            logger.error(f"JDQuoteApiClient: Cannot make request. Client is not operational (Base URL: {self.base_url}).")
            return None

        if not self.base_url: # Should be caught by is_operational, but as a safeguard
            logger.error("JDQuoteApiClient: Base URL not set. Cannot make API calls.")
            return None

        url = f"{self.base_url.rstrip('/')}/{endpoint.lstrip('/')}"
        headers = self.default_headers.copy()

        if requires_auth:
            auth_header = self._get_auth_header()
            if not auth_header: # No token available for an authenticated endpoint
                logger.error(f"JDQuoteApiClient: Authentication required for {url} but no token is available.")
                return None
            headers.update(auth_header)

        try:
            logger.debug(f"JDQuoteApiClient: Making {method} request to {url} with data: {data}, params: {params}")
            response = requests.request(method, url, headers=headers, json=data, params=params, timeout=30) # 30s timeout
            response.raise_for_status()  # Raise HTTPError for bad responses (4xx or 5xx)
            
            # Handle cases where response might be empty but successful (e.g., 204 No Content)
            if response.status_code == 204:
                return {"status": "success", "message": "Operation successful with no content returned."}
            
            return response.json() # Parse JSON response
        except requests.exceptions.HTTPError as http_err:
            logger.error(f"JDQuoteApiClient: HTTP error occurred: {http_err} - Response: {http_err.response.text if http_err.response else 'No response text'}")
        except requests.exceptions.ConnectionError as conn_err:
            logger.error(f"JDQuoteApiClient: Connection error occurred: {conn_err}")
        except requests.exceptions.Timeout as timeout_err:
            logger.error(f"JDQuoteApiClient: Timeout error occurred: {timeout_err}")
        except requests.exceptions.RequestException as req_err:
            logger.error(f"JDQuoteApiClient: An unexpected error occurred during request: {req_err}")
        except ValueError as json_err: # Includes json.JSONDecodeError
            logger.error(f"JDQuoteApiClient: Failed to decode JSON response from {url}: {json_err}")
        return None

    # --- API Methods ---

    def get_quote_details(self, quote_id: str) -> Optional[Dict[str, Any]]:
        """
        Retrieves details for a specific quote.
        Example: GET /quotes/{quote_id}
        """
        if not self.is_operational:
            logger.warning("JDQuoteApiClient: Cannot get quote details, client not operational.")
            return None
        
        logger.info(f"JDQuoteApiClient: Requesting details for quote ID: {quote_id}")
        
        # Make the actual API call
        return self._request("GET", f"quotes/{quote_id}")

    def submit_new_quote(self, quote_data: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """
        Submits a new quote.
        Example: POST /quotes
        """
        if not self.is_operational:
            logger.warning("JDQuoteApiClient: Cannot submit new quote, client not operational.")
            return None
        
        logger.info(f"JDQuoteApiClient: Submitting new quote")
        
        # Make the actual API call
        return self._request("POST", "quotes", data=quote_data)

    def update_existing_quote(self, quote_id: str, update_data: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """
        Updates an existing quote.
        Example: PUT /quotes/{quote_id}
        """
        if not self.is_operational:
            logger.warning("JDQuoteApiClient: Cannot update quote, client not operational.")
            return None
        
        logger.info(f"JDQuoteApiClient: Updating quote ID {quote_id}")
        
        # Make the actual API call
        return self._request("PUT", f"quotes/{quote_id}", data=update_data)

# Example Usage (for testing this module standalone)
if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - [%(module)s.%(funcName)s:%(lineno)d] - %(message)s')

    # Simulate Config object
    class MockConfigClient(Config):
        def __init__(self, settings_dict):
            self.settings = settings_dict
            super().__init__(env_path=".env.test_jd_client") # Dummy path for super init
            self.settings.update(settings_dict) # Ensure test settings override

    # Simulate JDAuthManager
    class MockJDAuthManager:
        def __init__(self, operational=True, token="test_access_token_123"):
            self.is_operational = operational
            self.token = token
        def get_access_token(self):
            return self.token if self.is_operational else None

    # --- Test Case 1: Client Operational ---
    print("\n--- Test Case 1: Client Operational ---")
    mock_env_client_ok = {"JD_API_BASE_URL": "https://mock-jd-api.example.com"}
    config_ok = MockConfigClient(mock_env_client_ok)
    auth_manager_ok = MockJDAuthManager()
    client_ok = JDQuoteApiClient(config=config_ok, auth_manager=auth_manager_ok)
    print(f"Client Operational: {client_ok.is_operational}, Base URL: {client_ok.base_url}")
    if client_ok.is_operational:
        # These are real API calls now, but they won't actually run since we're using a mock URL
        print(f"Get Quote Details: {client_ok.get_quote_details('Q-001')}")
        print(f"Submit New Quote: {client_ok.submit_new_quote({'customer_id': 'CUST-789', 'total_amount': 50000})}")

    # --- Test Case 2: Client Not Operational (Base URL missing) ---
    print("\n--- Test Case 2: Client Not Operational (Base URL missing) ---")
    mock_env_client_no_url = {} # JD_API_BASE_URL is missing
    config_no_url = MockConfigClient(mock_env_client_no_url)
    client_no_url = JDQuoteApiClient(config=config_no_url, auth_manager=auth_manager_ok)
    print(f"Client Operational: {client_no_url.is_operational}, Base URL: {client_no_url.base_url}")
    # Attempting a call should result in a warning and None
    print(f"Get Quote Details (should be None): {client_no_url.get_quote_details('Q-002')}")

    # --- Test Case 3: Client Operational but Auth Manager Not ---
    print("\n--- Test Case 3: Client Operational, Auth Manager Not ---")
    auth_manager_not_ok = MockJDAuthManager(operational=False)
    client_auth_issues = JDQuoteApiClient(config=config_ok, auth_manager=auth_manager_not_ok)
    print(f"Client Operational: {client_auth_issues.is_operational}")
    # In a real scenario with _request making actual calls, authenticated endpoints would fail here.
    # We'd expect _get_auth_header to return empty, and _request to log an error if auth is required.
    print(f"Submit New Quote (auth manager not ok): {client_auth_issues.submit_new_quote({'item': 'tractor'})}")


    # Clean up dummy .env file if created by MockConfigClient's super().__init__
    if os.path.exists(".env.test_jd_client"):
        os.remove(".env.test_jd_client")