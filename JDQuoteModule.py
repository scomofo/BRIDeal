import os
import json
import requests
import secrets
import urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
import threading
import logging
import time
import webbrowser
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QListWidget, QMessageBox, 
    QInputDialog, QHBoxLayout, QTextEdit, QApplication
)
from PyQt5.QtCore import Qt, QCoreApplication, QTimer
import atexit  # For saving cache on exit
import base64
# --- Constants ---
# Enable debug logging - Change to logging.INFO for less verbose output
logging.basicConfig(level=logging.DEBUG) 

# Deere OAuth2 Settings
TENANT_ID = "aus78tnlaysMraFhC1t7"  # Okta tenant ID for John Deere
CLIENT_ID = "0oao5jntk71YDUX9Q5d7"  # John Deere client ID
# IMPORTANT: Store secrets securely (e.g., env variables, Key Vault), not directly in code!
CLIENT_SECRET = os.environ.get("DEERE_CLIENT_SECRET")  # LOADED FROM ENVIRONMENT

# John Deere OAuth endpoints 
AUTH_BASE = "https://signin.johndeere.com/oauth2"
DEVICE_CODE_ENDPOINT = f"{AUTH_BASE}/{TENANT_ID}/v1/device/authorize"
TOKEN_ENDPOINT = f"{AUTH_BASE}/{TENANT_ID}/v1/token"
# API endpoints
API_BASE_URL = "https://api.deere.com/platform/maintain-quotes/v1"

# Scopes required for the API
SCOPES = ["axiom"]

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".jdquote_config.json")  # For saving Org ID
TOKEN_CACHE_PATH = os.path.join(os.path.expanduser("~"), ".jdquote_token.json")  # For saving tokens

class JDQuoteModule(QWidget):
    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window
        self.access_token = None
        self.refresh_token = None
        self.token_expiry = 0
        self.organization_id = self.load_org_id()
        self.device_flow_timer = None
        self.device_flow_data = None
        
        # Load tokens from cache if available
        self.load_tokens()
        
        # Register cache persistence on exit
        if not hasattr(QCoreApplication.instance(), '_token_cache_registered'):
            atexit.register(self.save_tokens)
            QCoreApplication.instance()._token_cache_registered = True
            logging.info("Registered token cache saving function on exit.")

        # Setup UI first so elements exist even if initialization fails
        self.setupUi() 
        logging.info("JDQuoteModule UI Initialized.")

        # Check if we have proper configuration
        try:
            if not CLIENT_ID:
                raise ValueError("CLIENT_ID is not configured properly")
                
            actual_client_secret = os.environ.get("DEERE_CLIENT_SECRET")
            if not actual_client_secret:
                logging.error("CRITICAL: DEERE_CLIENT_SECRET environment variable not set.")
                raise ValueError("Client secret environment variable 'DEERE_CLIENT_SECRET' is missing.")
            else:
                # Log only the first few characters for security
                logging.info(f"DEERE_CLIENT_SECRET is set and has length: {len(actual_client_secret)}")
                self._enable_module_ui()
            
            if self.access_token:
                self._update_status("Previously authenticated token loaded")
            
        except ValueError as val_err:
            logging.error(f"Failed to initialize Okta client: {val_err}", exc_info=True)
            QMessageBox.critical(self, "Auth Config Error", f"Configuration error:\n{val_err}")
            self._disable_module_ui("Auth Config Error")
        except Exception as app_init_err:
            logging.error(f"Unexpected error initializing: {app_init_err}", exc_info=True)
            QMessageBox.critical(self, "Init Error", f"Unexpected error initializing:\n{app_init_err}")
            self._disable_module_ui("Init Error")

    def load_tokens(self):
        """Load tokens from cache file"""
        try:
            if os.path.exists(TOKEN_CACHE_PATH):
                with open(TOKEN_CACHE_PATH, 'r') as f:
                    data = json.load(f)
                    self.access_token = data.get("access_token")
                    self.refresh_token = data.get("refresh_token")
                    self.token_expiry = data.get("expiry", 0)
                    
                    # Check if token is expired
                    if self.token_expiry < time.time():
                        logging.info("Cached token is expired, will need to reauthenticate")
                        self.access_token = None
                        self.refresh_token = None
                    else:
                        logging.info("Loaded valid token from cache")
        except Exception as e:
            logging.warning(f"Failed to load tokens from cache: {e}")
            self.access_token = None
            self.refresh_token = None
            self.token_expiry = 0

    def save_tokens(self):
        """Save tokens to cache file"""
        if self.access_token:
            try:
                with open(TOKEN_CACHE_PATH, 'w') as f:
                    json.dump({
                        "access_token": self.access_token,
                        "refresh_token": self.refresh_token,
                        "expiry": self.token_expiry
                    }, f)
                logging.info("Saved tokens to cache")
            except Exception as e:
                logging.error(f"Failed to save tokens to cache: {e}")

    def _enable_module_ui(self):
        """Enables UI elements after successful initialization."""
        self.auth_btn.setEnabled(True)
        self.org_btn.setEnabled(True)
        self.refresh_btn.setEnabled(True)
        self.new_btn.setEnabled(True)
        self.delete_btn.setEnabled(True)

    def _disable_module_ui(self, reason="Initialization Error"):
        """Disables UI elements if initialization fails."""
        logging.warning(f"Disabling JDQuoteModule UI due to: {reason}")
        if hasattr(self, 'status_label'):
            self.status_label.setText(f"Error: {reason}")
        else:
            print(f"Cannot set status label text - UI setup likely incomplete due to error: {reason}")
              
        if hasattr(self, 'auth_btn'):
            self.auth_btn.setEnabled(False)
        if hasattr(self, 'org_btn'):
            self.org_btn.setEnabled(False)
        if hasattr(self, 'refresh_btn'):
            self.refresh_btn.setEnabled(False)
        if hasattr(self, 'new_btn'):
            self.new_btn.setEnabled(False)
        if hasattr(self, 'delete_btn'):
            self.delete_btn.setEnabled(False)

    def setupUi(self):
        """Creates the UI elements for the module."""
        layout = QVBoxLayout(self)

        self.status_label = QLabel("Status: Initializing.")
        layout.addWidget(self.status_label)

        btn_layout_top = QHBoxLayout()
        self.auth_btn = QPushButton("🔑 Authenticate") 
        self.auth_btn.setToolTip("Login via John Deere Device Flow")
        self.auth_btn.clicked.connect(self.authenticate)
        btn_layout_top.addWidget(self.auth_btn)

        self.org_btn = QPushButton("🏢 Get/Set Org ID") 
        self.org_btn.setToolTip("Fetch available organizations and select one")
        self.org_btn.clicked.connect(self.fetch_organization_id)
        btn_layout_top.addWidget(self.org_btn)
        layout.addLayout(btn_layout_top)

        self.org_id_label = QLabel(f"Current Org ID: {self.organization_id or 'Not Set'}")
        layout.addWidget(self.org_id_label)
        
        self.refresh_btn = QPushButton("🔄 Load Quotes") 
        self.refresh_btn.setToolTip("Fetch quotes for the selected organization")
        self.refresh_btn.clicked.connect(self.load_quotes)
        layout.addWidget(self.refresh_btn)

        self.quote_list = QListWidget()
        self.quote_list.setToolTip("List of quotes fetched from the API")
        layout.addWidget(self.quote_list)

        btn_layout_bottom = QHBoxLayout()
        self.new_btn = QPushButton("➕ Create Quote") 
        self.new_btn.setToolTip("Create a new quote with a description")
        self.new_btn.clicked.connect(self.create_quote)
        btn_layout_bottom.addWidget(self.new_btn)

        self.delete_btn = QPushButton("🗑 Delete Selected Quote") 
        self.delete_btn.setToolTip("Delete the quote selected in the list above")
        self.delete_btn.clicked.connect(self.delete_quote)
        btn_layout_bottom.addWidget(self.delete_btn)
        layout.addLayout(btn_layout_bottom)

        log_label = QLabel("API Response / Log:")
        layout.addWidget(log_label)
        self.response_view = QTextEdit()
        self.response_view.setReadOnly(True)
        self.response_view.setLineWrapMode(QTextEdit.NoWrap)
        self.response_view.setStyleSheet("font-family: Consolas, Courier New, monospace; background-color: #f0f0f0;")
        self.response_view.setMinimumHeight(100)
        layout.addWidget(self.response_view)

        # Initially disable buttons until initialization completes
        self._disable_module_ui("Initializing")

    def load_org_id(self):
        """Loads the last used organization ID from a config file."""
        try:
            if os.path.exists(CONFIG_PATH):
                with open(CONFIG_PATH, 'r') as f:
                    data = json.load(f)
                    org_id = data.get("organization_id")
                    logging.info(f"Loaded Org ID from {CONFIG_PATH}: {org_id}")
                    return org_id
        except Exception as e:
            logging.warning(f"Could not load org ID from {CONFIG_PATH}: {e}")
        return None

    def save_org_id(self, org_id):
        """Saves the selected organization ID to a config file."""
        try:
            with open(CONFIG_PATH, 'w') as f:
                json.dump({"organization_id": org_id}, f)
            logging.info(f"Saved Org ID to {CONFIG_PATH}: {org_id}")
        except Exception as e:
            logging.error(f"Failed to save org ID to {CONFIG_PATH}: {e}")
            QMessageBox.warning(self, "Config Save Error", f"Could not save organization ID config:\n{e}")

    def _update_status(self, message, is_error=False):
        """Helper to update status label and log."""
        prefix = "Status: "
        if is_error:
            prefix = "Error: "
            logging.error(message)
        else:
            logging.info(message)
        
        if hasattr(self, 'status_label'):
            self.status_label.setText(prefix + message)
            if QApplication.instance():
                QApplication.processEvents()
        else:
            print(f"Status Label not ready. Message: {prefix + message}")

    def authenticate(self):
        # New device code flow using a POST request:
        try:
            payload = {
                "client_id": CLIENT_ID,
                "scope": " ".join(SCOPES)
            }
            headers = {"Content-Type": "application/x-www-form-urlencoded"}
            response = requests.post(DEVICE_CODE_ENDPOINT, data=payload, headers=headers)
            response.raise_for_status()  # Raises an exception for HTTP errors
            device_data = response.json()
            
            device_code = device_data.get("device_code")
            verification_uri = device_data.get("verification_uri")
            user_code = device_data.get("user_code")
            
            if not device_code or not verification_uri or not user_code:
                raise ValueError("Missing device code data in response")
            
            # Instruct the user to manually complete the device code verification.
            info_message = (f"To complete authentication, please visit {verification_uri} "
                            f"and enter the following code: {user_code}")
            logging.info(info_message)
            QMessageBox.information(self, "Device Code Authentication", info_message)
            
            # Start polling for the token using the device code.
            # You can implement a polling loop (with a timer) that POSTs to the TOKEN_ENDPOINT
            # with the device_code until you receive a token or timeout.
            self._poll_for_token(device_code)
            
        except Exception as e:
            self._update_status(f"Device code authentication error: {e}", is_error=True)
            QMessageBox.critical(self, "Device Code Auth Error", f"Error initiating device code authentication:\n{e}")
    import base64 # <--- Add this import at the top of your file

# ... (keep other imports and constants)

class JDQuoteModule(QWidget):
    # ... (keep __init__, load_tokens, save_tokens, UI setup, etc.)

    def authenticate(self):
        """Initiates the OAuth 2.0 Device Authorization Grant flow."""
        logging.info("Attempting Device Code Flow authentication...")

        # --- Make sure Client Secret is loaded ---
        actual_client_secret = os.environ.get("DEERE_CLIENT_SECRET")
        if not actual_client_secret:
            error_msg = "CRITICAL: DEERE_CLIENT_SECRET environment variable not set."
            logging.error(error_msg)
            QMessageBox.critical(self, "Configuration Error", error_msg)
            self._update_status(error_msg, is_error=True)
            return # Stop authentication if secret is missing
        # --- End Secret Check ---

        try:
            # --- Prepare Basic Auth Header ---
            # Combine client ID and secret with a colon
            client_credentials = f"{CLIENT_ID}:{actual_client_secret}"
            # Encode to bytes
            client_credentials_bytes = client_credentials.encode('utf-8')
            # Base64 encode
            base64_credentials_bytes = base64.b64encode(client_credentials_bytes)
            # Decode back to string for the header
            base64_credentials_string = base64_credentials_bytes.decode('utf-8')
            auth_header = f"Basic {base64_credentials_string}"
            logging.debug(f"Using Authorization Header: Basic [REDACTED]") # Avoid logging the full header
            # --- End Basic Auth Header ---

            payload = {
                "client_id": CLIENT_ID, # Still include client_id in payload as per original code
                "scope": " ".join(SCOPES)
            }
            headers = {
                "Content-Type": "application/x-www-form-urlencoded",
                "Authorization": auth_header # <--- Add the Basic Auth header
            }

            self._update_status("Requesting device code from John Deere...")
            response = requests.post(DEVICE_CODE_ENDPOINT, data=payload, headers=headers)

            logging.debug(f"Device code request status: {response.status_code}")
            # Log response text only on error for debugging, be mindful of sensitive info
            if response.status_code != 200:
                 logging.error(f"Device code response text: {response.text}")

            response.raise_for_status() # Raises an exception for HTTP errors (like the 401)

            device_data = response.json()
            logging.debug(f"Received device data: {device_data}") # Be careful if this contains sensitive info

            device_code = device_data.get("device_code")
            verification_uri = device_data.get("verification_uri_complete") or device_data.get("verification_uri") # Use complete if available
            user_code = device_data.get("user_code")
            expires_in = device_data.get("expires_in") # Optional: use for timeout
            interval = device_data.get("interval", 5) # Use suggested interval or default to 5

            if not device_code or not verification_uri or not user_code:
                raise ValueError("Missing device code data (device_code, verification_uri, user_code) in response")

            # --- User Interaction ---
            info_message = (f"Sign-in required.\n\n"
                            f"1. Open this URL in your browser:\n{verification_uri}\n\n"
                            f"2. Enter this code when prompted:\n{user_code}\n\n"
                            f"Waiting for you to approve...")
            logging.info(f"Instructing user: Visit {verification_uri} and enter code {user_code}")

            # Display message and optionally open browser
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Information)
            msg_box.setWindowTitle("Device Code Authentication")
            msg_box.setText(info_message)
            msg_box.setStandardButtons(QMessageBox.Ok | QMessageBox.Open)
            result = msg_box.exec_()

            if result == QMessageBox.Open:
                 webbrowser.open(verification_uri)

            self._update_status(f"Waiting for user to enter code: {user_code}")
            # --- End User Interaction ---


            # --- Start Polling ---
            # Pass interval to the polling function
            self._poll_for_token(device_code, interval)
            # --- End Polling ---

        except requests.exceptions.RequestException as req_err:
             # Handle specific request errors (network, DNS, etc.)
             error_msg = f"Network or connection error during device code request: {req_err}"
             logging.error(error_msg, exc_info=True)
             QMessageBox.critical(self, "Network Error", f"{error_msg}")
             self._update_status(error_msg, is_error=True)
        except ValueError as val_err:
             # Handle missing data in response
             error_msg = f"Invalid response from device code endpoint: {val_err}"
             logging.error(error_msg, exc_info=True)
             QMessageBox.critical(self, "API Response Error", f"{error_msg}")
             self._update_status(error_msg, is_error=True)
        except Exception as e:
            # Catch other potential errors (JSON decoding, etc.)
            error_msg = f"Device code authentication error: {e}"
            logging.error(error_msg, exc_info=True) # Log full traceback
            # Show specific error details if possible (e.g., from response)
            details = ""
            if hasattr(e, 'response') and e.response is not None:
                try:
                    details = f"\n\nDetails: {e.response.status_code} - {e.response.text}"
                except Exception:
                    details = f"\n\nStatus Code: {e.response.status_code}" # Fallback if text fails
            QMessageBox.critical(self, "Device Code Auth Error", f"{error_msg}{details}")
            self._update_status(f"Device code auth failed: {e}", is_error=True)


    # Make sure you only have ONE definition of the authenticate function
    # Remove the duplicate one if it exists


    def _poll_for_token(self, device_code, interval=5): # Add interval parameter
        """
        Start polling the token endpoint using the device_code.
        Uses the interval suggested by the authorization server if provided.
        """
        # Stop existing timer if any
        if hasattr(self, "_poll_timer") and self._poll_timer:
            self._poll_timer.stop()
            logging.debug("Stopped existing poll timer.")

        polling_interval_ms = max(interval, 2) * 1000 # Ensure interval is at least 2 seconds, convert to ms

        self._poll_timer = QTimer(self)
        # Use lambda to pass device_code correctly
        self._poll_timer.timeout.connect(lambda dc=device_code: self._attempt_token_exchange(dc))
        self._poll_timer.start(polling_interval_ms)
        logging.info(f"Started polling for token every {polling_interval_ms / 1000} seconds using device_code.")
        self._update_status("Polling for token...")

    def _attempt_token_exchange(self, device_code):
        """
        Attempt to exchange the device code for an access token.
        Handles polling logic (authorization_pending, slow_down).
        """
        logging.debug(f"Attempting token exchange for device_code: ...{device_code[-4:]}") # Log partial code

        # --- Make sure Client Secret is loaded ---
        # Needed for the token exchange step!
        actual_client_secret = os.environ.get("DEERE_CLIENT_SECRET")
        if not actual_client_secret:
            error_msg = "CRITICAL: DEERE_CLIENT_SECRET environment variable not set. Cannot exchange token."
            logging.error(error_msg)
            if hasattr(self, "_poll_timer") and self._poll_timer: # Stop polling on critical error
                 self._poll_timer.stop()
                 self._poll_timer = None
            QMessageBox.critical(self, "Configuration Error", error_msg)
            self._update_status(error_msg, is_error=True)
            return
        # --- End Secret Check ---

        try:
             # --- Prepare Basic Auth Header for Token Endpoint ---
             # This endpoint *definitely* needs client auth.
             client_credentials = f"{CLIENT_ID}:{actual_client_secret}"
             base64_credentials_string = base64.b64encode(client_credentials.encode('utf-8')).decode('utf-8')
             auth_header = f"Basic {base64_credentials_string}"
             # --- End Basic Auth ---

             payload = {
                 "client_id": CLIENT_ID, # Some providers require client_id here too
                 "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                 "device_code": device_code
             }
             headers = {
                 "Content-Type": "application/x-www-form-urlencoded",
                 "Authorization": auth_header # <--- Add the Basic Auth header
             }

             response = requests.post(TOKEN_ENDPOINT, data=payload, headers=headers)

             # --- Handle Successful Token Response ---
             if response.status_code == 200:
                 token_data = response.json()
                 self.access_token = token_data.get("access_token")
                 self.refresh_token = token_data.get("refresh_token") # Might be null if not requested/granted
                 expires_in = token_data.get("expires_in", 3600) # Default to 1 hour if missing
                 self.token_expiry = time.time() + expires_in

                 # Stop the timer
                 if hasattr(self, "_poll_timer") and self._poll_timer:
                     self._poll_timer.stop()
                     self._poll_timer = None
                     logging.debug("Stopped poll timer.")

                 self.save_tokens() # Save the new tokens immediately
                 logging.info("Device code polling successful, token received and saved.")
                 self._update_status("Authentication successful!")
                 QMessageBox.information(self, "Auth Success", "Successfully authenticated with John Deere API.")
                 # Potentially trigger other actions now that you are authenticated (e.g., load orgs)
                 self.fetch_organization_id() # Example: Try fetching org ID after auth
                 return # Exit function on success

             # --- Handle Errors and Pending Status ---
             else:
                 try:
                     data = response.json()
                     error = data.get("error")
                     error_desc = data.get("error_description", "No description provided.")
                     logging.warning(f"Token exchange non-200 response: {response.status_code} - {error} - {error_desc}")

                     if error == "authorization_pending":
                         logging.info("Authorization pending... continuing poll.")
                         self._update_status("Waiting for user approval...")
                         # Timer will fire again, no need to stop it here
                     elif error == "slow_down":
                         logging.warning("Slow down requested; increasing polling interval.")
                         self._update_status("Polling slower...")
                         if hasattr(self, "_poll_timer") and self._poll_timer:
                             current_interval_ms = self._poll_timer.interval()
                             self._poll_timer.setInterval(current_interval_ms + 5000) # Increase interval by 5s
                     elif error == "expired_token":
                         # This means the device_code itself expired
                         logging.error("Device code expired. User did not authorize in time.")
                         if hasattr(self, "_poll_timer") and self._poll_timer:
                            self._poll_timer.stop()
                            self._poll_timer = None
                         self._update_status("Authorization failed (Code Expired)", is_error=True)
                         QMessageBox.warning(self, "Authorization Expired", "The authorization request expired because it wasn't approved in time. Please try authenticating again.")
                     else: # Handle other, unexpected errors
                         logging.error(f"Unhandled error during token polling: {response.status_code} - {error} - {error_desc}")
                         if hasattr(self, "_poll_timer") and self._poll_timer:
                             self._poll_timer.stop()
                             self._poll_timer = None
                         self._update_status(f"Token polling error: {error}", is_error=True)
                         QMessageBox.critical(self, "Token Polling Error", f"Error exchanging device code for token:\n{error} - {error_desc}")

                 except json.JSONDecodeError: # Handle cases where response isn't valid JSON
                     logging.error(f"Failed to decode JSON response during token poll. Status: {response.status_code}, Response: {response.text}")
                     if hasattr(self, "_poll_timer") and self._poll_timer:
                         self._poll_timer.stop()
                         self._poll_timer = None
                     self._update_status("Token polling error (Invalid Response)", is_error=True)
                     QMessageBox.critical(self, "API Error", f"Received an invalid response from the server during token polling (Status {response.status_code}).")

        except requests.exceptions.RequestException as req_err:
            # Handle network errors during polling
            logging.error(f"Network error during token polling attempt: {req_err}", exc_info=True)
            # Optionally stop polling or just log and let it retry? For now, log and continue.
            self._update_status("Network error during polling...", is_error=True)
            # Consider adding a retry limit or stopping the timer after several network errors.
        except Exception as e:
            # Catch-all for unexpected errors during the attempt
            logging.error(f"Unexpected exception during token polling attempt: {e}", exc_info=True)
            if hasattr(self, "_poll_timer") and self._poll_timer:
                self._poll_timer.stop()
                self._poll_timer = None
            self._update_status(f"Token polling failed: {e}", is_error=True)
            QMessageBox.critical(self, "Polling Error", f"An unexpected error occurred while polling for the token:\n{e}")

    # ... (keep _refresh_token, _headers, API call methods, etc.)

    # IMPORTANT: Ensure _refresh_token also includes the Basic Auth header
    # when calling TOKEN_ENDPOINT with grant_type=refresh_token.
    # The current code for _refresh_token seems incorrect as it puts
    # client_id and client_secret in the body, not the header.

    def _refresh_token(self):
        """Refresh the access token using the refresh token"""
        logging.info("Attempting to refresh token...")
        if not self.refresh_token:
            logging.warning("No refresh token available to refresh.")
            raise ValueError("No refresh token available")

        # --- Make sure Client Secret is loaded ---
        actual_client_secret = os.environ.get("DEERE_CLIENT_SECRET")
        if not actual_client_secret:
            error_msg = "CRITICAL: DEERE_CLIENT_SECRET environment variable not set. Cannot refresh token."
            logging.error(error_msg)
            # Don't raise here, let the caller handle the failure if needed
            # but log the critical config issue.
            # We might still have a valid (short-lived) access token.
            # Or we might need to force re-auth.
            # Forcing re-auth might be safer:
            self.access_token = None
            self.refresh_token = None
            self.token_expiry = 0
            self.save_tokens() # Clear saved tokens
            raise ValueError(error_msg) # Raise to signal failure
        # --- End Secret Check ---

        try:
            # --- Prepare Basic Auth Header ---
            client_credentials = f"{CLIENT_ID}:{actual_client_secret}"
            base64_credentials_string = base64.b64encode(client_credentials.encode('utf-8')).decode('utf-8')
            auth_header = f"Basic {base64_credentials_string}"
            # --- End Basic Auth ---

            payload = {
                "grant_type": "refresh_token",
                "refresh_token": self.refresh_token,
                "scope": " ".join(SCOPES) # Request same scopes again (good practice)
                # DO NOT include client_id/secret in the payload for refresh_token grant
            }
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "Authorization": auth_header # <--- Add the Basic Auth header
            }

            response = requests.post(TOKEN_ENDPOINT, data=payload, headers=headers)
            logging.debug(f"Refresh token response status: {response.status_code}")

            if response.status_code != 200:
                 logging.error(f"Refresh token response text: {response.text}")
            response.raise_for_status() # Raise exception for non-200 status

            token_data = response.json()

            # Update tokens
            self.access_token = token_data.get("access_token")
            # Check if a *new* refresh token was issued (some providers rotate them)
            if "refresh_token" in token_data:
                self.refresh_token = token_data.get("refresh_token")
                logging.info("Received a new refresh token during refresh.")
            else:
                 # Keep the old refresh token if a new one wasn't provided
                 logging.info("Existing refresh token remains valid.")

            expires_in = token_data.get("expires_in", 3600)
            self.token_expiry = time.time() + expires_in
            logging.info(f"Token successfully refreshed. New expiry in {expires_in} seconds.")

            # Save to cache immediately
            self.save_tokens()
            self._update_status("Token refreshed successfully.")

        except requests.exceptions.HTTPError as http_err:
             logging.error(f"HTTP error refreshing token: {http_err}")
             # If refresh fails (e.g., 400 Bad Request, often means invalid refresh token),
             # clear the tokens to force re-authentication next time.
             if http_err.response.status_code in [400, 401, 403]:
                 logging.warning("Refresh token seems invalid or expired. Clearing tokens.")
                 self.access_token = None
                 self.refresh_token = None
                 self.token_expiry = 0
                 self.save_tokens() # Save cleared state
                 self._update_status("Token refresh failed. Please re-authenticate.", is_error=True)
             # Re-raise the exception so _headers knows it failed
             raise RuntimeError(f"Failed to refresh token: {http_err}") from http_err
        except Exception as e:
             logging.error(f"Unexpected error refreshing token: {e}", exc_info=True)
             # Also clear tokens on unexpected errors during refresh? Maybe safer.
             self.access_token = None
             self.refresh_token = None
             self.token_expiry = 0
             self.save_tokens()
             self._update_status("Token refresh failed unexpectedly. Please re-authenticate.", is_error=True)
             # Re-raise the exception
             raise RuntimeError(f"Unexpected error during token refresh: {e}") from e


    # ... (rest of the class)
    def _headers(self):
        """Returns necessary headers for API calls, checking authentication status first."""
        if not self.access_token:
            raise RuntimeError("Authentication required. Please click Authenticate.")
            
        if self.token_expiry <= time.time():
            if self.refresh_token:
                try:
                    self._refresh_token()
                except Exception as e:
                    logging.error(f"Failed to refresh token: {e}")
                    self.refresh_token = None
                    raise RuntimeError("Authentication token expired. Please authenticate again.")
            else:
                raise RuntimeError("Authentication token expired. Please authenticate again.")
        
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/vnd.deere.maintain-quotes.v1+json", 
            "Content-Type": "application/json"
        }

    def _exchange_code_for_token(self, code):
        """Exchange authorization code for tokens."""
        try:
            response = requests.post(
                TOKEN_ENDPOINT,
                data={
                    'grant_type': 'authorization_code',
                    'code': code,
                    'redirect_uri': 'http://localhost:9090/callback',
                    'client_id': CLIENT_ID,
                    'client_secret': CLIENT_SECRET
                },
                headers={
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            )
            response.raise_for_status()
            token_data = response.json()
            
            # Save tokens
            self.access_token = token_data.get('access_token')
            self.refresh_token = token_data.get('refresh_token')
            expires_in = token_data.get('expires_in', 3600)
            self.token_expiry = time.time() + expires_in
            
            # Save to cache
            self.save_tokens()
            
            # Update UI
            self._update_status("Authentication successful")
            QMessageBox.information(self, "Auth Success", "Successfully authenticated with John Deere API.")
        except requests.RequestException as e:
            error_msg = str(e)
            if hasattr(e, 'response') and e.response and e.response.content:
                try:
                    error_data = e.response.json()
                    error_msg = error_data.get('error_description', error_data.get('error', str(e)))
                except:
                    error_msg = e.response.text
            self._update_status(f"Token exchange error: {error_msg}", is_error=True)
            QMessageBox.critical(self, "Auth Error", f"Error exchanging code for token:\n{error_msg}")
        except Exception as e:
            self._update_status(f"Token exchange error: {e}", is_error=True)
            QMessageBox.critical(self, "Auth Error", f"Unexpected error exchanging code for token:\n{e}")

    def _handle_api_error(self, operation_name, error):
        """Logs API errors and shows messages, extracting details from response if possible."""
        logging.error(f"Error during {operation_name}: {error}", exc_info=True)
        error_details = ""
        status_code = "N/A"

        if isinstance(error, RuntimeError) and "Authentication required" in str(error):
            QMessageBox.warning(self, "Authentication Needed", str(error))
            return 
        elif isinstance(error, requests.exceptions.RequestException):
            message = f"Network error during {operation_name}:\n{error}"
            if error.response is not None:
                status_code = error.response.status_code
                try:
                    error_data = error.response.json() 
                    err_msg = error_data.get('message', error_data.get('error', {}).get('message', 'No details found.')) 
                    error_details = f"\nStatus: {status_code}\nDetails: {err_msg}"
                    logging.error(f"API Error Response Body: {json.dumps(error_data)}") 
                except json.JSONDecodeError:
                    error_details = f"\nStatus: {status_code}\nRaw Response: {error.response.text[:500]}..." 
                if status_code in [401, 403]:
                    self.access_token = None 
                    error_details += "\n\nAuthentication may have expired. Please try Authenticating again."
                    self._update_status("Authentication may be required", is_error=True)
            self.response_view.setPlainText(f"Request failed:\n{error}{error_details}")
        else: 
            message = f"Unexpected error during {operation_name}:\n{error}"
            self.response_view.setPlainText(f"Operation failed:\n{error}")

        self._update_status(f"{operation_name} failed (Status: {status_code})", is_error=True)
        QMessageBox.critical(self, f"{operation_name} Error", message + error_details)

    def fetch_organization_id(self):
        """Fetches organizations from the API and prompts user selection."""
        try:
            headers = self._headers()
            orgs_endpoint = f"{API_BASE_URL}/organizations"
            
            self._update_status("Fetching organizations...")
            response = requests.get(orgs_endpoint, headers=headers, timeout=15)
            response.raise_for_status()
            
            data = response.json()
            self.response_view.setPlainText(f"Fetched Organizations:\n{json.dumps(data, indent=2)}")
            
            orgs = data.get("values", [])
            if not orgs:
                self._update_status("No organizations found", is_error=True)
                QMessageBox.warning(self, "No Orgs Found", "No organizations found for your user account.")
                return
                
            items = [f"{org['id']} - {org.get('name', 'Unnamed Organization')}" for org in orgs]
            current_index = -1
            if self.organization_id:
                current_index = next((i for i, item in enumerate(items) if item.startswith(self.organization_id)), -1)
                
            item, ok = QInputDialog.getItem(self, "Select Organization", 
                                          "Choose organization:", items, 
                                          current=max(0, current_index), editable=False)
                                          
            if ok and item:
                selected_org_id = item.split(" - ")[0]
                if selected_org_id != self.organization_id:
                    self.organization_id = selected_org_id
                    self.save_org_id(self.organization_id)
                    self._update_status(f"Organization set: {self.organization_id}")
                else:
                    self._update_status(f"Organization confirmed: {self.organization_id}")
                self.org_id_label.setText(f"Current Org ID: {self.organization_id}")
                self.quote_list.clear()
            else:
                self._update_status("Organization selection cancelled")
                
        except Exception as e:
            self._handle_api_error("Fetch Organizations", e)

    def load_quotes(self):
        """Loads quotes for the currently selected organization ID."""
        if not self.organization_id:
            QMessageBox.warning(self, "Missing Org ID", "Please fetch and select an organization first using 'Get/Set Org ID'.")
            self._update_status("Cannot load quotes - Org ID not set", is_error=True)
            return
            
        self._update_status(f"Loading quotes for Org ID: {self.organization_id}.")
        self.quote_list.clear()
        self.quote_list.addItem("Loading.")
        if QApplication.instance():
            QApplication.processEvents()
            
        try:
            headers = self._headers()
            params = {'organizationId': self.organization_id}
            quotes_endpoint = f"{API_BASE_URL}/quotes"
            
            response = requests.get(quotes_endpoint, headers=headers, params=params, timeout=30)
            response.raise_for_status()
            
            data = response.json()
            self.quote_list.clear()
            self.response_view.setPlainText(f"Loaded Quotes (Org: {self.organization_id}):\n{json.dumps(data, indent=2)}")
            
            quotes = data.get("values", [])
            if not quotes:
                self.quote_list.addItem("No quotes found for this organization.")
                self._update_status("No quotes found")
            else:
                for quote in quotes:
                    desc = quote.get('description', '<No Description>')
                    status = quote.get('status', {}).get('value', 'Unknown')
                    quote_id = quote.get('id', '<Missing ID>')
                    display = f"{quote_id}: {desc} (Status: {status})"
                    self.quote_list.addItem(display)
                self._update_status(f"Loaded {len(quotes)} quotes")
                
        except Exception as e:
            self.quote_list.clear()
            self.quote_list.addItem("Error loading quotes.")
            self._handle_api_error("Load Quotes", e)

    def create_quote(self):
        """Creates a new quote via the API."""
        if not self.organization_id:
            QMessageBox.warning(self, "Missing Org ID", "Please fetch and select an organization first using 'Get/Set Org ID'.")
            self._update_status("Cannot create quote - Org ID not set", is_error=True)
            return
            
        text, ok = QInputDialog.getText(self, "New Quote", "Enter description for the new quote:")
        if ok and text:
            self._update_status(f"Creating quote '{text}'.")
            payload = {
                "description": text,
                "organizationId": self.organization_id
            }
            
            try:
                headers = self._headers()
                create_endpoint = f"{API_BASE_URL}/quotes"
                
                response = requests.post(create_endpoint, headers=headers, 
                                        data=json.dumps(payload), timeout=15)
                response.raise_for_status()
                
                new_quote_data = response.json()
                new_id = new_quote_data.get('id', 'N/A')
                self.response_view.setPlainText(f"Quote Created:\n{json.dumps(new_quote_data, indent=2)}")
                self._update_status(f"Quote '{text}' created (ID: {new_id})")
                QMessageBox.information(self, "Success", f"Quote created successfully!\nID: {new_id}")
                
                self.load_quotes()
                
            except Exception as e:
                self._handle_api_error(f"Create Quote '{text}'", e)

    def delete_quote(self):
        """Deletes the selected quote from the list via the API."""
        current = self.quote_list.currentItem()
        if not current or "No quotes found" in current.text() or "Loading" in current.text():
            QMessageBox.information(self, "Select Quote", "Please select a valid quote from the list to delete.")
            return
            
        try:
            quote_id_part = current.text().split(":")[0]
            quote_id = quote_id_part.strip()
            if not quote_id or quote_id == '<Missing ID>':
                raise ValueError("Invalid or missing quote ID selected.")
        except (IndexError, ValueError) as parse_err:
            QMessageBox.warning(self, "Selection Error", 
                              f"Could not parse valid quote ID from selected item: '{current.text()}'.\nError: {parse_err}")
            return
            
        confirm = QMessageBox.question(self, "Confirm Delete", 
                                     f"Are you sure you want to delete quote ID:\n{quote_id}?", 
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if confirm == QMessageBox.Yes:
            self._update_status(f"Deleting quote {quote_id}.")
            
            try:
                headers = self._headers()
                delete_endpoint = f"{API_BASE_URL}/quotes/{quote_id}"
                
                response = requests.delete(delete_endpoint, headers=headers, timeout=15)
                response.raise_for_status()
                
                status = response.status_code
                self.response_view.setPlainText(f"Quote Deleted (ID: {quote_id}, Status: {status})")
                self._update_status(f"Quote {quote_id} deleted successfully")
                QMessageBox.information(self, "Deleted", f"Quote {quote_id} deleted successfully.")
                
                self.load_quotes()
                
            except Exception as e:
                self._handle_api_error(f"Delete Quote {quote_id}", e)

    def _cleanup_auth_server(self):
        """Clean up the auth server resources."""
        if hasattr(self, 'auth_server') and self.auth_server:
            logging.info("Shutting down auth server")
            try:
                self.auth_server.shutdown()
                self.auth_server.server_close()
            except Exception as e:
                logging.error(f"Error shutting down auth server: {e}")
            self.auth_server = None
        if hasattr(self, 'server_thread') and self.server_thread and self.server_thread.is_alive():
            self.server_thread.join(1.0)
            self.server_thread = None

    def __del__(self):
        """Destructor to ensure cleanup of the authentication server."""
        try:
            self._cleanup_auth_server()
        except Exception as e:
            logging.error(f"Error during cleanup in __del__: {e}")

# --- Example of running standalone for testing (Optional) ---
if __name__ == '__main__':
    import sys
    from PyQt5.QtWidgets import QApplication, QMainWindow

    # Ensure DEERE_CLIENT_SECRET is set in environment for testing
    if not os.environ.get("DEERE_CLIENT_SECRET"):
        print("\nWARNING: DEERE_CLIENT_SECRET environment variable not set.")
    app = QApplication(sys.argv)
    # Create a simple mock main window
    class MockMainWindow(QMainWindow):
        def __init__(self):
            super().__init__()
            self.statusBar = self.statusBar()
    main_window_mock = MockMainWindow() 

    jd_module = JDQuoteModule(main_window=main_window_mock)
    jd_module.show()
    sys.exit(app.exec_())
