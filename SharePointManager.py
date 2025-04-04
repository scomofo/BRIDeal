import os
import requests
from urllib.parse import quote
from PyQt5.QtWidgets import QMessageBox # Import QMessageBox for error popups
import logging # Use logging
import json # For email payload

# Assuming auth.py is in the same directory or accessible via PYTHONPATH
try:
    from auth import get_access_token
except ImportError as e:
    logging.critical(f"CRITICAL ERROR: Could not import 'get_access_token' from auth.py: {e}.")
    def get_access_token():
        logging.error("ERROR: auth.get_access_token function is missing or auth.py failed to import!")
        raise RuntimeError("Authentication function not found or import failed.")

logger = logging.getLogger('AMSDeal.SharePointManager') # Get logger instance

class SharePointExcelManager:
    """
    Manages authentication and interaction with SharePoint/Graph API.
    Includes methods for Excel reading/writing and sending Email.
    """
    def __init__(self):
        """
        Initializes the manager, loading configuration directly from environment variables.
        Authentication is lazy-loaded (occurs on first API call).
        """
        logger.info("Initializing SharePoint Manager...")
        # Load configuration directly from environment variables
        self.site_id = os.getenv("SITE_ID") # Expects SITE_ID (GUID or composite)
        self.file_name = os.getenv("FILE_PATH") # Expects just the filename
        self.site_name = os.getenv("SHAREPOINT_SITE_NAME") # Optional, for logging
        self.sender_email = os.getenv("SENDER_EMAIL") # Email address to send from

        # Construct the full relative path assuming file is at root of default library
        self.full_relative_path = self.file_name
        if self.full_relative_path:
             logger.debug(f"Using relative path (from FILE_PATH env var): '{self.full_relative_path}'")

        self.graph_url = "https://graph.microsoft.com/v1.0"
        self.access_token = None # Initialize token as None
        self._file_id = None # Cache for file ID

        # Validation (Log errors but don't necessarily crash init)
        if not self.site_id: logger.error("CONFIG ERROR: SITE_ID environment variable missing.")
        else: logger.debug(f"Using SITE_ID from env: '{self.site_id}'")
        if not self.file_name: logger.error("CONFIG ERROR: FILE_PATH environment variable missing.")
        if not self.full_relative_path: logger.error("CONFIG ERROR: Path construction failed.")
        if not self.sender_email: logger.error("CONFIG ERROR: SENDER_EMAIL environment variable missing.")

        logger.info("SharePoint Manager Initialized (token will be acquired on first use).")

    def _authenticate(self):
        """Gets the access token using the function imported from auth.py."""
        if self.access_token: return True # Add expiry check later if needed
        logger.info("Attempting authentication via auth.get_access_token()...")
        try:
            self.access_token = get_access_token()
            if not self.access_token: raise ValueError("Auth function returned no token.")
            logger.info("Authentication successful (token received).")
            return True
        except Exception as e:
            logger.error(f"Authentication call failed: {e}", exc_info=True)
            self.access_token = None
            QMessageBox.critical(None, "Authentication Error", f"Failed to authenticate with SharePoint/Graph:\n{e}")
            return False

    def _get_headers(self):
        """Gets headers, attempting authentication if needed. Returns None on failure."""
        if self.access_token is None:
             logger.info("Token not found, attempting authentication now...")
             if not self._authenticate():
                 logger.error("Authentication failed when getting headers.")
                 return None
        if self.access_token:
            return { 'Authorization': 'Bearer ' + self.access_token, 'Content-Type': 'application/json' }
        else:
             logger.error("Could not get headers because access token is still None after auth attempt.")
             return None

    def _get_file_id(self, headers):
        """Helper method to get (and cache) the SharePoint file ID."""
        if self._file_id:
            logger.debug(f"Returning cached File ID: {self._file_id}")
            return self._file_id

        if not self.full_relative_path:
             logger.error("Cannot get File ID: File path is not set.")
             return None

        file_info_url = f"{self.graph_url}/sites/{self.site_id}/drive/root:/{quote(str(self.full_relative_path))}:/"
        logger.info(f"Getting File ID: GET {file_info_url}")
        try:
            response_file = requests.get(file_info_url, headers=headers)
            if response_file.status_code == 400:
                 try:
                     error_body = response_file.json()
                     if error_body.get('error', {}).get('message') == 'Invalid hostname for this tenancy':
                         raise requests.exceptions.RequestException(f"Invalid hostname (Site ID: {self.site_id})", response=response_file)
                 except ValueError: pass
            response_file.raise_for_status()
            file_info = response_file.json()
            file_id = file_info.get('id')
            if not file_id: raise ValueError("Could not get file ID from Graph API response")
            logger.info(f"Obtained File ID: {file_id}")
            self._file_id = file_id # Cache the ID
            return file_id
        except requests.exceptions.RequestException as e:
            logger.error(f"Failed to get File ID: {e}", exc_info=True)
            error_details = f"API Request Failed getting File ID:\n{e}"
            # --- Refactored error body parsing ---
            if hasattr(e, 'response') and e.response is not None:
                 logger.error(f"Response Status: {e.response.status_code}")
                 try:
                     error_body = e.response.json()
                     logger.error(f"Response Body: {error_body}")
                     ms_error = error_body.get('error', {}).get('message', '')
                     if ms_error: error_details += f"\nDetails: {ms_error}"
                 except ValueError: # Handle cases where response is not JSON
                     logger.error(f"Response Body (non-JSON): {e.response.text[:500]}")
                     error_details += f"\nResponse: {e.response.text[:200]}..."
            # --- End Refactor ---
            QMessageBox.critical(None, "SharePoint Error", error_details)
            return None
        except Exception as e:
            logger.error(f"Unexpected error getting File ID: {e}", exc_info=True)
            QMessageBox.critical(None, "SharePoint Error", f"Unexpected Error getting File ID:\n{e}")
            return None

    # --- NEW: Read Excel Sheet Method ---
    def read_excel_sheet(self, sheet_name):
        """
        Reads the used range data from a specific sheet in the Excel file.

        Args:
            sheet_name (str): The name of the worksheet to read.

        Returns:
            list[list[str]] | None: A list of lists representing the rows and cells,
                                     or None if an error occurs.
        """
        logger.info(f"Attempting to read sheet '{sheet_name}'...")
        headers = self._get_headers()
        if headers is None: return None # Auth failed

        file_id = self._get_file_id(headers)
        if file_id is None: return None # Failed to get file ID

        read_range_url = f"{self.graph_url}/sites/{self.site_id}/drive/items/{file_id}/workbook/worksheets/{quote(sheet_name)}/usedRange?$select=values"
        logger.info(f"Reading sheet data: GET {read_range_url}")

        try:
            response = requests.get(read_range_url, headers=headers)
            response.raise_for_status() # Check for HTTP errors (4xx, 5xx)
            data = response.json()
            sheet_values = data.get("values")
            if sheet_values: logger.info(f"Successfully read {len(sheet_values)} rows from sheet '{sheet_name}'."); return sheet_values
            else: logger.warning(f"Sheet '{sheet_name}' found, but no data in used range."); return []

        except requests.exceptions.RequestException as e:
             logger.error(f"Graph API request failed reading sheet '{sheet_name}': {e}", exc_info=True)
             error_details = f"API Request Failed reading sheet '{sheet_name}':\n{e}"
             # --- Refactored error body parsing ---
             if hasattr(e, 'response') and e.response is not None:
                 logger.error(f"Response Status: {e.response.status_code}")
                 try:
                     error_body = e.response.json(); logger.error(f"Response Body: {error_body}"); ms_error = error_body.get('error', {}).get('message', '');
                     if ms_error: error_details += f"\nDetails: {ms_error}"
                 except ValueError:
                     logger.error(f"Response Body (non-JSON): {e.response.text[:500]}")
                     error_details += f"\nResponse: {e.response.text[:200]}..."
                 if e.response.status_code == 404: error_details += f"\n\nVerify sheet named '{sheet_name}' exists."
             # --- End Refactor ---
             QMessageBox.critical(None, "SharePoint Read Error", error_details)
             return None
        except Exception as e:
             logger.error(f"Unexpected error reading sheet '{sheet_name}': {e}", exc_info=True)
             QMessageBox.critical(None, "SharePoint Read Error", f"Unexpected Error reading sheet '{sheet_name}':\n{e}")
             return None

    # --- send_html_email Method ---
    def send_html_email(self, recipient_list, subject, html_body):
        """Sends an email using Microsoft Graph API."""
        if not self.sender_email: logger.error("Cannot send email: SENDER_EMAIL not configured."); QMessageBox.critical(None, "Email Config Error", "Sender email not configured."); return False
        headers = self._get_headers();
        if headers is None: logger.error("Cannot send email: Auth failed."); return False
        if not recipient_list: logger.warning("Email not sent: No recipients."); QMessageBox.warning(None, "Email Error", "No recipients specified."); return False
        to_recipients_payload = [{"emailAddress": {"address": email.strip()}} for email in recipient_list if email.strip()]
        if not to_recipients_payload: logger.warning("Email not sent: No valid recipients."); QMessageBox.warning(None, "Email Error", "No valid recipients."); return False
        email_payload = { "message": { "subject": subject, "body": { "contentType": "HTML", "content": html_body }, "toRecipients": to_recipients_payload }, "saveToSentItems": "true" }
        send_mail_url = f"{self.graph_url}/users/{self.sender_email}/sendMail"; logger.info(f"Sending email via Graph API: POST {send_mail_url}")
        try:
            response = requests.post(send_mail_url, headers=headers, json=email_payload)
            if response.status_code == 202: logger.info("Graph API Send Mail request accepted (Status 202)."); return True
            else: response.raise_for_status(); logger.warning(f"Send Mail returned unexpected status: {response.status_code}"); return False
        except requests.exceptions.RequestException as e:
             logger.error(f"Microsoft Graph API Send Mail request failed: {e}", exc_info=True)
             error_details = f"API Request Failed:\n{e}"
             # --- Refactored error body parsing ---
             if hasattr(e, 'response') and e.response is not None:
                 logger.error(f"Response Status: {e.response.status_code}")
                 try:
                     error_body = e.response.json(); logger.error(f"Response Body: {error_body}"); ms_error = error_body.get('error', {}).get('message', '');
                     if ms_error: error_details += f"\nDetails: {ms_error}"
                 except ValueError:
                     logger.error(f"Response Body (non-JSON): {e.response.text[:500]}")
                     error_details += f"\nResponse: {e.response.text[:200]}..."
             # --- End Refactor ---
             QMessageBox.critical(None, "Email Send Error", error_details); return False
        except Exception as e: logger.error(f"Unexpected error during email sending: {e}", exc_info=True); QMessageBox.critical(None, "Email Send Error", f"Unexpected Error:\n{e}"); return False


    # --- update_excel_data Method ---
    def update_excel_data(self, data_to_save):
        """Updates the Excel file table on SharePoint. Expects data_to_save as list of dicts."""
        if not self.site_id: logger.error("Cannot update Excel: Site ID missing."); QMessageBox.critical(None, "Config Error", "SharePoint Site ID missing."); return False
        if not self.full_relative_path: logger.error("Cannot update Excel: File Path missing."); QMessageBox.critical(None, "Config Error", "SharePoint File Path missing."); return False
        if not data_to_save: logger.warning("No data provided to save."); return False

        logger.info(f"Attempting Graph API call to update '{self.full_relative_path}' on site '{self.site_id}' with {len(data_to_save)} rows.")
        try:
            headers = self._get_headers()
            if headers is None: return False # Auth failed

            file_id = self._get_file_id(headers) # Use helper
            if file_id is None: return False # Failed to get file ID

            excel_headers = ["Payment", "Customer", "Equipment", "Stock Number", "Amount", "Trade", "Attached to stk#", "Trade STK#", "Amount2", "Salesperson", "Email Date", "Status", "Timestamp"]
            values_payload = [[row_dict.get(h, "") for h in excel_headers] for row_dict in data_to_save]
            if not values_payload: raise ValueError("No valid rows formatted.")

            table_name = "App"
            add_rows_url = f"{self.graph_url}/sites/{self.site_id}/drive/items/{file_id}/workbook/tables/{quote(table_name)}/rows/add"
            api_payload = {"index": None, "values": values_payload}
            logger.info(f"Adding rows to table: POST {add_rows_url}")
            response = requests.post(add_rows_url, headers=headers, json=api_payload)
            response.raise_for_status()
            logger.info(f"Graph API Add Rows Response Status: {response.status_code}")
            return True

        except requests.exceptions.RequestException as e:
             logger.error(f"Microsoft Graph API request failed: {e}", exc_info=True)
             error_details = f"API Request Failed:\n{e}"
             # --- Refactored error body parsing ---
             if hasattr(e, 'response') and e.response is not None:
                 logger.error(f"Response Status: {e.response.status_code}")
                 try:
                     error_body = e.response.json(); logger.error(f"Response Body: {error_body}"); ms_error = error_body.get('error', {}).get('message', '');
                     if ms_error: error_details += f"\nDetails: {ms_error}"
                 except ValueError:
                     logger.error(f"Response Body (non-JSON): {e.response.text[:500]}")
                     error_details += f"\nResponse: {e.response.text[:200]}..."
             # Add context about potential causes
             if 'file_info_url' in locals() and hasattr(e, 'request') and e.request.url == file_info_url:
                 if hasattr(e, 'response') and e.response is not None and 'Invalid hostname' in str(e): error_details += "\n\nCheck SITE_ID in .env matches Tenant ID."
                 else: error_details += f"\n\nCheck SITE_ID and FILE_PATH in .env. Path used: '{self.full_relative_path}'."
             elif 'add_rows_url' in locals() and hasattr(e, 'request') and e.request.url == add_rows_url:
                  error_details += f"\n\nCheck Table Name ('{table_name}') exists in Excel."
             # --- End Refactor ---
             QMessageBox.critical(None, "SharePoint Error", error_details)
             return False
        except Exception as e:
             logger.error(f"Unexpected error during SharePoint update: {e}", exc_info=True)
             QMessageBox.critical(None, "SharePoint Error", f"Unexpected Error during update:\n{e}")
             return False

