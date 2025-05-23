"""
SharePoint Manager module for handling SharePoint operations.
Provides functionality to interact with SharePoint files via Microsoft Graph API.
"""

import os
import requests
import pandas as pd
from datetime import datetime
import io
import traceback
import time
import json
import logging
logger = logging.getLogger(__name__)
# ... other imports ...
from .auth import get_access_token
from dotenv import load_dotenv
# *** Use RELATIVE import since auth.py is in the same 'modules' directory ***
from .auth import get_access_token

# Load environment variables if not already loaded
if 'SHAREPOINT_SITE_ID' not in os.environ:
    try:
        # Assume .env is in the parent directory (where main.py runs)
        dotenv_path = os.path.join(os.path.dirname(__file__), '..', '.env')
        if os.path.exists(dotenv_path):
            load_dotenv(dotenv_path=dotenv_path)
            print(f"Loaded environment variables from {dotenv_path}")
        else:
             print(f"Warning: .env file not found at {dotenv_path}")
    except Exception as e:
        print(f"Warning: Could not load .env file: {e}")

class SharePointExcelManager:
    

    def __init__(self):
        """Initialize the SharePoint Excel Manager."""
        print("Initializing SharePointExcelManager...")
        logger.info("Initializing SharePointExcelManager...")
        self.is_operational = False  # Initialize the attribute HERE

        # Get SharePoint and file information from environment variables
        self.site_id = os.environ.get('SHAREPOINT_SITE_ID')
        self.excel_file_path = os.environ.get('FILE_PATH')
        self.sender_email = os.environ.get('SENDER_EMAIL', '')
        # ... other attributes like site_name, sharepoint_url ...

        self.local_backup_dir = os.path.join(os.path.expanduser("~"), "brideal_sp_backups")
        # ... (backup dir creation) ...

        missing_vars = []
        required_env_vars = ['SHAREPOINT_SITE_ID', 'FILE_PATH']
        for var_name in required_env_vars:
            if not os.environ.get(var_name):
                missing_vars.append(var_name)

        if missing_vars:
            logger.error(f"SharePointExcelManager: Missing required environment variables: {', '.join(missing_vars)}. Manager will not be operational.")
            self.access_token = None
            return  # self.is_operational remains False

        try:
            import openpyxl
            logger.debug("openpyxl is installed and available.")
        except ImportError:
            logger.error("Required dependency 'openpyxl' is missing. SharePointExcelManager will not be operational.")
            self.access_token = None
            return # self.is_operational remains False

        self.graph_base_url = "https://graph.microsoft.com/v1.0"

        try:
            self.access_token = get_access_token()
            if not self.access_token:
                logger.error("Failed to acquire access token during initialization. SharePoint operations will fail.")
                # self.is_operational remains False
            else:
                logger.info("Successfully acquired initial access token for SharePoint operations.")
                self.is_operational = True  # Set to True on success
        except Exception as auth_err:
             logger.error(f"Exception during initial token acquisition: {auth_err}", exc_info=True)
             self.access_token = None
             # self.is_operational remains False
        # Get SharePoint and file information from environment variables
        self.site_id = os.environ.get('SHAREPOINT_SITE_ID')
        self.site_name = os.environ.get('SHAREPOINT_SITE_NAME', '')
        self.sharepoint_url = os.environ.get('SHAREPOINT_URL', '')
        self.excel_file_path = os.environ.get('FILE_PATH')  # Path to Excel file in SharePoint
        self.sender_email = os.environ.get('SENDER_EMAIL', '')

        # Set up local backup directory
        self.local_backup_dir = os.path.join(os.path.expanduser("~"), "ams_backup")
        if not os.path.exists(self.local_backup_dir):
            try:
                os.makedirs(self.local_backup_dir)
                print(f"Created local backup directory: {self.local_backup_dir}")
            except Exception as e:
                print(f"Warning: Could not create backup directory: {e}")

        # Validate required environment variables
        missing_vars = []
        for var_name in ['SHAREPOINT_SITE_ID', 'FILE_PATH']:
            if not os.environ.get(var_name):
                missing_vars.append(var_name)

        if missing_vars:
            print(f"ERROR: Missing required environment variables: {', '.join(missing_vars)}")
        else:
            print(f"SharePoint config loaded: Site ID: {self.site_id}, File: {self.excel_file_path}")

        # Check for required dependencies
        try:
            import openpyxl
            print("openpyxl is installed and available.")
        except ImportError:
            print("ERROR: Required dependency 'openpyxl' is missing. Please install with: pip install openpyxl")

        # Microsoft Graph API endpoints
        self.graph_base_url = "https://graph.microsoft.com/v1.0"

        # Authenticate and get access token
        try:
            self.access_token = get_access_token() # This now uses the function imported relatively
            if not self.access_token:
                print("ERROR: Failed to acquire access token. SharePoint operations will fail.")
            else:
                print("Successfully acquired access token for SharePoint operations.")
        except Exception as auth_err:
             print(f"ERROR: Exception during initial token acquisition: {auth_err}")
             self.access_token = None


    def _get_headers(self):
        """
        Generate headers for Graph API requests.

        Returns:
            dict: Request headers with access token.
        """
        # Refresh token if needed (or if initial acquisition failed)
        if not self.access_token:
            try:
                print("Attempting to refresh/acquire access token...")
                self.access_token = get_access_token()
                if not self.access_token:
                    print("ERROR: Failed to refresh/acquire access token.")
                    return None # Return None if token cannot be obtained
            except Exception as auth_err:
                print(f"ERROR: Exception during token refresh/acquisition: {auth_err}")
                self.access_token = None
                return None # Return None

        return {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }

    def get_excel_file_info(self):
        """
        Get information about the Excel file in SharePoint.

        Returns:
            dict: File information including ID, webUrl, etc. if found, None otherwise.
        """
        headers = self._get_headers()
        if headers is None:
            print("ERROR: Cannot get file info without valid authentication headers.")
            return None

        try:
            # Parse the file path
            file_path = self.excel_file_path.strip('/')
            print(f"Looking for file: {file_path}")

            # Try direct access first
            url = f"{self.graph_base_url}/sites/{self.site_id}/drive/root:/{file_path}"

            response = requests.get(url, headers=headers) # Use headers potentially refreshed
            if response.status_code == 200:
                file_data = response.json()
                print(f"Found file with ID: {file_data.get('id')}")
                return file_data

            # If direct access fails, try searching for the file
            parts = file_path.split('/')

            if len(parts) == 1:
                # File is in root folder
                file_name = parts[0]
                url = f"{self.graph_base_url}/sites/{self.site_id}/drive/root/children"
            else:
                # File is in a subfolder
                folder_path = '/'.join(parts[:-1])
                file_name = parts[-1]
                url = f"{self.graph_base_url}/sites/{self.site_id}/drive/root:/{folder_path}:/children"

            print(f"Searching for {file_name} in folder...")
            response = requests.get(url, headers=headers) # Use headers
            response.raise_for_status()

            for item in response.json().get('value', []):
                if item.get('name') == file_name:
                    print(f"Found file with ID: {item.get('id')}")
                    return item

            print(f"ERROR: Excel file '{file_name}' not found in specified path.")
            return None

        except Exception as e:
            print(f"ERROR: Failed to get Excel file info: {e}")
            traceback.print_exc()
            return None

    def get_excel_data(self, sheet_name=None):
        """
        Get current data from the Excel file.

        Args:
            sheet_name (str or int, optional): The name or index of the sheet to read.
                                               Defaults to None (reads the first sheet if 0 is not specified).
        Returns:
            pd.DataFrame: DataFrame containing the Excel data if successful, None otherwise.
        """
        headers = self._get_headers()
        if headers is None:
            print("ERROR: Cannot get excel data without valid authentication headers.")
            return None

        try:
            # Check for openpyxl before proceeding
            try:
                import openpyxl
            except ImportError:
                print("ERROR: Cannot read Excel file - openpyxl package is missing.")
                return None

            file_info = self.get_excel_file_info()
            if not file_info:
                return None

            file_id = file_info.get('id')

            # Get the Excel file content
            url = f"{self.graph_base_url}/sites/{self.site_id}/drive/items/{file_id}/content"
            print(f"Downloading Excel file content...")
            # Use headers without Content-Type for download
            download_headers = {'Authorization': headers['Authorization']}
            response = requests.get(url, headers=download_headers)
            response.raise_for_status()

            # Read Excel content into DataFrame
            excel_data = io.BytesIO(response.content)
            try:
                current_sheet_target = sheet_name if sheet_name is not None else 0 # Default to first sheet (index 0) if None
                if isinstance(sheet_name, str):
                    print(f"Parsing Excel data (Sheet: {sheet_name})...")
                else: # Could be int or None
                    print(f"Parsing Excel data (Sheet index: {current_sheet_target if sheet_name is not None else 'first sheet (default)'})...")

                df = pd.read_excel(excel_data, engine='openpyxl', sheet_name=current_sheet_target)
                print(f"Successfully read Excel sheet with {len(df)} rows and {len(df.columns)} columns.")
                return df
            except Exception as ex:
                print(f"ERROR in pandas read_excel (Target sheet: {sheet_name}): {ex}")
                traceback.print_exc()
                return None

        except Exception as e:
            print(f"ERROR: Failed to get Excel data: {e}")
            traceback.print_exc()
            return None

    def _update_excel_file_direct(self, updated_df, file_id, max_retries=3, retry_delay=2):
        """
        Update the Excel file directly with new data, with retry logic.

        Args:
            updated_df (pd.DataFrame): DataFrame with updated data.
            file_id (str): ID of the file to update.
            max_retries (int): Maximum number of retries.
            retry_delay (int): Delay in seconds between retries.

        Returns:
            bool: True if successful, False otherwise.
        """
        for attempt in range(max_retries):
            headers = self._get_headers()
            if headers is None:
                print(f"ERROR (Direct Update, attempt {attempt+1}): Cannot update without valid authentication headers.")
                if attempt < max_retries - 1: time.sleep(retry_delay); continue
                else: return False

            try:
                # Convert DataFrame to Excel bytes
                print(f"Converting DataFrame to Excel format (attempt {attempt+1}/{max_retries})...")
                excel_bytes = io.BytesIO()
                updated_df.to_excel(excel_bytes, index=False, engine='openpyxl')
                excel_bytes.seek(0)
                excel_content = excel_bytes.getvalue()

                # Upload the updated file
                url = f"{self.graph_base_url}/sites/{self.site_id}/drive/items/{file_id}/content"
                # Don't set Content-Type for file uploads, use only Authorization
                upload_headers = {'Authorization': headers['Authorization']}

                print(f"Uploading updated Excel file ({len(excel_content)} bytes)...")
                response = requests.put(url, headers=upload_headers, data=excel_content)

                if response.status_code in (200, 201):
                    print(f"Successfully updated Excel file.")
                    return True
                elif response.status_code == 423:  # Locked resource
                    print(f"WARNING: File is locked (attempt {attempt+1}/{max_retries}). Waiting before retry...")
                    time.sleep(retry_delay)
                    continue
                else:
                    print(f"ERROR: Failed to update Excel file. Status code: {response.status_code}")
                    print(f"Response: {response.text}")
                    if attempt < max_retries - 1:
                        print(f"Retrying in {retry_delay} seconds...")
                        time.sleep(retry_delay)
                    else:
                        return False

            except Exception as e:
                print(f"ERROR: Failed to update Excel file (attempt {attempt+1}/{max_retries}): {e}")
                traceback.print_exc()
                if attempt < max_retries - 1:
                    print(f"Retrying in {retry_delay} seconds...")
                    time.sleep(retry_delay)
                else:
                    return False

        return False  # All retries failed


    def _update_excel_via_session(self, updated_df, file_info, target_sheet_name=None):
        """
        Update Excel using the Excel session API, appending ALL provided rows to a specified sheet.

        Args:
            updated_df (pd.DataFrame): DataFrame with ALL new rows to append.
            file_info (dict): File information including ID and parentReference.
            target_sheet_name (str, optional): The name of the sheet to append to. 
                                             If None, uses the first sheet.
        Returns:
            bool: True if successful, False otherwise.
        """
        session_headers = self._get_headers()
        if session_headers is None:
            print("ERROR (Session Update): Cannot update without valid authentication headers.")
            return False

        session_id = None
        drive_id = file_info.get('parentReference', {}).get('driveId')
        file_id = file_info.get('id')

        try:
            print("Attempting to update Excel via session API (works with locked files)...")

            if not file_id or not drive_id:
                print("ERROR: Missing file ID or drive ID for Excel session API.")
                return False

            if updated_df.empty:
                print("WARNING: Updated DataFrame is empty, nothing to append.")
                return True

            num_cols = len(updated_df.columns)
            if num_cols == 0:
                print("ERROR: DataFrame has no columns.")
                return False

            session_url = f"{self.graph_base_url}/drives/{drive_id}/items/{file_id}/workbook/createSession"
            session_data = {"persistChanges": True}
            session_response = requests.post(session_url, headers=session_headers, json=session_data)

            if session_response.status_code != 201:
                print(f"ERROR: Failed to create Excel session. Status: {session_response.status_code}, Resp: {session_response.text}")
                return False
            session_id = session_response.json().get('id')
            print(f"Created Excel session with ID: {session_id}")
            session_headers['Workbook-Session-Id'] = session_id

            worksheet_name_to_use = target_sheet_name
            if not worksheet_name_to_use:
                worksheets_url = f"{self.graph_base_url}/drives/{drive_id}/items/{file_id}/workbook/worksheets"
                worksheets_response = requests.get(worksheets_url, headers=session_headers)
                if worksheets_response.status_code != 200:
                    raise Exception(f"Failed to get worksheets. Status: {worksheets_response.status_code}")
                worksheets = worksheets_response.json().get('value', [])
                if not worksheets:
                    raise Exception("No worksheets found in the Excel file.")
                worksheet_name_to_use = worksheets[0].get('name') # Default to first sheet
            
            print(f"Using worksheet: {worksheet_name_to_use}")

            range_url = f"{self.graph_base_url}/drives/{drive_id}/items/{file_id}/workbook/worksheets('{worksheet_name_to_use}')/usedRange(valuesOnly=true)"
            range_response = requests.get(range_url, headers=session_headers)
            if range_response.status_code != 200:
                raise Exception(f"Failed to get used range. Status: {range_response.status_code}, Resp: {range_response.text}")
            
            range_data = range_response.json()
            current_row_count = range_data.get('rowCount', 0)
            # If the sheet is completely empty (or only has headers and Graph API reports 0 rows for usedRange data),
            # we need to check if there are headers.
            # A common case is a sheet with headers but no data rows. Graph API might return rowCount=0 or 1.
            # If rowCount is 0 (or 1 and it's just a header), we start at row 1 (or 2).
            # This logic might need refinement based on actual Graph API behavior for empty/header-only sheets.
            # For simplicity, if usedRange gives 0, we assume we start appending at the first actual data row (A1 or A2 if A1 is header).
            # If usedRange shows rows, we append after them.
            # Let's assume if current_row_count is 0, it's an empty sheet, and we start at row 1.
            # If there are headers, and pd.read_excel read them, current_row_count would be 0 for the data portion.
            # We're appending data, so if current_row_count is 0 from usedRange, start at the actual first row (1).
            # If current_row_count is >0, it means there's data, so append after.
            
            # If the first row of updated_df matches the headers of an existing sheet, this implies we are writing data rows.
            # If current_row_count is 0, it means the sheet is empty or Graph API isn't counting a header-only row.
            # We will start appending at what Excel considers the next available row.
            start_row_excel = current_row_count + 1 # Excel row number (1-based)

            print(f"Current used range in '{worksheet_name_to_use}' has {current_row_count} rows. Appending will start at Excel row {start_row_excel}.")

            end_col_letter = self._col_num_to_letter(num_cols)

            for index, row_data in updated_df.iterrows():
                row_values_list = [row_data.tolist()]
                current_excel_row_to_add = start_row_excel + index

                range_address = f"{worksheet_name_to_use}!A{current_excel_row_to_add}:{end_col_letter}{current_excel_row_to_add}"
                print(f"  Appending row {index+1}/{len(updated_df)} to range: {range_address}")

                update_url = f"{self.graph_base_url}/drives/{drive_id}/items/{file_id}/workbook/worksheets('{worksheet_name_to_use}')/range(address='{range_address}')"
                update_payload = { "values": row_values_list }

                update_response = requests.patch(update_url, headers=session_headers, json=update_payload)
                if update_response.status_code != 200:
                    raise Exception(f"Failed to update row {index+1}. Status: {update_response.status_code}, Resp: {update_response.text}")

            print(f"Successfully appended {len(updated_df)} rows to '{worksheet_name_to_use}' via session API.")
            return True

        except Exception as e:
            print(f"ERROR: Failed during Excel session update: {e}")
            return False
        finally:
            if session_id and drive_id and file_id:
                try:
                    print("Attempting to close session...")
                    close_url = f"{self.graph_base_url}/drives/{drive_id}/items/{file_id}/workbook/closeSession"
                    base_headers_for_close = self._get_headers() 
                    if base_headers_for_close:
                        closing_headers = base_headers_for_close.copy()
                        closing_headers.pop('Content-Type', None) 
                        closing_headers['Workbook-Session-Id'] = session_id
                        close_response = requests.post(close_url, headers=closing_headers)
                        if close_response.status_code == 204: print("Closed session successfully.")
                        else: print(f"Warning: Failed to close session. Status: {close_response.status_code}, Resp: {close_response.text}")
                    else: print("Warning: Could not get headers to close session.")
                except Exception as close_e: print(f"Error closing session: {close_e}")

    def _col_num_to_letter(self, n):
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    def _save_local_backup(self, data, filename=None):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename_base = filename or f"ams_backup_{timestamp}"
            csv_path = os.path.join(self.local_backup_dir, f"{filename_base}.csv")
            df = pd.DataFrame(data)
            df.to_csv(csv_path, index=False)
            json_path = os.path.join(self.local_backup_dir, f"{filename_base}.json")
            with open(json_path, 'w') as f:
                json.dump(data, f, indent=2)
            print(f"Local backup saved to: {csv_path} and {json_path}")
            return csv_path
        except Exception as e:
            print(f"ERROR: Failed to save local backup: {e}")
            traceback.print_exc()
            return None

    def update_excel_data(self, new_data, target_sheet_name_for_append=None):
        """
        Add new rows to a specific Excel sheet with fallback options.
        If target_sheet_name_for_append is None, direct update rewrites the first sheet.
        Session update will use target_sheet_name_for_append or default to the first sheet.
        """
        try:
            if not new_data:
                print("ERROR: No data provided to update Excel file.")
                return False
            print(f"Updating Excel with {len(new_data)} new rows. Target sheet for append: {target_sheet_name_for_append if target_sheet_name_for_append else 'first sheet'}")

            backup_path = self._save_local_backup(new_data)
            if backup_path is None:
                print("CRITICAL ERROR: Failed to save local backup. Aborting update.")
                return False

            file_info = self.get_excel_file_info()
            if not file_info:
                print("ERROR: Failed to get Excel file info. Update aborted. Backup: {backup_path}")
                return False
            file_id = file_info.get('id')

            # Attempt Session API append first
            print("Attempting update via Session API (append)...")
            # For session append, we need to ensure `new_data` is structured as rows for the target sheet.
            # The `_update_excel_via_session` expects a DataFrame of new rows.
            df_for_session_append = pd.DataFrame(new_data)

            if self._update_excel_via_session(df_for_session_append, file_info, target_sheet_name=target_sheet_name_for_append):
                print("Update successful via Session API.")
                return True
            else:
                print("Session API append failed. Falling back to direct update (full rewrite of the first sheet or specified sheet if logic is added)...")

            # Fallback to Direct Update (Full Rewrite of the FIRST sheet usually)
            # This part needs to be clear about which sheet it's rewriting.
            # If target_sheet_name_for_append was provided, should direct update try to update only that sheet?
            # Current _get_excel_data and _update_excel_file_direct work on a single DataFrame (typically the first sheet if no name given).
            # For simplicity, direct update will read the first sheet, append, and write back the first sheet.
            # If a more sophisticated multi-sheet update is needed, _get_excel_data and _update_excel_file_direct would need changes.
            
            print("Attempting update via direct file upload (rewriting the first sheet)...")
            # Read the first sheet (index 0) for current data
            current_df = self._get_excel_data(sheet_name=0)
            if current_df is None:
                print("ERROR: Failed to get current Excel data (first sheet) for direct update. Update aborted. Backup: {backup_path}")
                return False

            # Convert new data to DataFrame and align columns to the first sheet's structure
            new_df_to_append = pd.DataFrame(new_data)
            # Ensure new_df_to_append has the same columns as current_df (first sheet) for concat
            # This might not be what's intended if new_data is for a *different* sheet structure
            # than the first one. This assumes new_data is compatible with the first sheet.
            aligned_new_df = pd.DataFrame(columns=current_df.columns) # Create empty df with target columns
            for record in new_data: # Iterate through new_data (list of dicts)
                aligned_new_df = pd.concat([aligned_new_df, pd.DataFrame([record])], ignore_index=True)
            
            # Ensure all columns from current_df exist in aligned_new_df, fill with NaN if not
            for col in current_df.columns:
                if col not in aligned_new_df.columns:
                    aligned_new_df[col] = pd.NA
            aligned_new_df = aligned_new_df[current_df.columns] # Reorder and select columns to match current_df

            updated_df = pd.concat([current_df, aligned_new_df], ignore_index=True)
            print(f"Updated DataFrame for direct upload (first sheet) has {len(updated_df)} rows.")

            if self._update_excel_file_direct(updated_df, file_id): # This writes updated_df back, effectively rewriting the file with this content
                print("Update successful via Direct Upload (first sheet rewritten).")
                return True

            print(f"WARNING: Both Session API and Direct Upload failed. Data saved locally at {backup_path}")
            return False

        except Exception as e:
            print(f"ERROR: Unhandled exception in update_excel_data: {e}")
            traceback.print_exc()
            if 'backup_path' not in locals() or backup_path is None and 'new_data' in locals():
                 try: backup_path_emergency = self._save_local_backup(new_data); print(f"Emergency backup: {backup_path_emergency}")
                 except: pass
            return False

    def send_html_email(self, recipients, subject, html_body):
        headers = self._get_headers()
        if headers is None:
            print("ERROR: Cannot send email without valid authentication headers.")
            return False
        try:
            if not self.sender_email:
                print("ERROR: Sender email not specified in environment variables.")
                return False
            if not recipients:
                print("ERROR: No recipients specified for email.")
                return False
            email_message = {
                "message": {
                    "subject": subject,
                    "body": {"contentType": "HTML", "content": html_body},
                    "toRecipients": [{"emailAddress": {"address": email}} for email in recipients]
                }
            }
            url = f"{self.graph_base_url}/users/{self.sender_email}/sendMail"
            response = requests.post(url, headers=headers, json=email_message)
            if response.status_code == 202:
                print(f"Successfully sent email to {len(recipients)} recipients.")
                return True
            else:
                print(f"ERROR: Failed to send email. Status: {response.status_code}, Resp: {response.text}")
                return False
        except Exception as e:
            print(f"ERROR: Failed to send email: {e}")
            traceback.print_exc()
            return False

if __name__ == "__main__":
    print("Testing SharePoint Excel Manager...")
    manager = SharePointExcelManager()
    file_info = manager._get_excel_file_info()
    if file_info:
        print(f"Successfully found file: {file_info.get('name')}")
        # Test reading specific sheet
        # df_sheet = manager._get_excel_data(sheet_name="Used AMS") # Example
        # if df_sheet is not None:
        #     print(f"Successfully read 'Used AMS' sheet with {len(df_sheet)} rows.")
        # else:
        #     print("Failed to read 'Used AMS' sheet.")

        # Test reading first sheet (default behavior if sheet_name is None or 0)
        df_first = manager._get_excel_data(sheet_name=0)
        if df_first is not None:
            print(f"Successfully read first sheet with {len(df_first)} rows and {len(df_first.columns)} columns.")
            print(f"Column names: {list(df_first.columns)}")
        else:
            print("Failed to read first sheet.")
    else:
        print("Failed to get file info")