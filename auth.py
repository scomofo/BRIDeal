import os
from dotenv import load_dotenv
import msal
import sys # Import sys for error stream

# --- Define Function to Get Token ---
def get_access_token():
    """
    Authenticates using MSAL Confidential Client Flow and returns the access token.

    Returns:
        str: The access token if successful.
        None: If authentication fails.
    """
    # Load environment variables within the function or ensure they are pre-loaded
    load_dotenv() # Load .env each time the function is called

    # *** Use AZURE_ prefixed variable names to match config.py and likely .env content ***
    tenant_id = os.getenv("AZURE_TENANT_ID")
    client_id = os.getenv("AZURE_CLIENT_ID")
    client_secret = os.getenv("AZURE_CLIENT_SECRET")

    # Validate required environment variables
    required_vars = ["AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET"]
    missing_vars = [var for var in required_vars if not os.getenv(var)]

    if missing_vars:
        print(f"ERROR: Missing required environment variables ({', '.join(missing_vars)}) for authentication.", file=sys.stderr)
        return None

    # Proceed with authentication using the loaded variables
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scope = ["https://graph.microsoft.com/.default"]
    access_token = None # Initialize token as None

    print("Attempting to acquire MSAL token...") # Log attempt

    try:
        app = msal.ConfidentialClientApplication(
            client_id,
            authority=authority,
            client_credential=client_secret
            # Consider adding token cache persistence here for efficiency
            # token_cache=msal.SerializableTokenCache()
        )

        # Simpler approach without cache for now: always acquire new token
        result = app.acquire_token_for_client(scopes=scope)


        if "access_token" in result:
            access_token = result['access_token']
            print("Successfully acquired access token!") # Confirmation log
            # print(f"Token starts with: {access_token[:10]}...") # Optional: Log partial token for verification
        else:
            error_desc = result.get("error_description", "No error description provided.")
            error_code = result.get("error", "N/A")
            print(f"ERROR: Failed to acquire access token.", file=sys.stderr)
            print(f"  Error Code: {error_code}", file=sys.stderr)
            print(f"  Description: {error_desc}", file=sys.stderr)
            # Optionally log more details if available:
            # print(f"  Full Error Response: {result}", file=sys.stderr)

    except Exception as e:
        print(f"ERROR: An exception occurred during MSAL authentication: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr) # Print traceback to stderr for debugging

    return access_token

# --- Optional: Standalone test ---
# This block allows running 'python auth.py' directly to test authentication
if __name__ == "__main__":
    print("Running auth.py standalone for testing...")
    # Ensure .env is loaded for standalone test
    if not (os.getenv("AZURE_TENANT_ID") and os.getenv("AZURE_CLIENT_ID") and os.getenv("AZURE_CLIENT_SECRET")):
         print("Loading .env for standalone test...")
         load_dotenv()

    token = get_access_token()
    if token:
        print("\nStandalone Test Result: SUCCESS")
        # print(f"Access Token Received (first 10 chars): {token[:10]}...")
    else:
        print("\nStandalone Test Result: FAILED")

