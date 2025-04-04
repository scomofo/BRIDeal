import sys
import pyautogui
import time
import pyperclip
import csv
import os

# Faster global pause for quicker movements.
pyautogui.PAUSE = 0.2

# --- Define the image directory dynamically ---
# Get the directory where the current script (TrafficAuto.py) is located
script_dir = os.path.dirname(os.path.abspath(__file__))
# Construct the path to the 'Script Images' directory relative to the script
# Assumes 'Script Images' folder is in the SAME directory as TrafficAuto.py
# If 'Script Images' is elsewhere (e.g., a subdirectory 'assets'), adjust os.path.join
images_dir = os.path.join(script_dir, "Script Images") # Use the new BRIDeal location relative to the script

# Add a check and log the path being used
if not os.path.isdir(images_dir):
    print(f"ERROR: Image directory not found at calculated path: {images_dir}")
    # Attempt fallback to original hardcoded path? Or just raise error?
    # For now, print error and continue, pyautogui will fail later if path is wrong.
    # Alternatively, exit:
    # sys.exit(f"ERROR: Image directory not found at {images_dir}")
else:
    print(f"DEBUG: Using image directory: {images_dir}")
# --- End image directory definition ---


def get_image_path(image_name):
    """Return the full path for an image file given its name."""
    return os.path.join(images_dir, image_name)

def click_element(image_path, description, timeout=10, click=True, conf=0.8, region=None):
    """
    Locate an on-screen element using an image.
    Optionally use a region to narrow the search.
    If found and click==True, move the cursor to it and click.
    Returns the location (x, y) or None if not found within the timeout.
    """
    start_time = time.time()
    location = None
    # Check if image file exists before trying to locate
    if not os.path.exists(image_path):
        print(f"ERROR: Image file not found for '{description}': {image_path}")
        return None

    while time.time() - start_time < timeout:
        try:
            # Use region if provided.
            if region:
                location = pyautogui.locateCenterOnScreen(image_path, confidence=conf, region=region)
            else:
                location = pyautogui.locateCenterOnScreen(image_path, confidence=conf)
        except pyautogui.ImageNotFoundException:
             # Expected exception if image isn't found, just continue loop
             location = None
        except Exception as e:
            print(f"Exception while locating {description}: {e}")
            location = None # Ensure location is None on other errors

        if location:
            print(f"Found {description} at {location}.")
            if click:
                try:
                    pyautogui.moveTo(location, duration=0.2)
                    pyautogui.click()
                except Exception as click_error:
                     print(f"Error clicking {description} at {location}: {click_error}")
                     return None # Return None if click fails
            return location
        time.sleep(0.2) # Small delay before retrying

    print(f"Timeout: {description} not found using image {os.path.basename(image_path)}!")
    return None

def click_and_type(image_path, text, description, timeout=10, conf=0.8):
    """
    Locate a field on-screen using its image, click it, then type text using the clipboard.
    """
    location = click_element(image_path, description, timeout, click=True, conf=conf)
    if location:
        time.sleep(0.2) # Short pause after click before typing
        try:
            pyperclip.copy(text)
            pyautogui.hotkey("ctrl", "v")
            print(f"Pasted '{text}' into {description}.")
        except Exception as paste_error:
             print(f"Error pasting text into {description}: {paste_error}")
    else:
        # Error message already printed by click_element on failure
        # print(f"Could not locate {description} for typing.")
        pass # click_element already printed the timeout message

def click_audit():
    """
    Attempt to locate and click the Audit image using a known region.
    If detection fails after multiple attempts, fallback to fixed coordinates.
    """
    audit_image = get_image_path("audit.png")
    attempts = 0
    max_attempts = 3
    # Adjust this region (x, y, width, height) to the area where the Audit pop-over usually appears.
    # These coordinates are screen-dependent!
    audit_region = (1300, 600, 300, 200) # Example region, ADJUST AS NEEDED
    print(f"Attempting to find Audit button in region: {audit_region}")
    while attempts < max_attempts:
        loc = click_element(audit_image, "Audit", timeout=3, region=audit_region)
        if loc:
            print("Audit clicked successfully via image in region.")
            return loc
        attempts += 1
        print(f"Retrying Audit click... attempt {attempts}")
        time.sleep(0.5)

    # Fallback if image detection fails.
    # These coordinates are screen-dependent and might need adjustment!
    fallback_coords = (1387, 667) # Example coordinates, ADJUST AS NEEDED
    print(f"Audit image not reliably detected, using fallback coordinates {fallback_coords}.")
    try:
        pyautogui.moveTo(fallback_coords[0], fallback_coords[1], duration=0.2)
        pyautogui.click()
        return fallback_coords
    except Exception as fallback_click_error:
        print(f"Error clicking Audit at fallback coordinates {fallback_coords}: {fallback_click_error}")
        return None

def run_automation(stock_item):
    """
    Runs the full sequence of automation steps for a given stock item.
    Returns:
        str: A status message ("Success" or an error description).
    """
    try:
        # Step 1: Click Create Traffic Ticket.
        print("Step 1: Click Create Traffic Ticket")
        if not click_element(get_image_path("create_traffic_ticket.png"), "Create Traffic Ticket", timeout=15): # Increased timeout
            return "Failed: Could not find 'Create Traffic Ticket' button."
        time.sleep(0.5) # Allow UI to react

        # Step 2: Click Audit using image recognition (with fallback).
        print("Step 2: Click Audit")
        if not click_audit():
             # click_audit prints its own errors/fallback info
             return "Failed: Could not click 'Audit'."
        # Pause for potential pop-up or transition after clicking Audit.
        time.sleep(1.5)

        # Step 3: Click Inbound.
        print("Step 3: Click Inbound")
        if not click_element(get_image_path("inbound.png"), "Inbound", timeout=10, conf=0.6):
             return "Failed: Could not find 'Inbound' button."
        time.sleep(0.5)

        # Step 4: Enter "CONV01" in the From Customer field.
        print("Step 4: Enter 'CONV01' in From Customer field")
        # click_and_type handles the click and type
        click_and_type(get_image_path("from_customer.png"), "CONV01", "From Customer Field", timeout=10)
        # Add a check? Or assume it worked if no exception. For now, assume.
        time.sleep(0.3)

        # Step 5: Click Save after typing CONV01.
        print("Step 5: Click Save (after CONV01)")
        if not click_element(get_image_path("save.png"), "Save Button (after CONV01)", timeout=10):
             return "Failed: Could not find 'Save' button after entering customer."
        time.sleep(2)  # Wait 2 seconds for the new page/section to load

        # Step 6: Enter the stock number in the Stock Number field.
        print(f"Step 6: Enter stock number '{stock_item}'")
        click_and_type(get_image_path("stock_number.png"), stock_item, "Stock Number Field", timeout=15) # Increased timeout
        time.sleep(0.3)

        # Step 7: Click Save after typing the stock number.
        print("Step 7: Click Save (after stock number)")
        if not click_element(get_image_path("save.png"), "Save Button (after stock number)", timeout=10):
             return "Failed: Could not find 'Save' button after entering stock number."
        time.sleep(0.5) # Shorter wait maybe sufficient

        # Step 8: Click the Status: Pending dropdown and press "c" to select "comp/pay".
        print("Step 8: Click Pending dropdown and press 'c'")
        if not click_element(get_image_path("pending.png"), "Pending Dropdown", timeout=10):
             return "Failed: Could not find 'Pending' dropdown."
        time.sleep(0.3) # Wait for dropdown potentially
        pyautogui.press("c")
        print("Pressed 'c' for Comp/Pay status.")
        time.sleep(0.3)

        # Step 9: Enter "BRITRK" in the Trucker field.
        print("Step 9: Enter 'BRITRK' in Trucker field")
        click_and_type(get_image_path("trucker.png"), "BRITRK", "Trucker Field", timeout=10)
        time.sleep(0.3)

        # Step 10: Enter "249" in the Salesperson field.
        print("Step 10: Enter '249' in Salesperson field")
        click_and_type(get_image_path("salesperson.png"), "249", "Salesperson Field", timeout=10)
        time.sleep(0.3)

        # Step 11: Click Save and Exit.
        print("Step 11: Click Save and Exit")
        if not click_element(get_image_path("save_and_exit.png"), "Save and Exit Button", timeout=10):
             return "Failed: Could not find 'Save and Exit' button."
        time.sleep(1.0)  # Increased pause after Save and Exit for potential dialogs

        # Step 12: Click Shipping Charge Save.
        print("Step 12: Click Shipping Charge Save")
        # This might pop up, might not. Use a shorter timeout.
        if not click_element(get_image_path("shipping_charge_save.png"), "Shipping Charge Save Button", timeout=5):
             print("Info: 'Shipping Charge Save' button not found (might not have appeared). Continuing...")
             # This might be okay, maybe it doesn't always appear.
        else:
             time.sleep(0.5) # Pause if it was clicked

        # Step 13: Click No button (likely related to printing).
        print("Step 13: Click No button")
         # This might pop up, might not. Use a shorter timeout.
        if not click_element(get_image_path("no.png"), "No Button (print prompt?)", timeout=5):
             print("Info: 'No' button not found (might not have appeared).")
             # This might be okay.
        else:
             time.sleep(0.5) # Pause if it was clicked

        success_message = f"Automation for stock '{stock_item}' completed."
        print(success_message)
        return "Success" # Return success status

    except Exception as e:
        # Catch any unexpected errors during the process
        error_message = f"ERROR during automation for stock '{stock_item}': {e}"
        print(error_message)
        import traceback
        traceback.print_exc() # Print full traceback for debugging
        return f"Failed: {e}" # Return error status

if __name__ == "__main__":
    # Check if stock numbers are provided as command-line arguments.
    if len(sys.argv) > 1:
        # Use provided arguments as the stock list (skip sys.argv[0]).
        stock_list = sys.argv[1:]
    else:
        # Read the CSV file "stocks.csv" if no command-line arguments are provided.
        stock_list = []
        stocks_csv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'stocks.csv') # Assume stocks.csv is in same dir
        print(f"No command line args, attempting to read from: {stocks_csv_path}")
        try:
            with open(stocks_csv_path, newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                stock_list = [row[0].strip() for row in reader if row and row[0].strip()] # Ensure row not empty and first cell not empty
            # Skip header if present.
            if stock_list and stock_list[0].lower() == "stock number":
                stock_list = stock_list[1:]
            print(f"Read {len(stock_list)} stock numbers from {stocks_csv_path}")
        except FileNotFoundError:
             print(f"Error: {stocks_csv_path} not found.")
             sys.exit(1)
        except Exception as e:
            print(f"Error reading {stocks_csv_path}: {e}")
            sys.exit(1)

    if not stock_list:
        print("No stock items found to process. Exiting.")
        sys.exit(1)

    # Loop through each stock item and run the automation steps.
    results_summary = []
    for stock in stock_list:
        print(f"\n--- Starting automation for stock: {stock} ---")
        status = run_automation(stock)
        results_summary.append(f"{stock}: {status}")
        print(f"--- Finished automation for stock: {stock} | Status: {status} ---")
        time.sleep(1)  # Optional: brief pause between iterations

    print("\n--- All automation tasks complete ---")
    print("Summary:")
    for result in results_summary:
        print(f"- {result}")
