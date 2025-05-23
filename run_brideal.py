# C:/Users/smorley/BRIDeal_refactored/run_BRIDeal.py
import sys
import os
import logging # For fallback logging if main app fails very early
import traceback # For more detailed error in final catch

# Ensure the 'app' directory (which is a sibling to this script) is discoverable.
# When this script is run from the project root, the project root is typically
# added to sys.path, making 'app' directly importable.
# This explicit addition can help in some edge cases or if the execution environment is unusual.
project_root = os.path.abspath(os.path.dirname(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

try:
    # Assuming run_application is defined in app.main
    from app.main import run_application 
except ModuleNotFoundError as e:
    # This might happen if the script is somehow not in the project root,
    # or if the 'app' directory is missing/misnamed.
    critical_logger = logging.getLogger("critical_launch_error_pre_import")
    logging.basicConfig(level=logging.CRITICAL, format='%(asctime)s - %(levelname)s - %(name)s - %(message)s')
    critical_logger.critical(f"Failed to import 'app.main.run_application'. Ensure 'run_BRIDeal.py' is in the project root ('BRIDeal_refactored') and the 'app' directory exists directly under it. Error: {e}", exc_info=True)
    critical_logger.critical(f"Current sys.path: {sys.path}")
    critical_logger.critical(f"Current working directory: {os.getcwd()}")
    critical_logger.critical(f"Project root (derived from __file__): {project_root}")
    # Attempt to show a GUI error if possible, otherwise print to stderr
    try:
        from PyQt5.QtWidgets import QApplication, QMessageBox
        if not QApplication.instance():
            _app_temp = QApplication(sys.argv) # Temporary app instance for dialog
        QMessageBox.critical(None, "Import Error", f"Could not import application components (ModuleNotFoundError: {e}).\n\nPlease ensure the application structure is correct and all dependencies are installed.\nCheck logs for details.")
    except ImportError:
        print(f"CRITICAL IMPORT ERROR: {e}. Could not load application components. PyQt5 might also be missing or not in PATH.", file=sys.stderr)
    sys.exit(1)


if __name__ == '__main__':
    try:
        run_application()
    except Exception as e:
        # Basic fallback logging for critical errors during bootstrap
        # In case the main app's logging isn't set up yet.
        logging.basicConfig(level=logging.CRITICAL, format='%(asctime)s - %(levelname)s - CRITICAL_LAUNCH_ERROR - %(message)s')
        critical_logger = logging.getLogger("critical_launch_error")
        critical_logger.critical(f"Unhandled exception during application launch: {e}", exc_info=True)
        
        # Optionally, show a simple GUI error message if PyQt5 can be imported at all
        try:
            from PyQt5.QtWidgets import QApplication, QMessageBox
            # Ensure QApplication instance exists if trying to show QMessageBox before main app init
            # This is tricky as app instance might not exist or might be partially initialized.
            if not QApplication.instance():
                 _app_temp = QApplication(sys.argv) # Create a temporary one if none exists
            QMessageBox.critical(None, "Application Launch Error", f"A critical error occurred: {e}\n\n{traceback.format_exc()}\n\nSee logs or console output for more details.")
            # print(f"CRITICAL ERROR: {e}", file=sys.stderr) # Already logged by critical_logger

        except ImportError: # PyQt5 itself might be the problem or unimportable here
            print(f"CRITICAL ERROR (PyQt5 not found for error dialog): {e}\n{traceback.format_exc()}", file=sys.stderr)
            
        sys.exit(1)
