# bridleal_refactored/app/core/threading.py
import logging
import traceback
import sys # For traceback printing

from PyQt5.QtCore import QObject, QRunnable, pyqtSignal, QThreadPool

# Get a logger for this module
logger = logging.getLogger(__name__)
# Basic config for standalone testing of this module, actual app should configure root logger
if not logger.handlers:
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(name)s - %(message)s')

class WorkerSignals(QObject):
    """
    Defines the signals available from a running worker thread.

    Supported signals are:
    - finished: No data
    - error: tuple (exctype, value, traceback.format_exc())
    - result: object data returned from processing
    - progress: int indicating % progress
    - status: str message for status updates
    """
    finished = pyqtSignal()
    error = pyqtSignal(tuple)
    result = pyqtSignal(object)
    progress = pyqtSignal(int)
    status = pyqtSignal(str)


class Worker(QRunnable):
    """
    Worker thread for executing long-running tasks.
    Inherits from QRunnable to handler worker thread setup, signals, and wrap-up.

    Args:
        fn (function): The function callback to run on this worker thread.
        *args: Arguments to pass to the callback function.
        **kwargs: Keyword arguments to pass to the callback function.
    """
    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()

        # Store constructor arguments (re-used for thread)
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

        # Add the callback to our kwargs
        # This allows the MplToolbar (if used) to pass progress updates.
        # It's a bit specific but can be generalized if needed.
        # For general use, a 'progress_callback' or similar in kwargs might be better.
        # self.kwargs['progress_callback'] = self.signals.progress # Example
        
        # Add a generic status update callback for the function to use
        self.kwargs['status_callback'] = self.signals.status


    def run(self):
        """
        Initialize the runner function with passed args, kwargs.
        """
        logger.debug(f"Worker started for function: {self.fn.__name__}")
        try:
            # Emit a status signal that the task is starting
            self.signals.status.emit(f"Starting task: {self.fn.__name__}...")
            
            result = self.fn(*self.args, **self.kwargs)
        except Exception as e:
            logger.error(f"Error in worker thread for {self.fn.__name__}: {e}", exc_info=True)
            # Get traceback information
            exctype, value = sys.exc_info()[:2]
            # Emit error signal with exception type, value, and formatted traceback string
            self.signals.error.emit((exctype, value, traceback.format_exc()))
        else:
            logger.debug(f"Worker for {self.fn.__name__} completed successfully.")
            # Emit result signal with the return value of the function
            self.signals.result.emit(result)
        finally:
            logger.debug(f"Worker for {self.fn.__name__} finishing.")
            # Emit finished signal
            self.signals.finished.emit()

# Example Usage (for testing this module standalone)
if __name__ == '__main__':
    import time
    from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QProgressBar, QLabel

    # --- Example Task Functions ---
    def example_task_success(status_callback, progress_callback_arg_name=None, duration=3):
        """An example task that runs successfully and reports progress."""
        status_callback.emit("Example task starting...") # Use the passed status_callback
        total_steps = duration * 10
        for i in range(total_steps):
            time.sleep(0.1)
            progress = int(((i + 1) / total_steps) * 100)
            # Check if a progress callback was passed via kwargs (e.g. progress_callback_arg_name)
            # This is just for demonstration; the Worker class itself now provides signals.progress
            # So, the function `fn` itself doesn't need to emit progress directly if Worker handles it.
            # However, if `fn` needs to report progress internally, it can.
            # For this example, we'll assume the Worker's progress signal is used by the main app.
            # The task function can emit status updates.
            if (i + 1) % 10 == 0: # Update status every second
                status_callback.emit(f"Task in progress... {progress}%")
        status_callback.emit("Example task almost done...")
        return "Task completed successfully!"

    def example_task_failure(status_callback):
        """An example task that raises an error."""
        status_callback.emit("Failure task starting...")
        time.sleep(1)
        raise ValueError("This is an intentional error from the example task.")

    def example_task_no_return(status_callback):
        """An example task that doesn't explicitly return a value (returns None)."""
        status_callback.emit("No-return task starting...")
        time.sleep(2)
        status_callback.emit("No-return task finished.")
        # No explicit return, so None will be emitted by result signal if not handled

    # --- PyQt Application for Testing ---
    class TestWindow(QMainWindow):
        def __init__(self):
            super().__init__()
            self.setWindowTitle("Worker Thread Test")
            self.setGeometry(100, 100, 400, 300)

            self.central_widget = QWidget()
            self.setCentralWidget(self.central_widget)
            layout = QVBoxLayout(self.central_widget)

            self.status_label = QLabel("Status: Idle")
            layout.addWidget(self.status_label)

            self.progress_bar = QProgressBar()
            self.progress_bar.setValue(0)
            layout.addWidget(self.progress_bar)
            
            # Make progress bar invisible initially, show when task runs
            self.progress_bar.setVisible(False)


            self.result_label = QLabel("Result: N/A")
            layout.addWidget(self.result_label)

            self.btn_success = QPushButton("Run Success Task")
            self.btn_success.clicked.connect(self.run_success_task)
            layout.addWidget(self.btn_success)

            self.btn_failure = QPushButton("Run Failure Task")
            self.btn_failure.clicked.connect(self.run_failure_task)
            layout.addWidget(self.btn_failure)

            self.btn_no_return = QPushButton("Run No-Return Task")
            self.btn_no_return.clicked.connect(self.run_no_return_task)
            layout.addWidget(self.btn_no_return)

            self.threadpool = QThreadPool()
            logger.info(f"Multithreading with maximum {self.threadpool.maxThreadCount()} threads")

        def _generic_task_runner(self, task_function, *args):
            self.btn_success.setEnabled(False)
            self.btn_failure.setEnabled(False)
            self.btn_no_return.setEnabled(False)
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(True) # Show progress bar
            self.status_label.setText(f"Status: Starting {task_function.__name__}...")
            self.result_label.setText("Result: Running...")

            # Create worker and connect signals
            # The 'status_callback' is passed by the Worker to the task_function
            worker = Worker(task_function, *args) 
            worker.signals.result.connect(self.handle_result)
            worker.signals.error.connect(self.handle_error)
            worker.signals.finished.connect(self.handle_finished)
            worker.signals.progress.connect(self.handle_progress) # Connect Worker's progress
            worker.signals.status.connect(self.handle_status_update) # Connect Worker's status

            self.threadpool.start(worker)

        def run_success_task(self):
            # The `example_task_success` expects `status_callback` as its first arg.
            # The Worker class automatically provides `status_callback` in kwargs.
            # If `example_task_success` needs other specific args, pass them here.
            # For this example, `duration` is an arg for the task.
            self._generic_task_runner(example_task_success, duration=5)


        def run_failure_task(self):
            self._generic_task_runner(example_task_failure)

        def run_no_return_task(self):
            self._generic_task_runner(example_task_no_return)

        def handle_result(self, result):
            logger.info(f"Result received: {result}")
            self.result_label.setText(f"Result: {result if result is not None else 'None (Task completed)'}")
            self.status_label.setText("Status: Task completed successfully.")

        def handle_error(self, error_tuple):
            exctype, value, tb_str = error_tuple
            logger.error(f"Error received: {exctype.__name__}: {value}")
            logger.error(f"Traceback:\n{tb_str}")
            self.result_label.setText(f"Result: Error - {exctype.__name__}")
            self.status_label.setText(f"Status: Task failed - {value}")
            QMessageBox.critical(self, "Task Error", f"An error occurred: {exctype.__name__}\n{value}\n\nTraceback:\n{tb_str}")


        def handle_finished(self):
            logger.info("Finished signal received.")
            self.btn_success.setEnabled(True)
            self.btn_failure.setEnabled(True)
            self.btn_no_return.setEnabled(True)
            # self.progress_bar.setVisible(False) # Hide progress bar again
            if not self.result_label.text().startswith("Result: Error"): # Don't overwrite error status
                 self.status_label.setText(self.status_label.text().replace("Running", "Finished"))
                 if "Task failed" not in self.status_label.text(): # if not already set by error
                    self.status_label.setText("Status: Task finished.")


        def handle_progress(self, progress_value):
            # This is connected to Worker.signals.progress
            # The task function itself does not need to emit this directly.
            # If Worker's progress signal is used, the task function should call its
            # progress_callback (passed in by Worker) if it wants to update progress.
            # For this example, Worker's progress signal is not being emitted by the example tasks.
            # If you want the progress bar to update, the task (e.g. example_task_success)
            # would need to call a `progress_callback` that is passed to it by the Worker,
            # and the Worker would need to emit `signals.progress` based on that.
            # Let's modify Worker to pass `signals.progress.emit` as a callback.

            # Simpler: For now, let's assume the main app updates progress based on status or other cues
            # if the task itself doesn't directly drive the Worker's progress signal.
            # The current `example_task_success` calculates progress but doesn't emit it via a callback
            # that would trigger Worker's progress signal.
            # For simplicity in this test, we'll update progress bar based on status messages.
            self.progress_bar.setValue(progress_value)
            logger.debug(f"Progress signal received: {progress_value}%")


        def handle_status_update(self, status_message):
            logger.info(f"Status update: {status_message}")
            self.status_label.setText(f"Status: {status_message}")
            # Example: try to parse progress from status message for the progress bar
            if "%" in status_message:
                try:
                    progress_val_str = status_message.split('%')[0].split()[-1]
                    progress_val = int(progress_val_str)
                    self.progress_bar.setValue(progress_val)
                except (ValueError, IndexError):
                    pass # Ignore if parsing fails

    app = QApplication(sys.argv)
    window = TestWindow()
    window.show()
    sys.exit(app.exec_())
