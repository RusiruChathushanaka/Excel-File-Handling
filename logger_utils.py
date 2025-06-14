# logger_utils.py

import logging
import os
from datetime import datetime, timedelta
import textwrap

LOG_FOLDER = "logs"
DAYS_TO_KEEP = 30

def cleanup_old_logs():
    """Deletes log files in the log folder older than DAYS_TO_KEEP."""
    if not os.path.exists(LOG_FOLDER):
        return # No folder, nothing to clean

    print(f"Running log cleanup. Keeping logs from the last {DAYS_TO_KEEP} days.")
    cutoff_date = datetime.now() - timedelta(days=DAYS_TO_KEEP)

    for filename in os.listdir(LOG_FOLDER):
        if filename.endswith(".log"):
            try:
                # Assumes filename is in 'YYYY-MM-DD.log' format
                file_date_str = filename.split('.')[0]
                file_date = datetime.strptime(file_date_str, '%Y-%m-%d')
                
                if file_date < cutoff_date:
                    file_path = os.path.join(LOG_FOLDER, filename)
                    os.remove(file_path)
                    print(f"Deleted old log file: {filename}")
            except (ValueError, IndexError):
                # Ignore files that don't match the date format
                continue

def setup_logger(config_settings: dict):
    """
    Configures and returns a logger.

    - Creates a daily log file in the LOG_FOLDER.
    - Cleans up logs older than DAYS_TO_KEEP.
    - Logs the provided configuration settings at the start.
    """
    # 1. Create log folder if it doesn't exist
    os.makedirs(LOG_FOLDER, exist_ok=True)

    # 2. Run the cleanup utility
    cleanup_old_logs()

    # 3. Basic logger configuration
    log_filename = os.path.join(LOG_FOLDER, f"{datetime.now().strftime('%Y-%m-%d')}.log")
    logger = logging.getLogger("DataUpdater")
    logger.setLevel(logging.INFO)

    # Prevent handlers from being added multiple times
    if logger.hasHandlers():
        logger.handlers.clear()

    # 4. Create handlers (file and console)
    # File handler for writing to the log file
    file_handler = logging.FileHandler(log_filename)
    file_handler.setLevel(logging.INFO)

    # Console handler for printing to the screen
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    # 5. Create formatter and add it to handlers
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # 6. Add handlers to the logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    # 7. Log the initial configuration settings
    config_header = "--- SCRIPT RUN STARTED WITH THE FOLLOWING CONFIGURATION ---"
    config_str = "\n".join([f"{key}: {value}" for key, value in config_settings.items()])
    
    logger.info(f"\n{textwrap.indent(config_header, '  ')}\n{textwrap.indent(config_str, '  ')}\n")

    return logger