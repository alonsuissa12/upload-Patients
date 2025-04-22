import logging
import os
from datetime import datetime


# Create logs directory if it doesn't exist
def setup_logger():
    # Create logs directory if it doesn't exist
    if not os.path.exists('logs'):
        os.makedirs('logs')

    # Create a unique log file name with timestamp
    log_filename = f"logs/clalit_run_{datetime.now().strftime('%Y-%m-%d__time_%H-%M')}.log"

    # Configure the logger
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        handlers=[
            logging.FileHandler(log_filename,encoding="windows-1255"),  # Use windows-1255 encoding
            logging.StreamHandler()  # This will also print logs to console
        ]
    )

    logging.info("Logger initialized")
    return logging.getLogger()




