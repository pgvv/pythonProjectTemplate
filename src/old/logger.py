# --------------------------- IMPORTS ---------------------------#
from datetime import datetime
import os
import logging

# --------------------------- CONSTANTS ---------------------------#
FILENAME = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"


# --------------------------- FUNCTIONS ---------------------------#
# --------------------------- MAIN CODE ---------------------------#
def get_logger(
    name: str = "default",
    log_file: str = FILENAME,
    level=logging.INFO,
) -> logging.Logger:
    formatter = logging.Formatter(
        "[%(asctime)s] - %(module)s - %(levelname)s - %(lineno)d - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    logger = logging.getLogger(name)
    logger.setLevel(level)

    # Prevent duplicate handlers
    if not logger.handlers:
        log_dir = os.path.join(os.getcwd(), "logs")
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, log_file)

        handler = logging.FileHandler(log_path)
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    logger.propagate = False  # Prevent duplicate logs to parent logger

    # Always add console handler for visibility
    # console = logging.StreamHandler()
    # console.setFormatter(formatter)
    # logger.addHandler(console)

    return logger
