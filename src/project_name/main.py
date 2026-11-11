# --------------------------- DOCUMENTATION ---------------------------#
"""
uv remove xyz
uv sync
"""

# --------------------------- IMPORTS ---------------------------#
from config import LOG_PATH
from loguru import logger

# --------------------------- CONSTANTS ---------------------------#

# --------------------------- FUNCTIONS ---------------------------#

# --------------------------- MAIN CODE ---------------------------#
if __name__ == "__main__":
    logger.remove()
    logger.add(LOG_PATH, rotation="1 MB", retention="7 days")
