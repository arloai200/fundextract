"""
cleanup.py — Immediately delete temporary PDF files after processing.
No financial data is retained on disk.
"""

import os
import logging

logger = logging.getLogger(__name__)


def delete_temp_file(path: str) -> None:
    """Silently delete a file; log a warning if deletion fails."""
    try:
        if path and os.path.exists(path):
            os.remove(path)
            logger.info("Deleted temp file: %s", path)
    except OSError as exc:
        logger.warning("Could not delete temp file %s: %s", path, exc)
