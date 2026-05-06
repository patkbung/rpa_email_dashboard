"""
config.py — Central configuration for the RPA Email Attachment Organizer.
"""

import os

# ── Base paths ────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DOWNLOADS_DIR   = os.path.join(BASE_DIR, "downloads")
REPORTS_DIR     = os.path.join(BASE_DIR, "reports")
TEST_FILES_DIR  = os.path.join(BASE_DIR, "test_files")

REPORT_FILE          = os.path.join(REPORTS_DIR, "daily_report.xlsx")
TEMP_DOWNLOADS_DIR   = os.path.join(BASE_DIR, "temp_downloads")

# ── Category sub-folders ──────────────────────────────────────────────────────
CATEGORY_DIRS = {
    "pdf":    os.path.join(DOWNLOADS_DIR, "pdf"),
    "word":   os.path.join(DOWNLOADS_DIR, "word"),
    "excel":  os.path.join(DOWNLOADS_DIR, "excel"),
    "image":  os.path.join(DOWNLOADS_DIR, "image"),
    "others": os.path.join(DOWNLOADS_DIR, "others"),
}

# ── File-type → category mapping ──────────────────────────────────────────────
EXTENSION_MAP = {
    ".pdf":  "pdf",
    ".doc":  "word",
    ".docx": "word",
    ".xls":  "excel",
    ".xlsx": "excel",
    ".csv":  "excel",
    ".png":  "image",
    ".jpg":  "image",
    ".jpeg": "image",
    ".gif":  "image",
    ".bmp":  "image",
    ".tiff": "image",
    ".webp": "image",
}

# ── Excel report column headers ───────────────────────────────────────────────
REPORT_COLUMNS = [
    "date",
    "time",
    "email_subject",
    "sender",
    "file_name",
    "file_type",
    "save_path",
    "status",
    "error_message",
]
