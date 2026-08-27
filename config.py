"""
Central configuration. Same env-var pattern as SharePoint-submittal-agent
so this can run locally (.env) or as a scheduled task unchanged.
"""
import os


def _require(name: str) -> str:
    val = os.environ.get(name)
    if not val:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return val


class Config:
    # --- Browser / open-tab connection ---
    # Chrome must be launched with remote debugging enabled, e.g.:
    #   chrome --remote-debugging-port=9222 --user-data-dir="/path/to/agent-profile"
    # with the SharePoint list view open and logged in, in any tab.
    CDP_URL = os.environ.get("CDP_URL", "http://localhost:9222")
    SP_TAB_URL_SUBSTRING = os.environ.get("SP_TAB_URL_SUBSTRING", "EWDC2")

    # --- SharePoint (REST via the open tab's session — no Azure app registration) ---
    # Confirmed from live screenshots (2026-08-27):
    SP_SITE_URL = os.environ.get(
        "SP_SITE_URL",
        "https://jacobsengineering.sharepoint.com/sites/CP703215CH/EWDC2",
    )
    SP_LIST_NAME = os.environ.get("SP_LIST_NAME", "Submittals")  # display name — used in the REST GetByTitle() call
    SP_LIST_URL_PATH = os.environ.get("SP_LIST_URL_PATH", "EWD C2 Submittals")  # URL slug — used only for navigation
    # Files matching this substring (case-insensitive) in their filename are
    # treated as the Jacobs/designer response document for that submittal
    # revision — confirmed from a live item: e.g.
    # "01561-01-R0-Air Quality Monitoring Plan-SRC.pdf" matches "SRC", while
    # the originally submitted "01561-01-R0 - Air Quality Monitoring Plan.pdf"
    # on the same item does not.
    SP_RESPONSE_FILE_PATTERN = os.environ.get("SP_RESPONSE_FILE_PATTERN", "SRC")

    # --- Dropbox ---
    DBX_APP_KEY = _require("DBX_APP_KEY")
    DBX_APP_SECRET = _require("DBX_APP_SECRET")
    DBX_REFRESH_TOKEN = _require("DBX_REFRESH_TOKEN")

    DBX_LOG_PATH = _require("DBX_LOG_PATH")  # e.g. "/EWD Contract 2/Shop Drawing Log.xlsx"
    DBX_SHEET_NAME = os.environ.get("DBX_SHEET_NAME", "Ongoing & Submitted")
    # Response files are saved under <this>/<TUL Submittal #>/<filename>.
    # The TUL Submittal # already encodes the revision (e.g. "134.R3"), so
    # each revision's response lands in its own subfolder automatically —
    # earlier revisions' responses are never overwritten by later ones.
    DBX_RESPONSES_FOLDER = os.environ.get(
        "DBX_RESPONSES_FOLDER", "/EWD Contract 2/Engineer Responses"
    )
    # Optional: if you generate/know a Dropbox shared-link base for this
    # folder, notification emails will link straight to files instead of
    # showing bare Dropbox paths.
    DBX_SHARED_LINK_BASE = os.environ.get("DBX_SHARED_LINK_BASE", "")

    # --- Log layout: unchanged from SharePoint-submittal-agent so
    # log_writer.py can be reused verbatim. ---
    LOG_HEADER_ROW = int(os.environ.get("LOG_HEADER_ROW", "3"))
    LOG_FIRST_DATA_ROW = int(os.environ.get("LOG_FIRST_DATA_ROW", "4"))
    LOG_COLUMNS = {
        "location": "B",
        "description": "C",
        "date_received_from_subs": "D",
        "tul_submittal_no": "E",
        "jacobs_submittal_no": "F",
        "date_submitted_to_jacobs": "G",
        "date_responded_by_jacobs": "H",
        "status": "I",
        "date_sent_to_subcontractor": "J",
    }
    SYNCED_FIELDS = {
        "description", "tul_submittal_no", "jacobs_submittal_no",
        "date_submitted_to_jacobs", "status",
    }

    # --- Notification ---
    GMAIL_ADDRESS = os.environ.get("GMAIL_ADDRESS", "")
    GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
    NOTIFY_TO = os.environ.get("NOTIFY_TO", "")

    DRY_RUN = os.environ.get("DRY_RUN", "0") == "1"
