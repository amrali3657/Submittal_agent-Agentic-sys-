"""
main.py — orchestrator

Sense -> Act -> Verify, applied to the whole workflow. Matches the real
process: a submittal/shop drawing is submitted, Jacobs/the designer
responds and uploads their response to that revision's list item, and this
agent's only job is to notice that, save it, and tell people:

  1. Attach to the already-open, already-logged-in SharePoint tab, then
     explicitly navigate it to the list view (see the navigation contract
     in sharepoint_tab_client.py) — fails loudly if the session's expired
     rather than silently reading nothing.
  2. Pull normalized submittal data via SharePoint REST, through that
     session (no DOM scraping, no Graph API, no Azure app registration).
     READ ONLY — nothing in this project writes back to SharePoint.
  3. Diff against the Shop Drawing Log workbook (a plain local file — see
     local_dropbox_client.py) and update only the sync-owned columns
     (log_writer.py, reused unchanged).
  4. Download each revision's Jacobs/designer response file (filename
     matching SP_RESPONSE_FILE_PATTERN, default "SRC") and save it into
     the matching local "Submittal 0XX - .../Revision 0X" folder, skipping
     ones already saved. Also a plain local file write — Dropbox's own
     desktop client handles syncing it, not this code.
  5. Email the team: new items, status changes, a current status
     breakdown, and paths to newly saved response files.

Every line printed is timestamped (see _log below) so a redirected log
file from an unattended nightly run is actually readable afterward, not
just a wall of text with no sense of when anything happened or how long
each step took.

Run:
    python main.py
Dry run (computes and prints the diff, writes nothing, sends no email,
downloads nothing from SharePoint):
    DRY_RUN=1 python main.py
"""
import asyncio
import sys
from datetime import datetime

from config import Config
from sharepoint_tab_client import SharePointTabClient, SharePointTabConfig
from local_dropbox_client import LocalDropboxClient
from log_writer import update_workbook
from attachments import sync_responses
from notify import notify_changes


def _log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}", flush=True)


async def async_main() -> None:
    _log("===== Run started =====")
    cfg = Config()

    sp_cfg = SharePointTabConfig(
        cdp_url=cfg.CDP_URL,
        site_url=cfg.SP_SITE_URL,
        list_title=cfg.SP_LIST_NAME,
        list_url_path=cfg.SP_LIST_URL_PATH,
        tab_url_substring=cfg.SP_TAB_URL_SUBSTRING,
        response_file_pattern=cfg.SP_RESPONSE_FILE_PATTERN,
    )

    dbx = LocalDropboxClient(cfg)

    _log(f"Attaching to open Chrome tab at {cfg.CDP_URL}...")
    async with SharePointTabClient(sp_cfg) as sp:
        _log(f"Navigated OK. Pulling submittals from '{cfg.SP_LIST_NAME}' (via open-tab session)...")
        submittals = await sp.get_submittals()
        _log(f"Pulled {len(submittals)} submittal(s).")

        _log(f"Reading current log: {cfg.DBX_LOG_PATH}")
        existing_bytes = dbx.load_workbook_bytes()

        _log("Computing changes...")
        new_bytes, added, changed = update_workbook(cfg, existing_bytes, submittals)

        _log("Checking for new Jacobs/designer response files...")
        saved = {} if cfg.DRY_RUN else await sync_responses(sp, dbx, submittals)
        if saved:
            total = sum(len(v) for v in saved.values())
            _log(f"Saved {total} new response file(s) across {len(saved)} submittal revision(s).")

    if not (added or changed or saved):
        _log("No changes detected. ===== Run finished =====")
        return

    _log(f"New submittals ({len(added)}): {added}")
    _log(f"Status changes ({len(changed)}): {changed}")

    if cfg.DRY_RUN:
        _log("DRY_RUN=1 set — not writing the log, not sending email, not downloading response files.")
        _log("===== Run finished (dry run) =====")
        return

    _log("Writing updated log...")
    dbx.save_workbook_bytes(new_bytes)

    _log("Notifying team...")
    notify_changes(cfg, added, changed, submittals, saved_responses=saved)
    _log("===== Run finished =====")


def main() -> None:
    try:
        asyncio.run(async_main())
    except Exception as e:
        _log(f"ERROR: {e}")
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
