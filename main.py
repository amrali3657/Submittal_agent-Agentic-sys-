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
     READ ONLY — this step and step 4 never write anything back to
     SharePoint; nothing in this project does.
  3. Diff against the current Dropbox "Shop Drawing Log" and update only
     the sync-owned columns (log_writer.py, reused unchanged).
  4. Download each revision's Jacobs/designer response file (filename
     matching SP_RESPONSE_FILE_PATTERN, default "SRC") and upload it to a
     per-revision Dropbox subfolder, skipping ones already saved.
  5. Email the team: new items, status changes, a current status
     breakdown, and links to newly saved response files.

Run:
    python main.py
Dry run (computes and prints the diff, writes nothing, sends no email,
downloads nothing from SharePoint):
    DRY_RUN=1 python main.py
"""
import asyncio
import sys

from config import Config
from sharepoint_tab_client import SharePointTabClient, SharePointTabConfig
from dropbox_client import DropboxClient
from log_writer import update_workbook
from attachments import sync_responses
from notify import notify_changes


async def async_main() -> None:
    cfg = Config()

    sp_cfg = SharePointTabConfig(
        cdp_url=cfg.CDP_URL,
        site_url=cfg.SP_SITE_URL,
        list_title=cfg.SP_LIST_NAME,
        list_url_path=cfg.SP_LIST_URL_PATH,
        tab_url_substring=cfg.SP_TAB_URL_SUBSTRING,
        response_file_pattern=cfg.SP_RESPONSE_FILE_PATTERN,
    )

    print(f"Attaching to open Chrome tab at {cfg.CDP_URL}...")
    async with SharePointTabClient(sp_cfg) as sp:
        print("Pulling submittals from SharePoint (via open-tab session)...")
        submittals = await sp.get_submittals()
        print(f"Pulled {len(submittals)} submittal(s) from '{cfg.SP_LIST_NAME}'.")

        dbx = DropboxClient(cfg)

        print(f"Downloading current log: {cfg.DBX_LOG_PATH}")
        existing_bytes = dbx.load_workbook_bytes()

        print("Computing changes...")
        new_bytes, added, changed = update_workbook(cfg, existing_bytes, submittals)

        print("Checking for new Jacobs/designer response files...")
        uploaded = {} if cfg.DRY_RUN else await sync_responses(
            sp, dbx, submittals, cfg.DBX_RESPONSES_FOLDER
        )
        if uploaded:
            total = sum(len(v) for v in uploaded.values())
            print(f"Uploaded {total} new response file(s) across {len(uploaded)} submittal revision(s).")

    if not (added or changed or uploaded):
        print("No changes detected.")
        return

    print(f"New submittals ({len(added)}): {added}")
    print(f"Status changes ({len(changed)}): {changed}")

    if cfg.DRY_RUN:
        print("\nDRY_RUN=1 set — not writing to Dropbox, not sending email, not downloading response files.")
        return

    print("Uploading updated log to Dropbox...")
    dbx.save_workbook_bytes(new_bytes)

    print("Notifying team...")
    notify_changes(
        cfg, added, changed, submittals,
        uploaded_responses=uploaded,
        dropbox_shared_link_base=cfg.DBX_SHARED_LINK_BASE,
    )
    print("Done.")


def main() -> None:
    try:
        asyncio.run(async_main())
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
