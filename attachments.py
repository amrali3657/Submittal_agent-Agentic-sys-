"""
attachments.py

Downloads the Jacobs/designer response file (see SP_RESPONSE_FILE_PATTERN)
for each submittal revision from SharePoint — read-only, through the open
tab's session — and uploads any not already present to a Dropbox subfolder
next to the Shop Drawing Log, so the notification email can point straight
at them. Nothing here writes back to SharePoint: it only ever reads the
list item and its attachments.
"""
from __future__ import annotations

from sharepoint_tab_client import SharePointTabClient


async def sync_responses(
    sp: SharePointTabClient,
    dbx,
    submittals: list[dict],
    dropbox_folder: str,
) -> dict[str, list[str]]:
    """
    Returns {tul_submittal_no: [dropbox_paths_uploaded_this_run...]}.

    tul_submittal_no already encodes the revision (e.g. "134.R3"), so each
    revision's response file lands in its own subfolder — a new revision
    never overwrites a prior one's saved response.

    Files already present at the destination path (matched by filename) are
    skipped, so re-runs don't re-download or re-upload unchanged files.
    """
    uploaded: dict[str, list[str]] = {}

    for s in submittals:
        response_files = sp.filter_response_attachments(s)
        if not response_files:
            continue

        tul_no = s["tul_submittal_no"]
        for att in response_files:
            dest_path = f"{dropbox_folder}/{tul_no}/{att['file_name']}"
            if dbx.file_exists(dest_path):
                continue
            content = await sp.download_attachment(att["server_relative_url"])
            dbx.upload_bytes(dest_path, content)
            uploaded.setdefault(tul_no, []).append(dest_path)

    return uploaded
