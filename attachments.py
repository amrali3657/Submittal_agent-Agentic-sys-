"""
attachments.py

Downloads the Jacobs/designer response file (filename matching
SP_RESPONSE_FILE_PATTERN, e.g. "SRC") for each submittal revision from
SharePoint — read-only, through the open tab's session — and saves it into
the matching local "Submittal 0XX - .../Revision 0X" folder under
DBX_SUBMITTALS_ROOT, confirmed against the real folder structure:

    Technicore Dropbox / <Project> / 18_Submittals and Shop Drawings /
        Submittal 012 - Air Quality Monitoring Plan /
            Revision 00 /
            Revision 01 /

The submittal folder is matched by Item No. (e.g. "01" -> "Submittal 001 - ...")
since that's the only part of the folder name guaranteed to be stable — the
rest is a free-text description. The revision folder is matched (or created,
since this is our own storage, not the portal) from the revision number
parsed out of the SUB No. field (e.g. "01561-01-R0" -> revision 0 ->
"Revision 00").

Nothing here writes back to SharePoint — only reads, per sharepoint_tab_client.py.
"""
from __future__ import annotations

from local_dropbox_client import LocalDropboxClient, parse_revision_number
from sharepoint_tab_client import SharePointTabClient


async def sync_responses(
    sp: SharePointTabClient,
    dbx: LocalDropboxClient,
    submittals: list[dict],
) -> dict[str, list[str]]:
    """
    Returns {tul_submittal_no: [relative_paths_saved_this_run...]}.
    Files already present at the destination are skipped, so re-runs don't
    re-download or re-save unchanged files.
    """
    saved: dict[str, list[str]] = {}

    for s in submittals:
        response_files = sp.filter_response_attachments(s)
        if not response_files:
            continue

        item_no = s.get("item_no")
        revision = parse_revision_number(s.get("jacobs_submittal_no"))
        if not item_no or revision is None:
            print(
                f"WARNING: skipping response file(s) for '{s['tul_submittal_no']}' — "
                f"could not determine folder (item_no={item_no!r}, "
                f"jacobs_submittal_no={s.get('jacobs_submittal_no')!r})"
            )
            continue

        submittal_folder = dbx.find_submittal_folder(item_no)
        if submittal_folder is None:
            print(
                f"WARNING: no local folder found for Item No. {item_no} "
                f"('{s['tul_submittal_no']}') under DBX_SUBMITTALS_ROOT — skipped."
            )
            continue

        revision_folder = dbx.find_or_create_revision_folder(submittal_folder, revision)

        tul_no = s["tul_submittal_no"]
        for att in response_files:
            dest_path = revision_folder / att["file_name"]
            if dbx.file_exists(dest_path):
                continue
            content = await sp.download_attachment(att["server_relative_url"])
            dbx.write_bytes(dest_path, content)
            saved.setdefault(tul_no, []).append(dbx.relative_path(dest_path))

    return saved
