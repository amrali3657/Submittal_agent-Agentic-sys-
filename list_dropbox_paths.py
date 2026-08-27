"""
list_dropbox_paths.py

One-time helper: lists folders/files under a given Dropbox path so you can
find the exact path string to use for DBX_LOG_PATH / DBX_RESPONSES_FOLDER,
instead of guessing at spacing or casing. Dropbox paths are case-insensitive
for lookups but the exact `path_display` casing is what shows up in the UI
and in shared links, so it's worth copying exactly.

Run:
    python list_dropbox_paths.py                # lists your Dropbox root
    python list_dropbox_paths.py "/EWD Contract 2"   # lists inside a folder
"""
import sys

import dropbox

from config import Config


def main():
    cfg = Config()
    dbx = dropbox.Dropbox(
        oauth2_refresh_token=cfg.DBX_REFRESH_TOKEN,
        app_key=cfg.DBX_APP_KEY,
        app_secret=cfg.DBX_APP_SECRET,
    )

    path = sys.argv[1] if len(sys.argv) > 1 else ""
    print(f"Listing: {path or '(root)'}\n")

    try:
        result = dbx.files_list_folder(path)
    except dropbox.exceptions.ApiError as e:
        print(f"Could not list '{path}': {e}")
        print(
            "\nIf this is a permission error, check the app's access type is "
            "'Full Dropbox' (not 'App folder') under dropbox.com/developers/apps."
        )
        return

    entries = result.entries
    while result.has_more:
        result = dbx.files_list_folder_continue(result.cursor)
        entries.extend(result.entries)

    if not entries:
        print("(empty)")
        return

    for e in sorted(entries, key=lambda x: x.name.lower()):
        kind = "FOLDER" if isinstance(e, dropbox.files.FolderMetadata) else "file"
        print(f"  [{kind:6}] {e.path_display}")

    print(
        "\nCopy the exact path_display shown above (including capitalization) "
        "into DBX_LOG_PATH / DBX_RESPONSES_FOLDER. Note: you don't need to "
        "pre-create the responses folder — files_upload creates any missing "
        "parent folders automatically on first upload."
    )


if __name__ == "__main__":
    main()
