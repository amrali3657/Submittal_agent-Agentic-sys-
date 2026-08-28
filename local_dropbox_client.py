"""
local_dropbox_client.py

Reads and writes files through the local filesystem path where the Dropbox
desktop client has already mounted "Technicore Dropbox" — NOT the Dropbox
API. There's no app to register, no OAuth, no admin approval needed: this
is exactly as much access as any other program running on this machine
already has to that folder, because the Dropbox client is the thing
actually doing the syncing in the background. Writing a file here and
having it show up in teammates' synced copies is Dropbox doing its
ordinary job, not something this code manages.

Trade-off vs. the API approach: this only works on a machine that has the
Dropbox desktop client installed, signed into the Technicore Dropbox team
space, and running. That's an easy bar to clear on the same machine that's
already running Chrome for the SharePoint side.

Reading a file that shows the cloud-only icon (not yet downloaded locally)
triggers Dropbox to hydrate it transparently on first access — a normal
file read may just take a moment longer the first time.
"""
from __future__ import annotations

import re
from pathlib import Path

REVISION_SUFFIX_PATTERN = re.compile(r"-R(\d+)\s*$", re.IGNORECASE)


def parse_revision_number(jacobs_submittal_no: str) -> int | None:
    """Extracts the revision number from a SUB No. like '01561-01-R0' -> 0."""
    m = REVISION_SUFFIX_PATTERN.search(str(jacobs_submittal_no or ""))
    return int(m.group(1)) if m else None


class LocalDropboxClient:
    def __init__(self, cfg):
        self.cfg = cfg
        self.root = Path(cfg.DROPBOX_ROOT_PATH)
        if not self.root.exists():
            raise RuntimeError(
                f"Dropbox root not found at {self.root} — confirm the Dropbox "
                f"desktop client is installed, signed in, and has synced the "
                f"Technicore Dropbox team folder on this machine."
            )

    # ---- log workbook ----
    def load_workbook_bytes(self) -> bytes | None:
        path = Path(self.cfg.DBX_LOG_PATH)
        if not path.exists():
            return None
        return path.read_bytes()

    def save_workbook_bytes(self, data: bytes) -> None:
        Path(self.cfg.DBX_LOG_PATH).write_bytes(data)

    # ---- response files ----
    def find_submittal_folder(self, item_no: str) -> Path | None:
        """
        Finds the folder for this submittal by Item No. (e.g. "01" -> a
        folder starting with "Submittal 001") under DBX_SUBMITTALS_ROOT.
        Matches by prefix rather than reconstructing the full folder name
        (which includes a free-text description) since the prefix is the
        only part guaranteed to be stable.
        """
        try:
            n = int(str(item_no).strip())
        except ValueError:
            return None
        prefix = f"Submittal {n:03d}"
        base = Path(self.cfg.DBX_SUBMITTALS_ROOT)
        matches = [p for p in base.iterdir() if p.is_dir() and p.name.startswith(prefix)]
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            print(f"WARNING: multiple folders match '{prefix}' under {base}: {matches}")
        return None

    def find_or_create_revision_folder(self, submittal_folder: Path, revision: int) -> Path:
        rev_folder = submittal_folder / f"Revision {revision:02d}"
        rev_folder.mkdir(exist_ok=True)
        return rev_folder

    def file_exists(self, path: Path) -> bool:
        return path.exists()

    def write_bytes(self, path: Path, data: bytes) -> None:
        path.write_bytes(data)

    def relative_path(self, path: Path) -> str:
        """Path relative to the Dropbox root, for use in notification emails —
        meaningful to any teammate's synced copy, unlike an absolute path
        that includes this machine's username."""
        try:
            return str(path.relative_to(self.root))
        except ValueError:
            return str(path)
