"""
Thin Dropbox wrapper, extended from SharePoint-submittal-agent's version
with a file_exists check so re-runs don't re-upload attachments already
pulled down in a prior run.
"""
import dropbox
from dropbox.exceptions import ApiError

from config import Config


class DropboxClient:
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.dbx = dropbox.Dropbox(
            oauth2_refresh_token=cfg.DBX_REFRESH_TOKEN,
            app_key=cfg.DBX_APP_KEY,
            app_secret=cfg.DBX_APP_SECRET,
        )

    def download_bytes(self, path: str) -> bytes | None:
        try:
            _, resp = self.dbx.files_download(path)
            return resp.content
        except ApiError as e:
            if isinstance(e.error, dropbox.files.DownloadError) and e.error.is_path():
                return None
            raise

    def upload_bytes(self, path: str, data: bytes) -> None:
        self.dbx.files_upload(data, path, mode=dropbox.files.WriteMode("overwrite"))

    def file_exists(self, path: str) -> bool:
        try:
            self.dbx.files_get_metadata(path)
            return True
        except ApiError:
            return False

    def load_workbook_bytes(self) -> bytes | None:
        return self.download_bytes(self.cfg.DBX_LOG_PATH)

    def save_workbook_bytes(self, data: bytes) -> None:
        self.upload_bytes(self.cfg.DBX_LOG_PATH, data)
