"""
One-time diagnostic: connects through the open tab and prints the first
list item's raw fields, so you can confirm internal field names (Status,
SUBNo, StartDate, etc.) match what sharepoint_tab_client.py assumes, before
trusting a real sync run. SharePoint's classic display names ("SUB No.",
"Spec Sec.") are never the actual field names the REST API returns.

Run:
    python list_fields.py
"""
import asyncio

from config import Config
from sharepoint_tab_client import SharePointTabClient, SharePointTabConfig


async def main():
    cfg = Config()
    sp_cfg = SharePointTabConfig(
        cdp_url=cfg.CDP_URL,
        site_url=cfg.SP_SITE_URL,
        list_title=cfg.SP_LIST_NAME,
        list_url_path=cfg.SP_LIST_URL_PATH,
        tab_url_substring=cfg.SP_TAB_URL_SUBSTRING,
        response_file_pattern=cfg.SP_RESPONSE_FILE_PATTERN,
    )
    async with SharePointTabClient(sp_cfg) as sp:
        data = await sp._rest_get(
            f"web/lists/GetByTitle('{cfg.SP_LIST_NAME}')/items?$top=1&$expand=AttachmentFiles"
        )
        items = data.get("value", [])
        if not items:
            print("List has no items to sample from.")
            return
        print(f"Fields on first item of '{cfg.SP_LIST_NAME}':\n")
        for key, val in items[0].items():
            print(f"  {key!r:30} -> {val!r}")
        print(
            "\nMatch these against what sharepoint_tab_client.get_submittals() reads "
            "(Title, Status, SUBNo, StartDate, Disposition, ItemNo) and adjust the "
            "field names in that file if any don't line up."
        )


if __name__ == "__main__":
    asyncio.run(main())
