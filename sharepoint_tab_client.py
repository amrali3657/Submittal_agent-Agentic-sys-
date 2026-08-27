"""
sharepoint_tab_client.py

Reads SharePoint submittal data and Jacobs/engineer response files through
an *already open, already logged-in* browser tab, by attaching Playwright
to a running Chrome instance over the Chrome DevTools Protocol (CDP) —
rather than launching a fresh browser and re-authenticating, and rather
than requiring an Azure AD app registration (which needs Jacobs' tenant
admin to approve, since it's their tenant, not yours).

============================== READ-ONLY, BY DESIGN ==============================
This client issues GET requests only. There is no method anywhere in this
file (or anywhere else in this project) that creates, updates, deletes, or
comments on a SharePoint list item or attachment. The workflow this
supports is: read the submittal log + download the response file Jacobs
already uploaded -> write the results to Dropbox -> email the team. Nothing
in this pipeline ever touches the Jacobs portal itself. If someone extends
this file, that guarantee needs to stay true — do not add POST/PATCH/DELETE
calls against SharePoint here.
====================================================================================

NAVIGATION CONTRACT (explicit, so this isn't "hope the open tab happens to
be on the right page"):
  1. Find an open tab whose URL contains SP_TAB_URL_SUBSTRING (proves we
     have a live, authenticated browser context to borrow cookies from —
     it does NOT need to already be on the exact list view).
  2. Explicitly navigate that tab to the list's "AllItems.aspx" view at
     `{site_url}/Lists/{list_title}/AllItems.aspx` and wait for the
     navigation to settle.
  3. Check the resulting URL for a Microsoft login redirect
     (`login.microsoftonline.com` / `/_forms/default.aspx`). If found, the
     session has expired — raise immediately with an actionable message
     instead of silently trying to read a login page as if it were data.
  4. Only once navigation is confirmed do we call the REST endpoint
     (`_api/web/lists/GetByTitle(...)`) that the list view itself just
     loaded from, using the same session cookies.
This module deliberately does NOT scrape the rendered DOM for data (step 4
uses REST, not HTML parsing) — step 2's navigation exists purely to prove
and refresh a valid session before trusting the REST call, and to fail
loudly and specifically if that session is stale, rather than the agent
silently pulling nothing or pulling stale cached data.

Trade-off vs. the Graph API app-only approach: this only works while the
tab stays open and your own session is valid, and it's scoped to whatever
permissions your own account has on the site (not a service principal).
That's an acceptable trade for not needing tenant-admin consent.

Explicitly NOT used here: cdp_input_driver.py's humanized mouse/keyboard
CDP dispatch from Claude-Web-Agent. There's no UI interaction to humanize —
this only issues authenticated GET requests through the existing session —
and that module's actual purpose (defeating isTrusted / bot-detection
checks) has no legitimate role in reading your own already-authenticated
SharePoint session.

Prerequisite: Chrome started with remote debugging enabled, e.g.:
    chrome --remote-debugging-port=9222 --user-data-dir="/path/to/agent-profile"
with the target SharePoint site open and logged in, in any tab.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from playwright.async_api import async_playwright, BrowserContext, Page

# Matches "TUI Submittal 002.R2 ..." embedded in the Title field — same
# convention as SharePoint-submittal-agent/src/sharepoint_client.py.
TUI_NO_PATTERN = re.compile(r"TUI\s+Submittal\s+([A-Za-z0-9.\-]+)", re.IGNORECASE)


@dataclass
class SharePointTabConfig:
    cdp_url: str
    site_url: str
    list_title: str          # display name, used for the REST GetByTitle() call
    list_url_path: str       # actual URL slug, used for navigation — NOT always
                              # the same as the display name (confirmed on this
                              # site: display name "Submittals", URL slug
                              # "EWD C2 Submittals")
    tab_url_substring: str
    # Files matching this substring (case-insensitive) in their filename are
    # treated as the Jacobs/designer response document for that revision —
    # confirmed from a live item: attachments included both the originally
    # submitted "...Air Quality Monitoring Plan.pdf" and the response files
    # "...Air Quality Monitoring Plan-SRC.docx" / "...-SRC.pdf". Matching on
    # plain "SRC" catches both response files regardless of the separator
    # character used in a given filename.
    response_file_pattern: str = "SRC"


class SharePointTabClient:
    def __init__(self, cfg: SharePointTabConfig):
        self.cfg = cfg
        self._playwright = None
        self._browser = None
        self.context: BrowserContext | None = None
        self.page: Page | None = None

    async def __aenter__(self) -> "SharePointTabClient":
        self._playwright = await async_playwright().start()
        self._browser = await self._playwright.chromium.connect_over_cdp(self.cfg.cdp_url)

        target_page = None
        for ctx in self._browser.contexts:
            for pg in ctx.pages:
                if self.cfg.tab_url_substring in pg.url:
                    target_page, self.context = pg, ctx
                    break
            if target_page:
                break

        if target_page is None:
            raise RuntimeError(
                f"No open tab found containing '{self.cfg.tab_url_substring}' at "
                f"{self.cfg.cdp_url}. Open the SharePoint site in a tab of that "
                f"Chrome instance (logged in) and re-run."
            )
        self.page = target_page

        await self._navigate_to_list()
        return self

    async def __aexit__(self, exc_type, exc, tb) -> None:
        # Deliberately do NOT close the browser/context/page — it's the
        # user's own long-running tab, not this agent's to tear down.
        if self._playwright:
            await self._playwright.stop()

    async def _navigate_to_list(self) -> None:
        """
        Step 2-3 of the navigation contract (see module docstring): drive
        the open tab to the exact list view, and fail loudly if that
        navigation lands on a Microsoft login page instead — which means
        the session has expired and needs a human to re-log-in the tab.
        This is a deliberate, visible step rather than an assumption that
        whatever page happened to be open already has current data.
        """
        assert self.page is not None
        # Use the URL slug, not the display name — SharePoint list URLs
        # don't always match what's shown in the UI (confirmed on this
        # site: displayed as "Submittals", lives at ".../Lists/EWD C2
        # Submittals/AllItems.aspx").
        list_path_url = self.cfg.list_url_path.replace(" ", "%20")
        list_view_url = f"{self.cfg.site_url}/Lists/{list_path_url}/AllItems.aspx"

        await self.page.goto(list_view_url, wait_until="networkidle", timeout=30_000)

        current_url = self.page.url.lower()
        if "login.microsoftonline.com" in current_url or "/_forms/default.aspx" in current_url:
            raise RuntimeError(
                f"Navigating to {list_view_url} redirected to a Microsoft login page "
                f"({self.page.url}) — the session in this tab has expired. Log back "
                f"into SharePoint in that tab manually, then re-run."
            )

    async def _rest_get(self, path: str):
        """The only network method in this class that talks to SharePoint.
        GET only — see the read-only banner in the module docstring."""
        assert self.context is not None
        url = f"{self.cfg.site_url}/_api/{path}"
        resp = await self.context.request.get(
            url, headers={"Accept": "application/json;odata=nometadata"}
        )
        if not resp.ok:
            body = await resp.text()
            raise RuntimeError(f"SharePoint REST call failed [{resp.status}]: {url}\n{body[:500]}")
        return await resp.json()

    async def get_submittals(self) -> list[dict]:
        """
        Pulls all items from the list via SharePoint REST, through the open
        tab's session. Returns the same normalized shape the Graph-based
        client in SharePoint-submittal-agent produces, so log_writer.py
        works unmodified against either source.
        """
        list_title = self.cfg.list_title.replace("'", "''")
        items: list[dict] = []
        path = (
            f"web/lists/GetByTitle('{list_title}')/items"
            f"?$top=500&$expand=AttachmentFiles"
        )
        # Follow pagination explicitly rather than assuming one page covers
        # everything — this list currently has 472 items (fits in one page
        # at $top=500), but this shouldn't silently start dropping rows if
        # it grows past that.
        while path:
            data = await self._rest_get(path)
            items.extend(data.get("value", []))
            next_link = data.get("odata.nextLink") or data.get("__next")
            if not next_link:
                break
            # next_link is a full URL; _rest_get expects a path appended to
            # {site_url}/_api/, so strip that prefix back off.
            api_prefix = f"{self.cfg.site_url}/_api/"
            path = next_link[len(api_prefix):] if next_link.startswith(api_prefix) else None
            if path is None:
                break

        normalized, skipped = [], []
        for item in items:
            title = item.get("Title", "") or ""
            m = TUI_NO_PATTERN.search(title)
            tul_no = m.group(1) if m else str(item.get("ItemNo") or item.get("Id") or "").strip()
            if not tul_no:
                skipped.append(title or item.get("Id"))
                continue

            status = str(item.get("Status", "") or "").strip()
            disposition = str(item.get("Disposition", "") or "").strip()
            # Disposition is the more accurate/current field per site review —
            # lead with it, keep Status as secondary context rather than the
            # other way around.
            if disposition:
                combined_status = f"{disposition} ({status})" if status else disposition
            else:
                combined_status = status

            raw_attachments = item.get("AttachmentFiles") or []
            if isinstance(raw_attachments, dict):  # odata=verbose shape fallback
                raw_attachments = raw_attachments.get("results", [])

            attachments = [
                {"file_name": a["FileName"], "server_relative_url": a["ServerRelativeUrl"]}
                for a in raw_attachments
            ]

            normalized.append({
                "tul_submittal_no": tul_no,
                "description": title,
                "jacobs_submittal_no": item.get("SUBNo", ""),
                "status": combined_status,
                "start_date": item.get("StartDate", ""),
                "item_no": item.get("ItemNo", ""),
                "disposition": disposition,
                "modified": item.get("Modified", ""),
                "attachments": attachments,
                "sp_item_id": item.get("Id"),
            })

        if skipped:
            print(
                f"WARNING: {len(skipped)} item(s) had no parseable submittal "
                f"number — skipped: {skipped[:5]}{'...' if len(skipped) > 5 else ''}"
            )
        return normalized

    async def download_attachment(self, server_relative_url: str) -> bytes:
        """Downloads one attachment through the same authenticated session."""
        assert self.context is not None
        tenant_root = "/".join(self.cfg.site_url.split("/")[:3])  # scheme://host
        full_url = tenant_root + server_relative_url
        resp = await self.context.request.get(full_url)
        if not resp.ok:
            raise RuntimeError(f"Attachment download failed [{resp.status}]: {full_url}")
        return await resp.body()

    def filter_response_attachments(self, submittal: dict) -> list[dict]:
        """
        Picks out the Jacobs/designer response file from everything attached
        to this submittal revision item — not the shop drawing that was
        originally submitted, not transmittal cover sheets, just the file
        matching response_file_pattern (default "SRC").
        """
        kw = self.cfg.response_file_pattern.lower()
        return [a for a in submittal.get("attachments", []) if kw in a["file_name"].lower()]
