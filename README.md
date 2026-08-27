# Submittal Notification Agent — orchestrated version

Combines the good parts of `SharePoint-submittal-agent` and
`Claude-Web-Agent` into one workflow, matched to the real submittal
process: a shop drawing gets submitted, Jacobs/the designer responds and
uploads their response to that revision's SharePoint item (submittals
often go through several revisions, e.g. `134-R0`, `134-R1`, `134-R3`),
and this agent's job is to notice that, save the response, update the log,
and tell the team — never to touch the portal itself.

```
[Chrome, running long-term, logged into the Jacobs SharePoint site]
                 |  (CDP, port 9222)
                 v
      sharepoint_tab_client.py
      1. finds the open tab, 2. EXPLICITLY navigates it to the list view,
      3. checks for a login redirect (fails loudly if session expired),
      4. reads list items + attachments via SharePoint REST, through that
      session's cookies. READ-ONLY — no write calls exist in this project.
                 |
                 v
      log_writer.py (unchanged from SharePoint-submittal-agent)
      -- diffs against the Dropbox "Shop Drawing Log", updates only the
      sync-owned columns, leaves manual columns alone
                 |
                 v
      attachments.py -- downloads the Jacobs/designer response file for
      each revision (filename matching "SRC"), uploads to
      /EWD Contract 2/Engineer Responses/<TUL#>/, skips duplicates.
      Revision is already encoded in TUL#, so R1/R2/R3 responses each
      land in their own subfolder — nothing gets overwritten.
                 |
                 v
      dropbox_client.py -- writes the updated workbook + response files
                 |
                 v
      notify.py -- emails the team: new items, status changes, a status
      breakdown, and links to newly saved response files
```

## 1. Navigation — explicit, not assumed

Earlier version of this just checked that *some* open tab's URL contained
a substring and called the REST API directly. That's not good enough to
hand off — if the tab happened to be sitting on a stale page, or the
session had quietly expired, it would look like "0 submittals" instead of
failing clearly. `sharepoint_tab_client.py` now follows a stated contract:

1. Find an open tab whose URL contains `SP_TAB_URL_SUBSTRING` (proves we
   have a live browser context to borrow cookies from).
2. **Explicitly navigate** that tab to
   `{SP_SITE_URL}/Lists/{SP_LIST_NAME}/AllItems.aspx` and wait for the
   page to settle.
3. Check the resulting URL for a Microsoft login redirect. If found, the
   session has expired — the agent **raises immediately** with a clear
   "log back into the tab" message instead of silently proceeding.
4. Only then does it call the REST endpoint the list view itself just
   loaded from (`_api/web/lists/GetByTitle(...)`), using that same
   session.

This is all in `SharePointTabClient._navigate_to_list()` — read it before
pointing this at the live portal, and adjust `AllItems.aspx` if your
Jacobs site uses a custom view URL instead of the default.

## 2. Read-only, by construction

There is no method anywhere in this project that creates, updates,
deletes, or comments on anything in SharePoint — only `GET` requests
through `_rest_get()` and `download_attachment()`. The agent's entire
write surface is: the Dropbox log file, the Dropbox response-file
subfolder, and the notification email. That's stated explicitly in
`sharepoint_tab_client.py`'s module docstring as a constraint for anyone
extending it later, not just as a design note.

## 3. Response files — the `SRC` pattern and revisions

`SP_RESPONSE_FILE_PATTERN` (default `SRC`) identifies which attachment on
a submittal item is the actual Jacobs/designer response, as opposed to the
originally submitted shop drawing or a transmittal cover sheet sitting on
the same item. Only files matching that pattern get downloaded.

Because the SharePoint list `Title` field encodes the revision (e.g.
`TUI Submittal 134.R3`), and that's exactly what's used as both the log's
match key and the Dropbox subfolder name, each revision's response lands
at its own path — `/EWD Contract 2/Engineer Responses/134.R3/...` — so a
new revision's response never overwrites an earlier one's, and the log's
revision history stays intact.

## Setup

### 1. Launch Chrome with remote debugging, log into SharePoint

```bash
# macOS
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
  --remote-debugging-port=9222 --user-data-dir="$HOME/chrome-agent-profile"

# Windows
chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\chrome-agent-profile"
```

Log into `jacobsengineering.sharepoint.com/sites/CP703215CH/EWDC2` in any
tab of this Chrome instance and leave it running — the agent attaches to
it and navigates it itself; it doesn't launch its own browser.

### 2. Confirm field names

```bash
pip install -r requirements.txt
playwright install chromium
python list_fields.py
```

Compare the printed internal names against what `sharepoint_tab_client.py`
assumes (`Title`, `Status`, `SUBNo`, `StartDate`, `Disposition`, `ItemNo`).

### 3. Dropbox app (same as SharePoint-submittal-agent)

Reuse the same `DBX_APP_KEY` / `DBX_APP_SECRET` / `DBX_REFRESH_TOKEN` if
you already set one up — **but check its access type first**:

1. **dropbox.com/developers/apps** → your app → confirm it's **Scoped
   access / Full Dropbox**, not **App folder**. App folder sandboxes the
   app to its own isolated `/Apps/<AppName>/` directory and can't reach
   your existing Shop Drawing Log wherever it actually lives. If the
   existing app was created as App folder, create a new one as Full
   Dropbox — access type can't be changed after creation.
2. **Permissions** tab → enable `files.content.read`, `files.content.write`,
   and `files.metadata.read` (the last one is needed for the duplicate
   check before uploading a response file).
3. Generate a refresh token the same way as before (see the original
   repo's README) and set `DBX_APP_KEY` / `DBX_APP_SECRET` /
   `DBX_REFRESH_TOKEN`.

### 4. Find the exact Dropbox paths

Don't guess at spacing/casing by eye — Dropbox paths need to match
exactly. Once the app credentials above are set:

```bash
python list_dropbox_paths.py                    # your Dropbox root
python list_dropbox_paths.py "/EWD Contract 2"   # inside a folder
```

This prints the exact `path_display` for everything at that level. Copy
the log file's path into `DBX_LOG_PATH`. For `DBX_RESPONSES_FOLDER`, pick
any path you want (existing or new) — you don't need to pre-create it,
`files_upload` creates any missing parent folders automatically on first
upload.

### 5. Environment variables

```
CDP_URL=http://localhost:9222
SP_TAB_URL_SUBSTRING=EWDC2
SP_SITE_URL=https://jacobsengineering.sharepoint.com/sites/CP703215CH/EWDC2
SP_LIST_NAME=Submittals
SP_RESPONSE_FILE_PATTERN=SRC

DBX_APP_KEY=...
DBX_APP_SECRET=...
DBX_REFRESH_TOKEN=...
DBX_LOG_PATH=/EWD Contract 2/Shop Drawing Log.xlsx
DBX_SHEET_NAME=Ongoing & Submitted
DBX_RESPONSES_FOLDER=/EWD Contract 2/Engineer Responses
DBX_SHARED_LINK_BASE=          # optional

GMAIL_ADDRESS=...
GMAIL_APP_PASSWORD=...
NOTIFY_TO=...
```

### 6. Dry run, then real run

```bash
DRY_RUN=1 python main.py   # prints the diff, downloads/writes/emails nothing
python main.py              # real run
```

## Known gaps to sanity-check before trusting this on the live log

- **`AllItems.aspx` URL assumption**: step 2 of the navigation contract
  assumes the default modern-list view URL pattern. If Jacobs' site uses
  a custom view, adjust `_navigate_to_list()` accordingly — `list_fields.py`
  won't catch this since it calls REST directly without navigating first.
- **Session lifetime**: if the SharePoint session expires between runs,
  the agent now fails clearly (rather than silently) — but it still needs
  a human to re-log-in the tab. Not self-healing like a service principal.
- **`SRC` pattern**: confirm this matches your actual response filenames
  before the first real run — `python list_fields.py` shows attachment
  data too if you want to check filenames on a live item first.
- **"Date responded by Jacobs" could now be automated**: `SharePoint-submittal-agent`'s
  README originally left this log column as manual-only, reasoning that
  "SharePoint has no actual responded date, only a Target/due date." A live
  item screenshot shows that's not quite right — there's a `Returned to
  Contractor Date` field on each item that looks like exactly this. Not
  wired in yet since changing what's synced vs. manual is a real behavior
  change worth confirming first — if you want it added to `SYNCED_FIELDS`,
  say the word and I'll map it in.
- **Whole-file overwrite**: each run re-uploads the entire log workbook.
  Not safe if someone's editing it in Dropbox at the exact moment a run
  fires.
- **Scheduling**: needs a live local Chrome instance, so this can't run on
  GitHub Actions like the Graph version — needs a machine (or persistent
  VM) where that Chrome + logged-in tab lives, via cron / Task Scheduler.
- **Generalizing to RFIs**: this same shape (navigate → read-only REST
  pull → log diff/update → response-file sync → status-breakdown email)
  should carry over to an RFI log with a different list name, column
  mapping, and Dropbox path once this is validated on submittals.
