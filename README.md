# Submittal Notification Agent — orchestrated version

Combines the good parts of `SharePoint-submittal-agent` and
`Claude-Web-Agent` into one workflow, matched to the real submittal
process: a shop drawing gets submitted, Jacobs/the designer responds and
uploads their response to that revision's SharePoint item (submittals go
through several revisions), and this agent's job is to notice that, save
the response, update the log, and tell the team — never to touch the
portal itself.

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
      -- diffs against the Shop Drawing Log workbook, updates only the
      sync-owned columns, leaves manual columns alone
                 |
                 v
      attachments.py -- downloads the Jacobs/designer response file for
      each revision (filename matching "SRC"), saves into the matching
      local "Submittal 0XX - .../Revision 0X" folder, skips duplicates
                 |
                 v
      local_dropbox_client.py -- plain filesystem reads/writes through
      wherever the Dropbox desktop client has mounted "Technicore
      Dropbox" on this machine. NOT the Dropbox API.
                 |
                 v
      notify.py -- emails the team: new items, status changes, a status
      breakdown, and paths to newly saved response files
```

## Why local filesystem instead of the Dropbox API

Your screenshots showed the actual setup: "Technicore Dropbox" is mounted
as an ordinary folder via the Dropbox desktop client (visible in File
Explorer's sidebar, with cloud-sync status icons). That means there's
nothing to build against an API for — the agent just reads and writes
files at that path like any other program on the machine, and the
already-running Dropbox client handles syncing them in the background.

This sidesteps the whole earlier Dropbox-app-permission question: there's
no app to register, so there's nothing for a Business/Team admin to
approve or block. The only requirement is that the machine running this
has the Dropbox desktop client installed, signed into the Technicore
Dropbox team space, and running — true by default on your machine already,
and easy to ensure on whatever machine ends up running this on a schedule.

## The real folder structure (confirmed from live screenshots)

```
Technicore Dropbox\
  12106 East to West DSTT\                          <- project folder
    18_Submittals and Shop Drawings\
      Submittal 012 - Air Quality Monitoring Plan\   <- matched by Item No.
        Revision 00\
          01561-01-R0 - Air Quality Monitoring Plan.pdf       (originally submitted)
          01561-01-R0-Air Quality Monitoring Plan-SRC.docx    (Jacobs response)
          01561-01-R0-Air Quality Monitoring Plan-SRC.pdf     (Jacobs response)
        Revision 01\
          ...
```

`local_dropbox_client.find_submittal_folder()` matches on Item No. (e.g.
`01` → a folder starting with `Submittal 001`) rather than trying to
reconstruct the full folder name, since the description text after the
number is free-form and not worth guessing exactly. The revision folder
is found (or created, since this is our own storage — not the portal) from
the revision number parsed out of the `SUB No.` field, e.g. `01561-01-R0`
→ revision `0` → `Revision 00`.

## 1. Navigation — explicit, not assumed

`sharepoint_tab_client.py` follows a stated contract rather than assuming
whatever tab happens to be open has current data:

1. Find an open tab whose URL contains `SP_TAB_URL_SUBSTRING` (proves we
   have a live browser context to borrow cookies from).
2. **Explicitly navigate** that tab to
   `{SP_SITE_URL}/Lists/{SP_LIST_URL_PATH}/AllItems.aspx` and wait for the
   page to settle. Note this uses the URL slug (`EWD C2 Submittals`), not
   the list's display name (`Submittals`) — confirmed these differ on this
   site from a live screenshot.
3. Check the resulting URL for a Microsoft login redirect. If found, the
   session has expired — the agent **raises immediately** with a clear
   "log back into the tab" message instead of silently proceeding.
4. Only then does it call the REST endpoint the list view itself just
   loaded from (`_api/web/lists/GetByTitle(...)`), using that same
   session, with pagination.

## 2. Read-only on SharePoint, by construction

There is no method anywhere in this project that creates, updates,
deletes, or comments on anything in SharePoint — only `GET` requests
through `_rest_get()` and `download_attachment()`. The agent's entire
write surface is: the local Shop Drawing Log file, the local response-file
folders, and the notification email.

## 3. Response files — the `SRC` pattern

`SP_RESPONSE_FILE_PATTERN` (default `SRC`) identifies which attachment on
a submittal item is the Jacobs/designer response, as opposed to the
originally submitted shop drawing sitting on the same item — confirmed
from a live item's attachments (see folder structure above).

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

### 3. Find your local Dropbox paths

No app or credentials needed — just real filesystem paths. In File
Explorer, navigate to the relevant folders, then Shift+right-click →
**Copy as path** (or right-click → Properties → Location, on some
Windows versions) to get the exact path string:

- `DROPBOX_ROOT_PATH` — the "Technicore Dropbox" folder itself, e.g.
  `C:\Users\Amr\Technicore Dropbox`
- `DBX_LOG_PATH` — full path to the Shop Drawing Log workbook
- `DBX_SUBMITTALS_ROOT` — the `18_Submittals and Shop Drawings` folder for
  this project, e.g.
  `C:\Users\Amr\Technicore Dropbox\12106 East to West DSTT\18_Submittals and Shop Drawings`

### 4. Environment variables

```
CDP_URL=http://localhost:9222
SP_TAB_URL_SUBSTRING=EWDC2
SP_SITE_URL=https://jacobsengineering.sharepoint.com/sites/CP703215CH/EWDC2
SP_LIST_NAME=Submittals
SP_LIST_URL_PATH=EWD C2 Submittals
SP_RESPONSE_FILE_PATTERN=SRC

DROPBOX_ROOT_PATH=C:\Users\Amr\Technicore Dropbox
DBX_LOG_PATH=C:\Users\Amr\Technicore Dropbox\12106 East to West DSTT\...\Shop Drawing Log.xlsx
DBX_SHEET_NAME=Ongoing & Submitted
DBX_SUBMITTALS_ROOT=C:\Users\Amr\Technicore Dropbox\12106 East to West DSTT\18_Submittals and Shop Drawings

GMAIL_ADDRESS=...
GMAIL_APP_PASSWORD=...
NOTIFY_TO=...
```

### 5. Dry run, then real run

```bash
DRY_RUN=1 python main.py   # prints the diff, downloads/writes/emails nothing
python main.py              # real run
```

## Running it: two ways

### Push to GitHub

Already done on your end. Worth noting: **GitHub is for storing the code,
not running it.** GitHub Actions runners are ephemeral and headless — they
can't attach to your already-logged-in Chrome tab or see your Dropbox
folder. This has to actually execute on a machine that has both of those
live, which for now means your own machine (or a VM you set up the same
way) rather than a CI runner.

### Scheduled nightly run (Windows Task Scheduler)

1. Make sure Chrome (with `--remote-debugging-port=9222`, logged into
   SharePoint) and the Dropbox desktop client are both left running on
   whatever machine will execute this.
2. Write a small wrapper batch file, e.g. `run_agent.bat`:
   ```bat
   cd /d C:\path\to\submittal-notification-agent
   call venv\Scripts\activate.bat
   python main.py >> logs\run.log 2>&1
   ```
3. Open **Task Scheduler → Create Task**:
   - **General**: run whether user is logged on or not (if the machine
     stays logged in, "Run only when user is logged on" is simpler and
     avoids Chrome-session complications).
   - **Triggers**: New → Daily → your chosen time (e.g. 6:00 AM, before
     the team's day starts).
   - **Actions**: New → Start a program → point at `run_agent.bat`.
   - **Conditions**: uncheck "Start the task only if the computer is on
     AC power" if this is a laptop; check "Wake the computer to run this
     task" if it might be asleep.
4. Test it once with **Run** in Task Scheduler before trusting the nightly
   trigger, and check `logs\run.log` afterward.

## How to watch it run / confirm it worked

**Running it yourself, live:** just run `python main.py` in a normal
terminal window with Chrome visible on screen. You'll see two things
happening in parallel: timestamped status lines printing in the terminal
(`Attaching to open Chrome tab...`, `Navigated OK...`, `Pulled N
submittal(s)...`, etc.), and the actual Chrome tab visibly jumping to the
SharePoint list view when `_navigate_to_list()` runs — it's not headless,
so you can watch it navigate in real time. This is the easiest way to
sanity-check a first run or a config change before trusting it unattended.

**Checking on a scheduled overnight run,** in order of how much detail
each gives you:

1. **Did the notification email arrive?** The fastest signal — if the
   team got the summary email, the whole pipeline completed successfully.
   No email doesn't necessarily mean failure though (it's also silent
   when there are no changes to report — see `main.py`'s early return).
2. **Task Scheduler → your task → History tab** (or the **Last Run
   Result** column on the main Task Scheduler list). `0x0` means it exited
   cleanly; anything else means `main.py` raised and you should check the
   log.
3. **`logs\run.log`** — every run appends a `===== Run started =====` /
   `===== Run finished =====` block with a timestamp on every line, so
   you can see exactly what happened and how long each step took, even
   days later. This is where you'll see the specific error if the
   SharePoint session had expired overnight, or a warning that a
   submittal's local folder couldn't be matched.

Worth running it on the actual nightly schedule (not just manually) for
the first few nights and checking `run.log` each morning, before trusting
it to run unattended for good — that's really the only way to learn how
long your SharePoint session survives overnight unattended.

## Known gaps to sanity-check before trusting this on the live log

- **`AllItems.aspx` URL assumption**: navigation assumes the default
  modern-list view URL pattern. If Jacobs' site uses a custom view, adjust
  `_navigate_to_list()` in `sharepoint_tab_client.py`.
- **Session lifetime**: if the SharePoint session expires between runs,
  the agent fails clearly (not silently) but still needs a human to
  re-log-in the tab. For a nightly unattended run, this is the main risk
  — worth checking the session is still valid each morning until you've
  seen how long it holds.
- **Multiple/zero folder matches**: if `find_submittal_folder()` finds
  more than one folder starting with `Submittal 0XX` (or none), it warns
  and skips that submittal's response rather than guessing — check
  `logs\run.log` for these warnings periodically.
- **"Date responded by Jacobs" could now be automated**: there's a
  `Returned to Contractor Date` field on each SharePoint item that looks
  like a real source for this previously-manual log column. Not wired in
  yet — say the word and I'll map it into `SYNCED_FIELDS`.
- **Whole-file overwrite**: each run rewrites the entire log workbook.
  Not safe if someone has it open and is editing at the exact moment a
  scheduled run fires.
- **Generalizing to RFIs**: this same shape (navigate → read-only REST
  pull → log diff/update → response-file sync → status-breakdown email)
  should carry over to an RFI log with a different list name and folder
  structure once this is validated on submittals.
