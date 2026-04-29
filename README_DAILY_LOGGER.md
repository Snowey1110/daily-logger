# Daily Logger Setup Guide (Windows + Excel)

## What this app does
- Opens a Command Prompt menu:
  - `J` for Journal
  - `R` for Recap (ask questions using your journal)
  - `RT` for Recap using thinking model
  - `C` for Chatbot (general chat)
  - `CT` for Chatbot using thinking model
  - `H` to list commands
  - `restore` to reopen latest unsaved journal window draft
  - `startup true` / `startup false` to manage Windows startup
  - `Enter` to skip/exit
- Saves data to Excel files in `daily_logs/`.
- Supports flexible date input such as:
  - `04/20/2026`
  - `4/26`
  - `Apr 26`
  - `April 26`

Missing year defaults to the current year.

## Files in this project
- `daily_logger.py` - main script
- `launch_daily_logger.bat` - easy launcher for Windows
- `daily_logs/` - generated folder with Excel files

## 1) Install requirements (one-time)
1. Install Python 3 (if needed).
2. Open Command Prompt in this folder.
3. Run:
   - `pip install openpyxl`
4. First time you use `R`, `C`, or `CT`, the app asks for your OpenAI API key and saves it locally in `settings/daily_logger_api_key.txt`.
   - The key file is git-ignored so you can share the project safely.
   - Optional override: you can still set `OPENAI_API_KEY` in environment.

## 2) Run the logger manually
- Double-click `launch_daily_logger.bat`

Or from Command Prompt:
- `python daily_logger.py`

On first run, the app asks what to name itself. Press Enter to keep the default `Daily Logger`.
The name is saved in `settings/daily_logger_prefs.json` (git-ignored).
Use `RENAME` (or quick `RENAME <name>` / `REANAME <name>`) from the command line to change it later.
On first run, the app also asks whether to open automatically on computer startup.

## 3) Start automatically when you log in
- In app command page:
  - `startup true` to enable
  - `startup false` to disable
- This manages a Windows Startup shortcut to `launch_daily_logger.bat`.

## How data is stored

### Journal
- File: `daily_logs/Journal.xlsx`
- Sheets:
  - `Master Journal` (always first sheet)
  - One date sheet per day, named like `2026-04-28`

Journal sync behavior:
- New journal entries are written to that date's page.
- `Master Journal` is rebuilt from date pages.
- If you delete a date page, those entries are removed from `Master Journal` on next launch/save.
- Date pages are reordered newest to oldest behind `Master Journal`.

## Input behavior
- **Date prompt:** `Entry date (mm/dd/yyyy, Enter for today <today's date>):`
- **Time prompt:** type `rn` (or `RN`) to use the current time; press Enter to save as `N/A`.
- **Journal delete shortcut:** at `What happened today?`, enter `DP` to delete the previous journal entry (with confirmation). If that was the only entry on its date page, that page is removed.
- **Journal window mode:** at `What happened today?`, enter `window` to open a GUI editor with text/date/time and image-path attach support.
- **Draft backup + restore:** journal window drafts auto-save to `settings/journal_window_draft.json`; use main command `restore` to reopen latest unsaved draft.
- **Recap mode (`R`):** asks ChatGPT questions using all journal entries as context; prompt text is `Recap:` and empty input exits.
- **Thinking recap (`RT`):** same as `R` but uses the thinking model.
- **Chatbot mode (`C`):** asks ChatGPT without journal context; press Enter on empty question to exit chat mode.
- **Chatbot commands (`C`/`CT`):** `help` shows chat commands locally (no AI call), `ts` takes a screenshot and attaches it to your next AI message, `rs` removes a pending screenshot attachment.
- **Thinking chatbot (`CT`):** same as `C` but uses the thinking model.
- After saving an entry, the app returns directly to the main menu.

## Add modules later
To add another module:
1. Add a prompt function that returns a row list.
2. Add a `ModuleConfig` entry in `MODULES` with:
   - workbook name
   - sheet name
   - headers
   - prompt function

Each module can use its own Excel file.
