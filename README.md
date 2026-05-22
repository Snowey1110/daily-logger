# Daily Logger

Daily Logger is a Windows-first personal journal for fast daily capture, rich review, and Excel-backed storage. It opens into a desktop Journal Window, saves entries locally, and can optionally use OpenAI features for recap, chat, transcription, and AI reports.

<p align="center">
  <a href="https://github.com/Snowey1110/daily-logger/releases/latest/download/DailyLogger.exe"><strong>Download the latest DailyLogger.exe</strong></a>
</p>

## Highlights

- Journal Window for quick writing, recording, transcription, and AI report generation.
- Virtual Journal Reader for browsing entries as a book in the browser.
- Editable reader pages with text, sketch, image, and layer-order controls.
- Excel storage in `Journal.xlsx` with per-day sheets and a rebuilt `Master Journal`.
- Background daily backup after the journal UI finishes loading.
- Optional OpenAI recap/chat tools; normal journaling works without an API key.

## Screenshots

| Journal Window - Dark Theme | Journal Window - Light Theme |
| --- | --- |
| <img src="images/Dark%20Theme.png" alt="Daily Logger current dark theme journal window" width="520" /> | <img src="images/Light%20Theme.png" alt="Daily Logger current light theme journal window" width="520" /> |

| Virtual Reader Cover | Virtual Reader Double-Page Journal |
| --- | --- |
| <img src="images/Virtual%20Reader%20Cover.png" alt="Actual Virtual Reader cover page" width="520" /> | <img src="images/Virtual%20Reader%20Spread.png" alt="Actual Virtual Reader double-page journal spread with text, sketch, and image layers" width="520" /> |

## Virtual Journal Reader

Virtual Journal Reader turns `Journal.xlsx` into a local browser-based book view. It opens from the Journal Window navigation rail with **Virtual Reader**, or from `launch_journal_reader.bat`.

Reader features:

- Book-style browsing with journal, speech-to-text, and AI report sections.
- Inline page editing for journal text, date, and time.
- Sketch layer with color, line width, eraser, undo/redo, and Shift straight-line drawing.
- Image layer with upload/paste, drag, resize, delete, and layer ordering.
- Theme settings, sort order, single-page mode, English/Chinese UI, and optional LAN viewing.

How the reader is found:

- PyInstaller builds can bundle `virtual-journal-reader/dist` and `serve_reader.py` through `DailyLogger.spec`.
- Source runs use `virtual-journal-reader/serve_reader.py` and `virtual-journal-reader/dist/`.
- If `virtual-journal-reader.zip` is copied next to the app, Daily Logger can extract it automatically the first time Virtual Reader opens.
- EXE builds launch the reader server through Daily Logger's internal `--serve-virtual-reader` mode, so opening the reader does not create another Daily Logger window.

The local reader server uses port `8765` by default. If an old reader server is already running, Daily Logger checks its health and refuses stale builds instead of silently opening the wrong UI.

## Quick Start

### Run the EXE

1. Download [DailyLogger.exe](https://github.com/Snowey1110/daily-logger/releases/latest/download/DailyLogger.exe).
2. Run the executable.
3. On first launch, choose an app name and whether to start with Windows.

No Python installation is required for the packaged EXE.

### Run from Source

```bash
python daily_logger.py
```

Or on Windows:

```text
launch_daily_logger.bat
```

Daily Logger checks required and optional Python packages at startup. Required packages are needed to run; optional packages enable features such as microphone recording and the calendar date picker.

### Build the Virtual Reader UI

Only needed when developing or rebuilding the browser reader:

```bash
cd virtual-journal-reader
npm install
npm run build
```

## Console Commands

The Journal Window has a built-in console. Press `H` or `HELP` to see the current command list.

| Command | Action |
| --- | --- |
| `J` | Open journal flow. |
| `J SETTINGS`, `J SETTING`, `JOURNAL SETTINGS`, `JS` | Open the journal command menu. |
| `R`, `RT` | Run AI recap, with `RT` using the thinking model. |
| `R [date range]`, `RT [date range]` | Recap entries within a date range, such as `4/27 - 4/30`. |
| `R [file]`, `RT [file]` | Recap using file text as context. |
| `C`, `CT` | Open chatbot, with `CT` using the thinking model. |
| `RESTORE` | Reopen the latest unsaved Journal Window draft. |
| `OPEN DIRECTORY` | Open the app data folder. |
| `OPEN JOURNAL` | Open `Journal.xlsx`. |
| `OPEN SCREENSHOTS` | Open the chat screenshots folder. |
| `STARTUP TRUE`, `STARTUP FALSE` | Enable or disable launch at Windows sign-in. |
| `DEFAULT WINDOWS`, `DEFAULT CONSOLE` | Choose whether `J` opens the window directly or shows journal choices. |
| `BACKUP START` | Create a backup zip now. |
| `BACKUP TRUE`, `BACKUP FALSE`, `BACKUP LIMITED` | Enable, disable, or limit automatic backups. |
| `TOKEN ADD [token]`, `TOKEN RESET`, `TOKEN COPY` | Manage the stored OpenAI API key. |
| `LAN cn`, `LAN en`, `LANGUAGE Chinese`, `LANGUAGE English` | Switch UI language. |
| `SB bat`, `SB journal`, `SB reader` | Create Start Menu search shortcuts. |
| `WIFI WARN [name]` | Warn when connected to the named Wi-Fi network. |
| `RENAME` | Change the app name shown in the UI. |
| `TS` | Take a screenshot for chat workflows. |
| `UNINSTALL`, `CONFIRM UNINSTALL` | Request and confirm local app-data cleanup. |

## Data and Storage

Daily Logger stores user data outside the repo so EXE and source runs share the same files:

```text
%APPDATA%\DailyLogger\
```

Important files and folders:

- `daily_logs/Journal.xlsx` - main Excel journal file.
- `daily_logs/Recording/` - saved recording files.
- `daily_logs/backup/` - backup zip files.
- `settings/daily_logger_prefs.json` - local preferences.
- `settings/journal_window_draft.json` - autosaved unsaved draft.
- `settings/journal_reader_sketches.json` - reader sketches and page overlays.
- `settings/daily_logger_api_key.txt` - optional saved OpenAI API key.

Journal workbook behavior:

- Entries save to date sheets such as `2026-05-22`.
- `Master Journal` is rebuilt from date sheets.
- Date sheets are ordered newest to oldest behind `Master Journal`.
- Old repo-local `daily_logs/` and `settings/` data can be migrated into `%APPDATA%\DailyLogger`.

Backup behavior:

- Automatic backup runs once per new day when enabled.
- Startup is kept responsive: the journal window opens first, then backup runs in the background.
- Limited backup mode keeps a small rotating set of backup zip files.

## OpenAI Features

OpenAI is optional. Without a key, local journaling, Excel storage, the reader, settings, backups, and shortcuts still work.

OpenAI-powered features include:

- AI recap over journal entries.
- General chatbot mode.
- AI report generation from journal text and speech-to-text.
- Speech transcription when recording dependencies are installed.

Set a key with:

```text
TOKEN ADD [OPENAI API TOKEN]
```

You can also use the `OPENAI_API_KEY` environment variable.

## Project Layout

```text
daily_logger.py                 Main desktop app
journal_i18n.py                 Journal Window translations
launch_daily_logger.bat         Windows app launcher
launch_journal_reader.bat       Virtual Reader launcher
DailyLogger.spec                PyInstaller build spec
dist/DailyLogger.exe            Packaged executable build
virtual-journal-reader/         React/Vite reader UI and server
virtual-journal-reader.zip      Runnable reader add-on package
images/                         README screenshots and feature mockups
```

## Notes

- This project is designed primarily for Windows.
- Settings, generated journals, backups, recordings, and API keys are local user data and should not be committed.
- The reader is served locally by default; LAN viewing can be toggled from the reader settings when needed.
