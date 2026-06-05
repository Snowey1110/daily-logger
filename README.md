# Daily Logger

Daily Logger is a Windows-first journal app for fast daily notes, recordings, speech-to-text, AI summaries, and book-style review. It stores your journal locally in Excel, opens quickly as a portable app, and keeps large optional features as separate downloads so the normal package stays small.

<p align="center">
  <a href="https://github.com/Snowey1110/daily-logger/releases/latest/download/DailyLoggerPortable.zip"><strong>Download DailyLoggerPortable.zip</strong></a>
</p>

## Downloads

| Package | Use it for | Link |
| --- | --- | --- |
| Daily Logger Portable | Main app with journal, cloud transcription, AI tools, reader, settings, and backups. | [Download](https://github.com/Snowey1110/daily-logger/releases/latest/download/DailyLoggerPortable.zip) |
| Local Transcription Addon | Optional offline Whisper helper for local `tiny`, `base`, `small`, and `medium` transcription models. | [Download](https://github.com/Snowey1110/daily-logger/releases/latest/download/DailyLoggerLocalTranscriptionAddon.zip) |
| Media Tools Addon | Optional video/audio conversion tools for iPhone videos, Voice Memos, large media splitting, and file preparation. | [Download](https://github.com/Snowey1110/daily-logger/releases/latest/download/DailyLoggerMediaToolsAddon.zip) |

The app can also install supported add-ons from **Settings > Download Manager**.

## Screenshots

| Journal Window - Dark | Journal Window - Light |
| --- | --- |
| <img src="images/Dark%20Theme.png" alt="Daily Logger dark journal window" width="620" /> | <img src="images/Light%20Theme.png" alt="Daily Logger light journal window" width="620" /> |

| iPhone Transfer | Download Manager |
| --- | --- |
| <img src="images/iPhone%20QR%20Transfer.png" alt="Sanitized iPhone QR upload window" width="620" /> | <img src="images/Download%20Manager.png" alt="Daily Logger Download Manager for models and add-ons" width="620" /> |

| Console Progress | Virtual Reader |
| --- | --- |
| <img src="images/Console%20Progress.png" alt="Daily Logger console progress log" width="620" /> | <img src="images/Virtual%20Reader%20Spread.png" alt="Virtual Reader double-page journal spread" width="620" /> |

| Virtual Reader Cover |
| --- |
| <img src="images/Virtual%20Reader%20Cover.png" alt="Virtual Reader cover page" width="620" /> |

The public iPhone screenshot hides private network details. Its QR code is a demo code, not a real local upload address.

## Current Features

- **Journal Window** for writing daily entries, recording audio, transcribing speech, restoring drafts, and saving to Excel.
- **Speech to text** with cloud transcription or optional local Whisper models.
- **Transcribe File** for app recordings, audio files, iPhone videos, and Voice Memos.
- **iPhone Inbox** for QR upload and Share Sheet Shortcut uploads while Daily Logger is open.
- **Download Manager** for local transcription models, add-ons, storage visibility, install, uninstall, and default model selection.
- **AI report, AI recap, and chatbot** using an optional OpenAI API key.
- **Virtual Journal Reader** with book-style browsing, journal pages, speech-to-text pages, AI report pages, sketches, images, and page editing.
- **Excel-backed storage** in `Journal.xlsx`, with per-day sheets and a rebuilt `Master Journal`.
- **Background backup** that lets the Journal Window open first, then runs daily backup work after startup.
- **Portable build** that avoids the slow single-file EXE unpack step.

## Quick Start

1. Download [DailyLoggerPortable.zip](https://github.com/Snowey1110/daily-logger/releases/latest/download/DailyLoggerPortable.zip).
2. Extract the zip.
3. Open the extracted `DailyLogger` folder.
4. Run `DailyLogger.exe`.
5. On first launch, choose the app name and startup preference.

No Python installation is required for the portable build. Keep the extracted folder together because `DailyLogger.exe` uses the bundled `_internal` runtime folder next to it.

## Speech To Text

Daily Logger supports two transcription paths:

- **Cloud models**: ready immediately when an OpenAI API key is saved.
- **Local models**: require the Local Transcription Addon and a downloaded model.

The transcription dropdown beside **Transcribe** and **Transcribe File** shows cloud models first, then installed local models. Use **Settings > Download Manager** to download or remove local models and add-ons.

For long videos, install the Media Tools Addon. Daily Logger can convert video to audio, split oversized audio into safe pieces, and append completed transcript parts as progress finishes.

## iPhone Transfer

Daily Logger can receive iPhone media while the app is open.

- **QR upload**: click **Receive from iPhone**, scan the QR code, choose videos, Voice Memos, or audio files, then upload from the phone page.
- **Share Sheet Shortcut**: copy the Shortcut URL from the iPhone Inbox window and use it in an iPhone Shortcut with `Get Contents of URL`.
- **Large videos**: for best reliability, use the Shortcut to encode audio-only M4A first, then upload that audio.
- **Incoming files**: received files are accepted or declined on the PC before transcription begins.

If Media Tools is missing and a video needs conversion, Daily Logger keeps the file pending and prompts for the add-on instead of losing the upload.

## Virtual Journal Reader

Virtual Reader turns `Journal.xlsx` into a local browser-based book view. It opens from the left navigation rail with **Virtual Reader**.

Reader features include:

- Double-page journal spreads.
- Journal, speech-to-text, and AI report bookmarks.
- Per-page sketch and image overlays.
- Inline page editing for text, date, time, sketches, images, and layer order.
- Scrollable right-page content for long journal overflow, transcripts, and AI reports.
- Theme settings, sort order, single-page mode, English/Chinese UI, and optional LAN viewing.

## AI Features

OpenAI is optional. Without a key, local journaling, Excel storage, settings, backups, the reader, and downloaded local transcription still work.

OpenAI-powered features include:

- AI recap over journal entries.
- Chatbot mode.
- AI report generation from journal text and speech-to-text.
- Cloud transcription.

Set a key in the Settings page, with the console command below, or with the `OPENAI_API_KEY` environment variable:

```text
TOKEN ADD [OPENAI API TOKEN]
```

## Console Commands

The Journal Window includes a console for quick actions. Type:

```text
HELP
```

The live help list is the source of truth because commands change with the app. Stable examples include:

| Command | Action |
| --- | --- |
| `J` | Open the Journal Window. |
| `R` / `RT` | Run AI recap, with `RT` using the thinking model. |
| `C` / `CT` | Open chatbot, with `CT` using the thinking model. |
| `RC` / `RECORD` | Start background recording. |
| `RS` / `RECORD STOP` | Stop background recording and save it. |
| `RESTORE` | Reopen the latest unsaved Journal Window draft. |
| `OPEN JOURNAL` | Open `Journal.xlsx`. |
| `OPEN DIRECTORY` | Open the Daily Logger app-data folder. |
| `BACKUP START` | Create a backup now. |
| `TOKEN ADD`, `TOKEN RESET`, `TOKEN COPY` | Manage the saved OpenAI API key. |
| `LANGUAGE English`, `LANGUAGE Chinese` | Switch UI language. |

## Data And Storage

Daily Logger stores user data outside the repo so source runs and portable builds share the same files:

```text
%APPDATA%\DailyLogger\
```

Important folders and files:

- `daily_logs/Journal.xlsx` - main Excel journal workbook.
- `daily_logs/Recording/` - saved app recordings and imported media.
- `daily_logs/Recording/iPhone Inbox/` - accepted iPhone uploads waiting for processing.
- `daily_logs/backup/` - backup zip files.
- `settings/daily_logger_prefs.json` - local preferences.
- `settings/journal_window_draft.json` - autosaved unsaved draft.
- `settings/journal_reader_sketches.json` - Virtual Reader page overlays.
- `settings/daily_logger_api_key.txt` - optional saved OpenAI API key.
- `addons/` - optional downloaded add-on runtimes.
- `models/whisper/` - optional downloaded local transcription models.

## Build From Source

Run the app from source:

```bash
python daily_logger.py
```

Build the Virtual Reader web UI:

```bash
cd virtual-journal-reader
npm install
npm run build
```

Build release artifacts:

```bash
python -m PyInstaller DailyLogger.spec
python build_local_transcription_addon.py
python build_media_tools_addon.py
```

## Project Layout

```text
daily_logger.py                      Main desktop app
journal_i18n.py                      English/Chinese UI text
local_transcriber_helper.py          Optional local transcription helper source
DailyLogger.spec                     Portable-folder PyInstaller build
build_local_transcription_addon.py   Local transcription add-on builder
build_media_tools_addon.py           Media Tools add-on builder
virtual-journal-reader/              React/Vite reader UI and server
images/                              README screenshots
```

## Notes

- Daily Logger is designed primarily for Windows.
- Settings, generated journals, backups, recordings, models, add-ons, and API keys are local user data and should not be committed.
- The normal portable package is cloud-ready and small; local transcription and media conversion are optional add-ons.
