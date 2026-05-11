from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
import base64
import contextlib
import ctypes
import io
import importlib
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import wave
from typing import Any, Callable, Dict, List, Optional, Tuple, Union
from urllib import error, request
import zipfile

from journal_i18n import UI_LANGUAGE_PREF_KEY, normalize_ui_language, ui_translate

_journal_ui_language_changed_hook: Optional[Callable[[str], None]] = None


def set_journal_ui_language_changed_hook(hook: Optional[Callable[[str], None]]) -> None:
    global _journal_ui_language_changed_hook
    _journal_ui_language_changed_hook = hook


try:
    from openpyxl import Workbook, load_workbook
except Exception:
    Workbook = None  # type: ignore[assignment]
    load_workbook = None  # type: ignore[assignment]
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception:
    tk = None
    filedialog = None
    messagebox = None
    ttk = None
try:
    from tkcalendar import Calendar, DateEntry
except Exception:
    Calendar = None  # type: ignore[assignment, misc]
    DateEntry = None  # type: ignore[assignment]
try:
    import msvcrt
except Exception:
    msvcrt = None
try:
    import readline as _readline  # type: ignore[assignment]
except ImportError:
    _readline = None


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

def get_user_data_root() -> Path:
    """Return a stable per-user storage root shared across EXE and source runs."""
    appdata = os.getenv("APPDATA", "").strip()
    if appdata:
        return Path(appdata) / "DailyLogger"
    return BASE_DIR


USER_DATA_ROOT = get_user_data_root()
DATA_DIR = USER_DATA_ROOT / "daily_logs"
RECORDING_DIR = DATA_DIR / "Recording"
BACKUP_DIR = DATA_DIR / "backup"
SETTINGS_DIR = USER_DATA_ROOT / "settings"
LEGACY_DATA_DIR = BASE_DIR / "daily_logs"
LEGACY_SETTINGS_DIR = BASE_DIR / "settings"
MASTER_JOURNAL_SHEET = "Master Journal"
JOURNAL_HEADERS_LEGACY = ["Date", "Time", "Journal"]
OPENAI_CHAT_COMPLETIONS_URL = "https://api.openai.com/v1/chat/completions"
OPENAI_TRANSCRIPTION_URL = os.getenv(
    "OPENAI_TRANSCRIPTION_URL", "https://api.openai.com/v1/audio/transcriptions"
).strip()
LIVE_STT_CHUNK_INTERVAL_SEC = 5.0
LIVE_STT_MIN_CHUNK_SAMPLES = int(16000 * 0.4)
# Journal waveform: int16 PCM RMS soft noise floor and display scale.
WAVEFORM_RMS_NOISE_FLOOR = 40.0
WAVEFORM_MAX_DRAW_SAMPLES = 4000
WAVEFORM_RMS_NORM = 6000.0
# Smaller input blocks when metering so the canvas updates often enough to feel live.
WAVEFORM_INPUT_BLOCK_SAMPLES = 512
# Journal STT / AI report: same button width (text units) and grid min width so text areas align.
JOURNAL_SIDE_ACTION_BTN_WIDTH_CH = 16
JOURNAL_SIDE_ACTION_GRID_MINSIZE = 130
# Whisper list price per audio minute (USD); verify at https://openai.com/pricing
WHISPER_USD_PER_MIN = 0.006
# Pre-send WAV cleanup (RMS on int16-scale, same ballpark as WAVEFORM_RMS_NOISE_FLOOR).
WHISPER_PRE_FRAME_MS = 25
WHISPER_PRE_SILENCE_RMS = 32.0
WHISPER_PRE_EDGE_PAD_MS = 120
WHISPER_PRE_MIN_SPEECH_MS = 50
WHISPER_PRE_MAX_INTERNAL_SILENCE_SEC = 1.25
WHISPER_PRE_KEEP_INTERNAL_SILENCE_SEC = 0.35
WHISPER_TRANSCRIBE_CHUNK_SEC = 8 * 60
# OpenAI Whisper multipart limit is ~25 MiB total; keep each mono int16 chunk smaller than that.
WHISPER_SAFE_CHUNK_PCM_BYTES = 20 * 1024 * 1024
WHISPER_SKIP_SINGLE_FILE_BYTES = 22 * 1024 * 1024
WHISPER_TRANSCRIBE_PROMPT_CHAR_LIMIT = 600
# Hover tooltips: narrow wrap → shorter line length, more lines (taller block).
TOOLTIP_WRAP_PX = 220
TOOLTIP_WRAP_PX_MAX = 280
JOURNAL_PREF_THEME_KEY = "journal_window_theme"


def migrate_legacy_storage_if_needed() -> None:
    """One-time best-effort migration from legacy BASE_DIR storage to USER_DATA_ROOT."""
    if USER_DATA_ROOT.resolve() == BASE_DIR.resolve():
        return
    try:
        USER_DATA_ROOT.mkdir(parents=True, exist_ok=True)
    except OSError:
        return

    def _journal_has_entries(path: Path) -> bool:
        if not path.exists():
            return False
        if load_workbook is None:
            # Fallback heuristic when openpyxl is unavailable.
            return path.stat().st_size > 16 * 1024
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            ws = wb[MASTER_JOURNAL_SHEET] if MASTER_JOURNAL_SHEET in wb.sheetnames else wb.active
            max_row = int(ws.max_row or 0)
            if max_row <= 1:
                wb.close()
                return False
            for row in ws.iter_rows(min_row=2, values_only=True):
                if any((str(cell).strip() if cell is not None else "") for cell in row):
                    wb.close()
                    return True
            wb.close()
            return False
        except Exception:
            return path.stat().st_size > 16 * 1024

    # Migrate daily logs. If new journal exists but is empty, replace it with legacy journal.
    try:
        legacy_journal = LEGACY_DATA_DIR / "Journal.xlsx"
        new_journal = DATA_DIR / "Journal.xlsx"
        if legacy_journal.exists() and (
            not new_journal.exists()
            or (not _journal_has_entries(new_journal) and _journal_has_entries(legacy_journal))
        ):
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            shutil.copy2(legacy_journal, new_journal)
        legacy_backup = LEGACY_DATA_DIR / "backup"
        new_backup = DATA_DIR / "backup"
        if legacy_backup.exists() and not new_backup.exists():
            shutil.copytree(legacy_backup, new_backup)
        legacy_recording = LEGACY_DATA_DIR / "Recording"
        new_recording = DATA_DIR / "Recording"
        if legacy_recording.exists() and not new_recording.exists():
            shutil.copytree(legacy_recording, new_recording)
    except OSError:
        pass

    # Migrate settings files only when target files are missing.
    try:
        if LEGACY_SETTINGS_DIR.exists():
            SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
            for src in LEGACY_SETTINGS_DIR.glob("*"):
                if not src.is_file():
                    continue
                dst = SETTINGS_DIR / src.name
                if not dst.exists():
                    shutil.copy2(src, dst)
    except OSError:
        pass


@dataclass(frozen=True)
class JournalWindowThemeSpec:
    """Colors and layout for the journal Tk window (light vs dark)."""

    id: str
    toggle_label: str
    surface: str
    panel: str
    field: str
    text: str
    muted: str
    accent: str
    border: str
    waveform: str
    btn_secondary: str
    btn_disabled: str
    disabled_fg: str
    hover_primary: str
    hover_save: str
    secondary_hover: str
    pad_outer: int
    pad_top_y: Tuple[int, int]
    pad_center_y: int
    pad_button_y: int
    date_label_font: Tuple[Any, ...]
    section_label_font: Tuple[Any, ...]
    is_dark: bool

    def toolbar_btn_config(self) -> Tuple[str, str, str, str]:
        """bg, fg, activebackground, activeforeground for Update Time / Open."""
        if self.is_dark:
            return (self.btn_secondary, self.text, self.accent, "white")
        return (self.btn_secondary, self.text, self.secondary_hover, self.text)

    def toolbar_hover(self) -> Tuple[str, str]:
        if self.is_dark:
            return (self.hover_primary, "white")
        return (self.secondary_hover, self.text)

    def toolbar_bind_rest(self) -> Tuple[str, str, str, str, str]:
        bg, fg, abg, afg = self.toolbar_btn_config()
        return ("normal", bg, fg, abg, afg)

    def side_action_config(self) -> Tuple[str, str, str, str]:
        """bg, fg, activebackground, activeforeground when action is enabled."""
        if self.is_dark:
            return (self.btn_secondary, self.text, self.accent, "white")
        return (self.accent, "white", self.hover_primary, "white")

    def side_action_bind_rest(self) -> Tuple[str, str, str, str, str]:
        bg, fg, abg, afg = self.side_action_config()
        return ("normal", bg, fg, abg, afg)

    def side_action_disabled(self) -> Tuple[str, str, str, str, str]:
        if self.is_dark:
            return (
                "disabled",
                self.btn_disabled,
                self.muted,
                self.btn_secondary,
                self.text,
            )
        return (
            "disabled",
            self.btn_disabled,
            self.disabled_fg,
            self.hover_primary,
            "white",
        )

    def transcribe_busy_config(self) -> Tuple[str, str, str, str, str]:
        if self.is_dark:
            return (
                self.btn_disabled,
                self.muted,
                self.btn_secondary,
                self.text,
                self.muted,
            )
        return (
            self.btn_disabled,
            self.disabled_fg,
            self.secondary_hover,
            self.text,
            self.disabled_fg,
        )

    def transcribe_idle_disabled_config(self) -> Tuple[str, str, str, str, str]:
        return self.transcribe_busy_config()

    def gen_bind_rest(self) -> Tuple[str, str, str, str, str]:
        if self.is_dark:
            bg, fg, abg, afg = self.side_action_config()
            return ("normal", bg, fg, abg, afg)
        return ("normal", self.accent, "white", self.hover_primary, "white")

    def gen_bind_disabled(self) -> Tuple[str, str, str, str, str]:
        if self.is_dark:
            return (
                "disabled",
                self.btn_disabled,
                self.muted,
                self.btn_secondary,
                self.text,
            )
        return (
            "disabled",
            self.btn_disabled,
            self.disabled_fg,
            self.hover_primary,
            "white",
        )

    def save_bind_disabled(self) -> Tuple[str, str, str, str, str]:
        return self.gen_bind_disabled()

    def ttk_combobox_kwargs(self) -> Dict[str, Any]:
        if self.is_dark:
            return {
                "fieldbackground": self.field,
                "background": self.panel,
                "foreground": self.text,
                "bordercolor": self.border,
                "lightcolor": self.panel,
                "darkcolor": self.field,
                "arrowcolor": self.muted,
                "padding": 4,
            }
        return {
            "fieldbackground": self.field,
            "background": self.btn_secondary,
            "foreground": self.text,
        }


JOURNAL_THEME_LIGHT = JournalWindowThemeSpec(
    id="light",
    toggle_label="Dark mode",
    surface="#F2F2F7",
    panel="#FFFFFF",
    field="#FFFFFF",
    text="#1D1D1F",
    muted="#6E6E73",
    accent="#0071E3",
    border="#D2D2D7",
    waveform="#0071E3",
    btn_secondary="#E8E8ED",
    btn_disabled="#E5E5EA",
    disabled_fg="#AEAEB2",
    hover_primary="#0077ED",
    hover_save="#0077ED",
    secondary_hover="#DCDCE0",
    pad_outer=14,
    pad_top_y=(14, 10),
    pad_center_y=10,
    pad_button_y=14,
    date_label_font=("Segoe UI", 10, "bold"),
    section_label_font=("Segoe UI", 10, "bold"),
    is_dark=False,
)

JOURNAL_THEME_DARK = JournalWindowThemeSpec(
    id="dark",
    toggle_label="Light mode",
    surface="#06060C",
    panel="#14141E",
    field="#0A0A12",
    text="#F5F5F7",
    muted="#98989D",
    accent="#0A84FF",
    border="#2C2C38",
    waveform="#64D2FF",
    btn_secondary="#24243A",
    btn_disabled="#101018",
    disabled_fg="#98989D",
    hover_primary="#339CFF",
    hover_save="#5CB0FF",
    secondary_hover="#339CFF",
    # Keep geometry/font metrics identical to light mode to avoid text reflow/shift
    # when toggling themes; only colors should differ between modes.
    pad_outer=14,
    pad_top_y=(14, 10),
    pad_center_y=10,
    pad_button_y=14,
    date_label_font=("Segoe UI", 10, "bold"),
    section_label_font=("Segoe UI", 10, "bold"),
    is_dark=True,
)


def normalize_journal_window_theme_key(raw: str) -> str:
    k = (raw or "").strip().lower()
    return "dark" if k == "dark" else "light"


def load_journal_window_theme_spec() -> JournalWindowThemeSpec:
    prefs = load_preferences()
    return (
        JOURNAL_THEME_DARK
        if normalize_journal_window_theme_key(prefs.get(JOURNAL_PREF_THEME_KEY, "light"))
        == "dark"
        else JOURNAL_THEME_LIGHT
    )


OPENAI_MODEL = "gpt-4o-mini"
OPENAI_THINKING_MODEL = "gpt-5.5"
API_KEY_FILE = SETTINGS_DIR / "daily_logger_api_key.txt"
PREFS_FILE = SETTINGS_DIR / "daily_logger_prefs.json"
WIFI_WARN_FILE = SETTINGS_DIR / "wifi_warn_list.json"
JOURNAL_WINDOW_DRAFT_FILE = SETTINGS_DIR / "journal_window_draft.json"
JOURNAL_WINDOW_CONSOLE_RESERVE_BOTTOM = 56
SCREENSHOT_DIR = DATA_DIR / "chat_screenshots"
STARTUP_SHORTCUT_NAME = "Daily Logger.lnk"


@dataclass
class ModuleConfig:
    name: str
    workbook_name: str
    sheet_name: str
    headers: List[str]
    prompt_builder: Callable[[], Optional[List[str]]]


PENDING_UNINSTALL_CONFIRM = False


def bind_openpyxl_symbols() -> bool:
    global Workbook, load_workbook
    try:
        openpyxl_module = importlib.import_module("openpyxl")
        Workbook = openpyxl_module.Workbook
        load_workbook = openpyxl_module.load_workbook
        return True
    except Exception:
        Workbook = None  # type: ignore[assignment]
        load_workbook = None  # type: ignore[assignment]
        return False


def _pip_install_packages(packages: List[str]) -> bool:
    if not packages:
        return True
    print("Installing:", ", ".join(packages))
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", *packages],
        capture_output=False,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        print("Package installation failed. Try installing manually:")
        print(f"  {sys.executable} -m pip install {' '.join(packages)}")
        return False
    return True


def _missing_modules(specs: List[Tuple[str, str]]) -> List[str]:
    """Return pip package names whose import modules are not available."""
    return [
        pip_name
        for module_name, pip_name in specs
        if importlib.util.find_spec(module_name) is None
    ]


def ensure_runtime_dependencies() -> bool:
    core_specs: List[Tuple[str, str]] = [
        ("openpyxl", "openpyxl"),
        ("mss", "mss"),
    ]
    optional_specs: List[Tuple[str, str, str]] = [
        ("sounddevice", "sounddevice", "microphone recording for journal speech-to-text"),
        ("numpy", "numpy", "audio buffers for journal speech-to-text"),
        ("tkcalendar", "tkcalendar", "calendar popup on the journal date field"),
    ]

    missing_core = _missing_modules(core_specs)
    optional_missing = [
        (pip_name, blurb)
        for module_name, pip_name, blurb in optional_specs
        if importlib.util.find_spec(module_name) is None
    ]

    if missing_core or optional_missing:
        if missing_core:
            print("Required packages:")
            for pip_name in missing_core:
                print(f"  - {pip_name}")
        if optional_missing:
            print("Optional packages for full journal window features:")
            for pip_name, blurb in optional_missing:
                print(f"  - {pip_name}: {blurb}")
        print("Install missing packages now? (y/N): ", end="")
        answer = input().strip().lower()
        if answer in ("y", "yes"):
            install_list = list(missing_core)
            install_list.extend([pip_name for pip_name, _blurb in optional_missing])
            if install_list and not _pip_install_packages(install_list):
                print("You can install them later with:")
                print(f"  {sys.executable} -m pip install {' '.join(install_list)}")
                if missing_core:
                    return False
        else:
            if missing_core:
                print("Skipped installation of required packages.")
            if optional_missing:
                print("Skipped optional packages. Speech-to-text needs sounddevice and numpy.")

        missing_core = _missing_modules(core_specs)
        if missing_core:
            print(
                "Cannot start: still missing "
                + ", ".join(missing_core)
                + ". Install them, then run the app again."
            )
            return False

    if not bind_openpyxl_symbols():
        print("openpyxl is required to run this app. Please install it and retry.")
        return False

    return True


def is_enter_equivalent(value: str) -> bool:
    return not value or value.upper() == "X"


def _normalize_journal_header_cell(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    return str(value).strip()


def migrate_journal_workbook_columns_if_needed(wb, new_headers: List[str]) -> bool:
    """Expand legacy 3-column journal sheets to five columns without deleting data."""
    if len(new_headers) != 5:
        return False
    legacy = JOURNAL_HEADERS_LEGACY
    modified = False
    for ws in wb.worksheets:
        if ws.max_row < 1:
            continue
        first_three = [_normalize_journal_header_cell(ws.cell(row=1, column=col).value) for col in (1, 2, 3)]
        if first_three != legacy:
            continue
        if ws.max_column == 3:
            ws.insert_cols(4, amount=2)
            ws.cell(row=1, column=4, value=new_headers[3])
            ws.cell(row=1, column=5, value=new_headers[4])
            modified = True
            continue
        d1 = _normalize_journal_header_cell(ws.cell(row=1, column=4).value)
        e1 = _normalize_journal_header_cell(ws.cell(row=1, column=5).value)
        want_d = new_headers[3].strip()
        want_e = new_headers[4].strip()
        if d1 != want_d or e1 != want_e:
            ws.cell(row=1, column=4, value=new_headers[3])
            ws.cell(row=1, column=5, value=new_headers[4])
            modified = True
    return modified


def red_text(value: str) -> str:
    return f"\033[31m{value}\033[0m"


def save_workbook_with_retry(wb, workbook_path: Path) -> None:
    while True:
        try:
            wb.save(workbook_path)
            return
        except PermissionError:
            print(
                f"Cannot save '{workbook_path.name}' because it is open in another program."
            )
            input("Close the file and press Enter to retry saving...")


def load_workbook_with_retry(workbook_path: Path):
    while True:
        try:
            return load_workbook(workbook_path)
        except PermissionError:
            print(
                f"Cannot open '{workbook_path.name}' because access is blocked (open/locked by another program or sync)."
            )
            input("Close the file or wait for sync, then press Enter to retry...")


def ensure_workbook(module: ModuleConfig) -> Path:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    workbook_path = DATA_DIR / module.workbook_name

    if not workbook_path.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = module.sheet_name
        ws.append(module.headers)
        save_workbook_with_retry(wb, workbook_path)
        return workbook_path

    wb = load_workbook_with_retry(workbook_path)
    if module.name == "Journal":
        if migrate_journal_workbook_columns_if_needed(wb, module.headers):
            save_workbook_with_retry(wb, workbook_path)
    if module.sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(module.sheet_name)
        ws.append(module.headers)
        save_workbook_with_retry(wb, workbook_path)
    else:
        ws = wb[module.sheet_name]
        if ws.max_row == 1:
            # Keep file resilient in case user removed headers manually.
            first_row = [cell.value for cell in ws[1]]
            if first_row != module.headers:
                ws.delete_rows(1, ws.max_row)
                ws.append(module.headers)
                save_workbook_with_retry(wb, workbook_path)

    return workbook_path


def append_row(module: ModuleConfig, row: List[str]) -> None:
    workbook_path = ensure_workbook(module)
    wb = load_workbook_with_retry(workbook_path)

    if module.name == "Journal":
        row_list = list(row)
        while len(row_list) < len(module.headers):
            row_list.append("")
        row = row_list[: len(module.headers)]
        daily_ws = get_or_create_journal_daily_sheet(wb, module, row[0])
        target_row = find_first_empty_data_row(daily_ws, len(module.headers))
        for col_index, value in enumerate(row, start=1):
            daily_ws.cell(row=target_row, column=col_index, value=value)

        rebuild_master_journal_from_daily_pages(wb, module)
        reorder_journal_sheets(wb)
        save_workbook_with_retry(wb, workbook_path)
        return

    ws = wb[module.sheet_name]
    target_row = find_first_empty_data_row(ws, len(module.headers))
    for col_index, value in enumerate(row, start=1):
        ws.cell(row=target_row, column=col_index, value=value)
    save_workbook_with_retry(wb, workbook_path)


def ensure_master_journal_sheet(wb, module: ModuleConfig):
    if MASTER_JOURNAL_SHEET in wb.sheetnames:
        ws = wb[MASTER_JOURNAL_SHEET]
    elif module.sheet_name in wb.sheetnames:
        ws = wb[module.sheet_name]
        ws.title = MASTER_JOURNAL_SHEET
    else:
        ws = wb.create_sheet(MASTER_JOURNAL_SHEET, 0)

    ensure_headers(ws, module.headers)
    return ws


def get_or_create_journal_daily_sheet(wb, module: ModuleConfig, date_value: str):
    date_obj = datetime.strptime(date_value, "%m/%d/%Y")
    daily_sheet_name = date_obj.strftime("%Y-%m-%d")
    if daily_sheet_name in wb.sheetnames:
        ws = wb[daily_sheet_name]
    else:
        ws = wb.create_sheet(daily_sheet_name)
    ensure_headers(ws, module.headers)
    return ws


def ensure_headers(ws, headers: List[str]) -> None:
    first_row = [ws.cell(row=1, column=index).value for index in range(1, len(headers) + 1)]
    normalized = [cell.strip() if isinstance(cell, str) else cell for cell in first_row]
    if normalized != headers:
        if ws.max_row > 0:
            ws.delete_rows(1, ws.max_row)
        ws.append(headers)


def reorder_journal_sheets(wb) -> None:
    if MASTER_JOURNAL_SHEET not in wb.sheetnames:
        return

    master = wb[MASTER_JOURNAL_SHEET]
    dated_sheets = []
    for sheet in wb.worksheets:
        if sheet.title == MASTER_JOURNAL_SHEET:
            continue
        try:
            sheet_date = datetime.strptime(sheet.title, "%Y-%m-%d")
            dated_sheets.append((sheet_date, sheet))
        except ValueError:
            continue

    ordered = sorted(dated_sheets, key=lambda item: item[0], reverse=True)
    ordered_daily = [item[1] for item in ordered]
    remaining = [
        sheet
        for sheet in wb._sheets
        if sheet is not master and sheet not in ordered_daily
    ]
    wb._sheets = [master] + ordered_daily + remaining


def rebuild_master_journal_from_daily_pages(wb, module: ModuleConfig) -> None:
    master_ws = ensure_master_journal_sheet(wb, module)
    entries: List[Tuple[datetime, int, List[str]]] = []

    for sheet in wb.worksheets:
        if sheet.title == MASTER_JOURNAL_SHEET:
            continue
        try:
            sheet_date = datetime.strptime(sheet.title, "%Y-%m-%d")
        except ValueError:
            continue

        for row_index in range(2, sheet.max_row + 1):
            values = [
                sheet.cell(row=row_index, column=col).value
                for col in range(1, len(module.headers) + 1)
            ]
            if is_row_empty(values):
                continue
            normalized = ["" if value is None else str(value) for value in values]
            entries.append((sheet_date, row_index, normalized))

    entries.sort(key=lambda item: (item[0], item[1]), reverse=True)

    if master_ws.max_row > 1:
        master_ws.delete_rows(2, master_ws.max_row - 1)

    for _, _, row_values in entries:
        master_ws.append(row_values)


def delete_latest_journal_entry() -> bool:
    module = MODULES["J"]
    workbook_path = ensure_workbook(module)
    wb = load_workbook_with_retry(workbook_path)

    latest_sheet = None
    latest_date = None
    for sheet in wb.worksheets:
        if sheet.title == MASTER_JOURNAL_SHEET:
            continue
        try:
            sheet_date = datetime.strptime(sheet.title, "%Y-%m-%d")
        except ValueError:
            continue
        if latest_date is None or sheet_date > latest_date:
            latest_date = sheet_date
            latest_sheet = sheet

    if latest_sheet is None:
        return False

    latest_row = None
    for row_index in range(latest_sheet.max_row, 1, -1):
        values = [
            latest_sheet.cell(row=row_index, column=col).value
            for col in range(1, len(module.headers) + 1)
        ]
        if not is_row_empty(values):
            latest_row = row_index
            break

    if latest_row is None:
        return False

    latest_sheet.delete_rows(latest_row, 1)

    has_remaining_data = False
    for row_index in range(2, latest_sheet.max_row + 1):
        values = [
            latest_sheet.cell(row=row_index, column=col).value
            for col in range(1, len(module.headers) + 1)
        ]
        if not is_row_empty(values):
            has_remaining_data = True
            break

    if not has_remaining_data:
        wb.remove(latest_sheet)

    rebuild_master_journal_from_daily_pages(wb, module)
    reorder_journal_sheets(wb)
    save_workbook_with_retry(wb, workbook_path)
    return True


def get_latest_journal_entry_for_edit() -> Optional[Dict[str, object]]:
    module = MODULES["J"]
    workbook_path = ensure_workbook(module)
    wb = load_workbook_with_retry(workbook_path)

    latest_sheet = None
    latest_date = None
    for sheet in wb.worksheets:
        if sheet.title == MASTER_JOURNAL_SHEET:
            continue
        try:
            sheet_date = datetime.strptime(sheet.title, "%Y-%m-%d")
        except ValueError:
            continue
        if latest_date is None or sheet_date > latest_date:
            latest_date = sheet_date
            latest_sheet = sheet

    if latest_sheet is None:
        return None

    latest_row = None
    latest_values: Optional[List[object]] = None
    for row_index in range(latest_sheet.max_row, 1, -1):
        values = [
            latest_sheet.cell(row=row_index, column=col).value
            for col in range(1, len(module.headers) + 1)
        ]
        if not is_row_empty(values):
            latest_row = row_index
            latest_values = values
            break

    if latest_row is None or latest_values is None:
        return None

    date_value = "" if latest_values[0] is None else str(latest_values[0])
    time_value = "" if latest_values[1] is None else str(latest_values[1])
    journal_value = "" if latest_values[2] is None else str(latest_values[2])
    speech_value = ""
    report_value = ""
    if len(latest_values) > 3 and latest_values[3] is not None:
        speech_value = str(latest_values[3])
    if len(latest_values) > 4 and latest_values[4] is not None:
        report_value = str(latest_values[4])
    return {
        "sheet_name": latest_sheet.title,
        "row_index": latest_row,
        "date": date_value,
        "time": time_value,
        "text": journal_value,
        "speech_transcript": speech_value,
        "ai_report": report_value,
        "images": [],
    }


def get_latest_journal_entry_for_delete() -> Optional[Dict[str, object]]:
    return get_latest_journal_entry_for_edit()


def update_journal_entry_at(sheet_name: str, row_index: int, row_values: List[str]) -> bool:
    module = MODULES["J"]
    workbook_path = ensure_workbook(module)
    wb = load_workbook_with_retry(workbook_path)
    if sheet_name not in wb.sheetnames:
        return False
    ws = wb[sheet_name]
    if row_index < 2:
        return False

    row_list = list(row_values)
    while len(row_list) < len(module.headers):
        row_list.append("")
    row_values = row_list[: len(module.headers)]

    for col_index, value in enumerate(row_values, start=1):
        ws.cell(row=row_index, column=col_index, value=value)

    rebuild_master_journal_from_daily_pages(wb, module)
    reorder_journal_sheets(wb)
    save_workbook_with_retry(wb, workbook_path)
    return True


def delete_journal_entry_at(sheet_name: str, row_index: int) -> bool:
    module = MODULES["J"]
    workbook_path = ensure_workbook(module)
    wb = load_workbook_with_retry(workbook_path)
    if sheet_name not in wb.sheetnames:
        return False
    ws = wb[sheet_name]
    if row_index < 2:
        return False
    ws.delete_rows(row_index, 1)

    has_remaining_data = False
    for candidate_row in range(2, ws.max_row + 1):
        values = [
            ws.cell(row=candidate_row, column=col).value
            for col in range(1, len(module.headers) + 1)
        ]
        if not is_row_empty(values):
            has_remaining_data = True
            break
    if not has_remaining_data:
        wb.remove(ws)

    rebuild_master_journal_from_daily_pages(wb, module)
    reorder_journal_sheets(wb)
    save_workbook_with_retry(wb, workbook_path)
    return True


def load_all_journal_entries() -> List[Tuple[datetime, str, str]]:
    module = MODULES["J"]
    workbook_path = ensure_workbook(module)
    wb = load_workbook_with_retry(workbook_path)
    entries: List[Tuple[datetime, str, str]] = []

    for sheet in wb.worksheets:
        if sheet.title == MASTER_JOURNAL_SHEET:
            continue
        try:
            sheet_date = datetime.strptime(sheet.title, "%Y-%m-%d")
        except ValueError:
            continue

        for row_index in range(2, sheet.max_row + 1):
            values = [
                sheet.cell(row=row_index, column=col).value
                for col in range(1, len(module.headers) + 1)
            ]
            if is_row_empty(values):
                continue
            date_value = "" if values[0] is None else str(values[0])
            time_value = "" if values[1] is None else str(values[1])
            journal_value = "" if values[2] is None else str(values[2])
            if journal_value.strip():
                entries.append((sheet_date, f"{date_value} {time_value}".strip(), journal_value))

    entries.sort(key=lambda item: item[0])
    return entries


def _journal_entry_calendar_day(entry: Tuple[datetime, str, str]) -> date:
    """Day used for recap filters: prefer the entry row date column, else the sheet tab date."""
    sheet_dt, when_value, _text = entry
    default_year = sheet_dt.year
    raw = (when_value or "").strip()
    if raw:
        first = raw.split(None, 1)[0]
        parsed = parse_flexible_date(first, default_year)
        if parsed is not None:
            return parsed.date()
    return sheet_dt.date()


def build_journal_context() -> str:
    return build_journal_context_for_range(None)


def build_journal_context_for_range(
    date_range: Optional[Tuple[datetime, datetime]]
) -> str:
    entries = load_all_journal_entries()
    if not entries:
        return "No journal entries available."
    if date_range is not None:
        start_date, end_date = date_range
        lo = start_date.date()
        hi = end_date.date()
        entries = [
            item
            for item in entries
            if lo <= _journal_entry_calendar_day(item) <= hi
        ]
        if not entries:
            return "No journal entries available in the selected date range."
    lines = []
    for _, when_value, text in entries:
        lines.append(f"- [{when_value}] {text}")
    return "\n".join(lines)


def build_journal_context_for_date_set(dates: Any) -> str:
    """Build journal context for sheet-days that match any calendar date in ``dates``."""
    if not dates:
        return "No dates selected."
    day_set: set[date] = set()
    for d in dates:
        if isinstance(d, datetime):
            day_set.add(d.date())
        elif isinstance(d, date):
            day_set.add(d)
    if not day_set:
        return "No dates selected."
    entries = load_all_journal_entries()
    if not entries:
        return "No journal entries available."
    filtered = [item for item in entries if _journal_entry_calendar_day(item) in day_set]
    if not filtered:
        return "No journal entries available for the selected day(s)."
    lines = [f"- [{when_value}] {text}" for _, when_value, text in filtered]
    return "\n".join(lines)


def parse_recap_date_range(raw_range: str, default_year: int) -> Optional[Tuple[datetime, datetime]]:
    cleaned = " ".join(raw_range.strip().split())
    if not cleaned:
        return None
    normalized = cleaned.replace(".", "/")
    tokens: List[str]
    if "-" in normalized:
        tokens = [part.strip() for part in normalized.split("-", 1)]
    else:
        parts = normalized.split()
        if len(parts) == 1:
            single = parse_flexible_date(parts[0], default_year)
            if single is None:
                return None
            return single, single
        if len(parts) != 2:
            return None
        tokens = parts
    if len(tokens) == 1:
        single = parse_flexible_date(tokens[0], default_year)
        if single is None:
            return None
        return single, single
    if len(tokens) != 2 or not tokens[0] or not tokens[1]:
        return None
    start = parse_flexible_date(tokens[0], default_year)
    end = parse_flexible_date(tokens[1], default_year)
    if start is None or end is None:
        return None
    if end < start:
        start, end = end, start
    return start, end


def list_journal_dates_in_range(date_range: Tuple[datetime, datetime]) -> List[str]:
    start_date, end_date = date_range
    lo = start_date.date()
    hi = end_date.date()
    matched_dates: set[str] = set()
    for entry in load_all_journal_entries():
        d = _journal_entry_calendar_day(entry)
        if lo <= d <= hi:
            matched_dates.add(d.strftime("%m/%d/%Y"))
    return sorted(
        matched_dates,
        key=lambda value: datetime.strptime(value, "%m/%d/%Y"),
    )


def load_recap_context_from_file(raw_path: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """Return (context_text, resolved_path, error_message)."""
    if not raw_path.strip():
        return None, None, "Missing file path."
    token = raw_path.strip().strip('"').strip("'")
    candidates = [
        Path(token),
        BASE_DIR / token,
        DATA_DIR / token,
    ]
    file_path: Optional[Path] = None
    for cand in candidates:
        if cand.exists() and cand.is_file():
            file_path = cand
            break
    if file_path is None:
        return None, None, f"File not found: {token}"
    suffix = file_path.suffix.lower()
    if suffix in (".xlsx", ".xlsm", ".xls"):
        return None, None, "Recap file lookup only supports text-like files (not Excel)."
    try:
        text = file_path.read_text(encoding="utf-8", errors="replace")
    except OSError as exc:
        return None, None, f"Could not read file: {exc}"
    if not text.strip():
        return None, None, "Selected file is empty."
    clipped = text
    if len(clipped) > 120000:
        clipped = clipped[:120000] + "\n\n[Truncated for recap]"
    header = f"Recap source file: {file_path.resolve()}\n"
    return header + clipped, str(file_path.resolve()), None


def resolve_recap_target(
    raw_arg: str, default_year: int
) -> Tuple[Optional[Tuple[datetime, datetime]], Optional[str], Optional[str], Optional[str]]:
    """Return (date_range, file_context, file_path, error)."""
    arg = raw_arg.strip()
    if not arg:
        return None, None, None, "Missing recap argument."
    recap_range = parse_recap_date_range(arg, default_year)
    if recap_range is not None:
        return recap_range, None, None, None
    file_context, file_path, file_error = load_recap_context_from_file(arg)
    if file_error is None:
        return None, file_context, file_path, None
    return None, None, None, (
        "Invalid recap target. Use a date range (e.g. 4/27 - 4/30) "
        "or a file path (e.g. notes.txt)."
    )


def get_openai_api_key() -> Optional[str]:
    env_key = os.getenv("OPENAI_API_KEY", "").strip()
    if env_key:
        return env_key
    if API_KEY_FILE.exists():
        try:
            file_key = API_KEY_FILE.read_text(encoding="utf-8").strip()
            if file_key:
                return file_key
        except OSError:
            return None
    return None


def save_openai_api_key(api_key: str) -> bool:
    try:
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        API_KEY_FILE.write_text(api_key.strip(), encoding="utf-8")
        return True
    except OSError:
        return False


def delete_openai_api_key() -> bool:
    if not API_KEY_FILE.exists():
        return True
    try:
        API_KEY_FILE.unlink()
        return True
    except OSError:
        return False


def copy_text_to_clipboard(text: str) -> bool:
    # Prefer native Windows clipboard command for CLI reliability.
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", "Set-Clipboard -Value @'\n" + text + "\n'@"],
            capture_output=True,
            text=True,
            check=False,
        )
        return result.returncode == 0
    except OSError:
        pass
    if tk is not None:
        try:
            root = tk.Tk()
            root.withdraw()
            root.clipboard_clear()
            root.clipboard_append(text)
            root.update()
            root.destroy()
            return True
        except Exception:
            return False
    return False


def ensure_openai_api_key_for_ai() -> bool:
    existing = get_openai_api_key()
    if existing:
        return True

    print("AI feature needs an OpenAI API key.")
    pasted = input("Paste your OpenAI API key (or press Enter to cancel): ").strip()
    if is_enter_equivalent(pasted):
        print("No API key entered. Returning to main menu.")
        return False
    if not save_openai_api_key(pasted):
        print("Could not save API key file. Check folder permissions and try again.")
        return False
    print("API key saved for future use.")
    return True


def load_preferences() -> Dict[str, str]:
    if not PREFS_FILE.exists():
        return {}
    try:
        raw = json.loads(PREFS_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(raw, dict):
        return {}
    result: Dict[str, str] = {}
    for key, value in raw.items():
        if isinstance(key, str) and isinstance(value, str):
            result[key] = value
    return result


def save_preferences(prefs: Dict[str, str]) -> bool:
    try:
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        PREFS_FILE.write_text(json.dumps(prefs, indent=2), encoding="utf-8")
        return True
    except OSError:
        return False


def _is_pref_true(value: str) -> bool:
    return value.strip().lower() == "true"


def ensure_backup_folder() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)


def _list_backup_zip_files() -> List[Path]:
    ensure_backup_folder()
    return sorted(
        [path for path in BACKUP_DIR.glob("*.zip") if path.is_file()],
        key=lambda item: item.stat().st_mtime,
        reverse=True,
    )


def run_backup_now() -> Optional[Path]:
    ensure_backup_folder()
    items_to_backup = [
        path
        for path in DATA_DIR.iterdir()
        if path.name.lower() != BACKUP_DIR.name.lower()
    ]
    if not items_to_backup:
        return None

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    zip_path = BACKUP_DIR / f"backup_{timestamp}.zip"

    with zipfile.ZipFile(zip_path, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        for item in items_to_backup:
            if item.is_file():
                archive.write(item, arcname=item.name)
                continue
            if item.is_dir():
                for nested in item.rglob("*"):
                    if nested.is_file():
                        archive.write(nested, arcname=str(nested.relative_to(DATA_DIR)))
    return zip_path


def trim_backups_if_limited(prefs: Dict[str, str]) -> None:
    if not _is_pref_true(prefs.get("backup_limited", "false")):
        return
    backups = _list_backup_zip_files()
    if len(backups) <= 3:
        return
    oldest_backup = backups[-1]
    try:
        oldest_backup.unlink()
        print(f"Backup limited mode: removed oldest backup {oldest_backup.name}")
    except OSError:
        print(f"Backup limited mode: could not remove {oldest_backup.name}")


def evict_oldest_backup_if_limited_full(prefs: Dict[str, str]) -> None:
    if not _is_pref_true(prefs.get("backup_limited", "false")):
        return
    backups = _list_backup_zip_files()
    if len(backups) < 3:
        return
    oldest_backup = backups[-1]
    try:
        oldest_backup.unlink()
        print(f"Backup limited mode: removed oldest backup {oldest_backup.name} before new backup")
    except OSError:
        print(f"Backup limited mode: could not remove {oldest_backup.name} before new backup")


def maybe_run_daily_auto_backup() -> None:
    prefs = load_preferences()
    backup_enabled = prefs.get("backup_enabled", "true")
    if not _is_pref_true(backup_enabled):
        return

    ensure_backup_folder()
    today = datetime.now().strftime("%Y-%m-%d")
    last_program_run_date = prefs.get("last_program_run_date", "").strip()
    if last_program_run_date != today:
        evict_oldest_backup_if_limited_full(prefs)
        backup_path = run_backup_now()
        if backup_path is None:
            print("Auto backup skipped: nothing in daily_logs to back up.")
        else:
            print(f"Auto backup created: {backup_path.name}")
            trim_backups_if_limited(prefs)
            prefs["last_backup_date"] = today

    prefs["backup_enabled"] = "true" if _is_pref_true(backup_enabled) else "false"
    prefs["last_program_run_date"] = today
    if not save_preferences(prefs):
        print("Warning: could not save backup preferences.")


def prompt_for_app_name() -> str:
    entered = input(
        "What would you like to name this app? (Press Enter for default name): "
    ).strip()
    if is_enter_equivalent(entered):
        return "Daily Logger"
    return entered


def get_or_create_app_name() -> str:
    prefs = load_preferences()
    app_name = prefs.get("app_name", "").strip()
    if app_name:
        return app_name
    app_name = prompt_for_app_name()
    prefs["app_name"] = app_name
    if not save_preferences(prefs):
        print("Warning: could not save app name preference.")
    return app_name


def rename_app_name() -> str:
    app_name = prompt_for_app_name()
    prefs = load_preferences()
    prefs["app_name"] = app_name
    if save_preferences(prefs):
        print(f'App renamed to "{app_name}".')
    else:
        print("Could not save new app name preference.")
    return app_name


def rename_app_name_to(new_name: str) -> str:
    app_name = new_name.strip() or "Daily Logger"
    prefs = load_preferences()
    prefs["app_name"] = app_name
    if save_preferences(prefs):
        print(f'App renamed to "{app_name}".')
    else:
        print("Could not save new app name preference.")
    return app_name


def get_startup_folder() -> Optional[Path]:
    appdata = os.getenv("APPDATA", "").strip()
    if not appdata:
        return None
    return Path(appdata) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"


def get_startup_shortcut_path() -> Optional[Path]:
    startup_dir = get_startup_folder()
    if startup_dir is None:
        return None
    return startup_dir / STARTUP_SHORTCUT_NAME


def create_startup_shortcut() -> bool:
    shortcut_path = get_startup_shortcut_path()
    if shortcut_path is None:
        return False
    try:
        shortcut_path.parent.mkdir(parents=True, exist_ok=True)
    except OSError:
        return False
    target_path = str((BASE_DIR / "launch_daily_logger.bat").resolve())
    ps_script = (
        "$WshShell = New-Object -ComObject WScript.Shell; "
        f"$Shortcut = $WshShell.CreateShortcut('{str(shortcut_path)}'); "
        f"$Shortcut.TargetPath = '{target_path}'; "
        f"$Shortcut.WorkingDirectory = '{str(BASE_DIR)}'; "
        "$Shortcut.Save();"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True,
            text=True,
            check=False,
        )
        return result.returncode == 0 and shortcut_path.exists()
    except OSError:
        return False


def remove_startup_shortcut() -> bool:
    shortcut_path = get_startup_shortcut_path()
    if shortcut_path is None:
        return False
    if not shortcut_path.exists():
        return True
    try:
        shortcut_path.unlink()
        return True
    except OSError:
        return False


def is_startup_enabled() -> bool:
    shortcut_path = get_startup_shortcut_path()
    return bool(shortcut_path and shortcut_path.exists())


def open_current_directory_in_explorer() -> bool:
    return open_path_with_default_app(USER_DATA_ROOT)


def open_path_with_default_app(path: Path) -> bool:
    try:
        if sys.platform == "win32":
            os.startfile(str(path))  # type: ignore[attr-defined]
            return True
        return False
    except OSError:
        return False


def _remove_path_quietly(path: Path) -> None:
    try:
        if path.is_dir():
            shutil.rmtree(path, ignore_errors=True)
        elif path.exists():
            path.unlink()
    except OSError:
        pass


def _schedule_windows_self_delete(exe_path: Path) -> None:
    script_path = Path(tempfile.gettempdir()) / f"daily_logger_uninstall_{int(time.time())}.cmd"
    script = (
        "@echo off\n"
        "timeout /t 2 /nobreak >nul\n"
        f'del /f /q "{exe_path}" >nul 2>&1\n'
        f'del /f /q "{script_path}" >nul 2>&1\n'
    )
    try:
        script_path.write_text(script, encoding="utf-8")
        subprocess.Popen(
            ["cmd", "/c", str(script_path)],
            creationflags=getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
        )
    except OSError:
        pass


def run_clean_uninstall() -> None:
    remove_startup_shortcut()
    _remove_path_quietly(DATA_DIR)
    _remove_path_quietly(SETTINGS_DIR)

    if getattr(sys, "frozen", False):
        exe_path = Path(sys.executable).resolve()
        _schedule_windows_self_delete(exe_path)
        print("Uninstall started. App data folders and this EXE will be removed after this window closes.")
        return

    # Dev-mode fallback: do not delete source code automatically.
    print("Uninstall cleaned app data folders (daily_logs/settings).")


def get_start_menu_programs_dir() -> Optional[Path]:
    appdata = os.getenv("APPDATA", "").strip()
    if not appdata:
        return None
    return Path(appdata) / "Microsoft" / "Windows" / "Start Menu" / "Programs"


def create_start_menu_search_shortcut(
    shortcut_path: Path,
    target_path: Path,
    working_directory: Path,
    description: str,
) -> bool:
    """Create a .lnk in Start Menu Programs so Windows Search can surface the target."""
    try:
        shortcut_path.parent.mkdir(parents=True, exist_ok=True)
    except OSError:
        return False
    target = str(target_path.resolve()).replace("'", "''")
    work_dir = str(working_directory.resolve()).replace("'", "''")
    desc = description.replace("'", "''")
    lnk = str(shortcut_path).replace("'", "''")
    ps_script = (
        "$WshShell = New-Object -ComObject WScript.Shell; "
        f"$Shortcut = $WshShell.CreateShortcut('{lnk}'); "
        f"$Shortcut.TargetPath = '{target}'; "
        f"$Shortcut.WorkingDirectory = '{work_dir}'; "
        f"$Shortcut.Description = '{desc}'; "
        "$Shortcut.Save();"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True,
            text=True,
            check=False,
        )
        return result.returncode == 0 and shortcut_path.exists()
    except OSError:
        return False


def sb_create_bat_search_shortcut() -> bool:
    programs = get_start_menu_programs_dir()
    if programs is None:
        return False
    folder = programs / "Daily Logger"
    shortcut_path = folder / "Daily Logger BAT Launcher.lnk"
    bat_path = BASE_DIR / "launch_daily_logger.bat"
    if not bat_path.exists():
        print(f"Missing launcher file: {bat_path}")
        return False
    ok = create_start_menu_search_shortcut(
        shortcut_path,
        bat_path,
        BASE_DIR,
        "Daily Logger batch launcher - search: Daily Logger, BAT, batch, logger",
    )
    if ok:
        print(f"Search shortcut created: {shortcut_path}")
        print("Try Windows search for: Daily Logger, BAT, or batch.")
    return ok


def sb_create_journal_search_shortcut() -> bool:
    programs = get_start_menu_programs_dir()
    if programs is None:
        return False
    journal_path = ensure_workbook(MODULES["J"])
    folder = programs / "Daily Logger"
    shortcut_path = folder / "Daily Logger Journal Excel.lnk"
    ok = create_start_menu_search_shortcut(
        shortcut_path,
        journal_path,
        journal_path.parent,
        "Daily Logger journal workbook - search: Journal, Excel, Daily Logger, xlsx",
    )
    if ok:
        print(f"Search shortcut created: {shortcut_path}")
        print("Try Windows search for: Daily Logger Journal, Excel, or Journal.")
    return ok


def load_wifi_warn_list() -> List[str]:
    if not WIFI_WARN_FILE.exists():
        return []
    try:
        parsed = json.loads(WIFI_WARN_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    if not isinstance(parsed, list):
        return []
    result: List[str] = []
    seen = set()
    for item in parsed:
        if not isinstance(item, str):
            continue
        cleaned = item.strip()
        key = cleaned.lower()
        if not cleaned or key in seen:
            continue
        seen.add(key)
        result.append(cleaned)
    return result


def save_wifi_warn_list(names: List[str]) -> bool:
    try:
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        WIFI_WARN_FILE.write_text(json.dumps(names, indent=2), encoding="utf-8")
        return True
    except OSError:
        return False


def add_wifi_warn_name(name: str) -> bool:
    cleaned = name.strip()
    if not cleaned:
        return False
    existing = load_wifi_warn_list()
    existing_lower = {item.lower() for item in existing}
    if cleaned.lower() in existing_lower:
        return True
    existing.append(cleaned)
    return save_wifi_warn_list(existing)


def get_current_wifi_name() -> Optional[str]:
    try:
        result = subprocess.run(
            ["netsh", "wlan", "show", "interfaces"],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError:
        return None
    if result.returncode != 0:
        return None
    for line in result.stdout.splitlines():
        stripped = line.strip()
        if not stripped.startswith("SSID"):
            continue
        if stripped.startswith("SSID BSSID"):
            continue
        if ":" not in stripped:
            continue
        value = stripped.split(":", 1)[1].strip()
        if value:
            return value
    return None


def maybe_warn_for_current_wifi() -> None:
    warned_names = load_wifi_warn_list()
    if not warned_names:
        return
    current_wifi = get_current_wifi_name()
    if not current_wifi:
        return
    warned_lower = {name.lower() for name in warned_names}
    if current_wifi.lower() in warned_lower:
        print(red_text(f'Warning: you are on "{current_wifi}" connection, it might not work.'))


def load_journal_window_draft() -> Optional[Dict[str, object]]:
    if not JOURNAL_WINDOW_DRAFT_FILE.exists():
        return None
    try:
        parsed = json.loads(JOURNAL_WINDOW_DRAFT_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(parsed, dict):
        return None
    return parsed


def save_journal_window_draft(draft: Dict[str, object]) -> bool:
    try:
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        JOURNAL_WINDOW_DRAFT_FILE.write_text(json.dumps(draft, indent=2), encoding="utf-8")
        return True
    except OSError:
        return False


def clear_journal_window_draft() -> None:
    try:
        if JOURNAL_WINDOW_DRAFT_FILE.exists():
            JOURNAL_WINDOW_DRAFT_FILE.unlink()
    except OSError:
        pass


def normalize_window_time_input(raw: str) -> Optional[str]:
    cleaned = raw.strip()
    if not cleaned:
        return datetime.now().strftime("%I:%M%p").lstrip("0")
    if cleaned.lower() in ("n/a", "na"):
        return "N/A"
    if cleaned.lower() == "rn":
        return datetime.now().strftime("%I:%M%p").lstrip("0")
    normalized = cleaned.upper().replace(" ", "")
    try:
        parsed = datetime.strptime(normalized, "%I:%M%p")
        return parsed.strftime("%I:%M%p").lstrip("0")
    except ValueError:
        return None


def _read_wav_mono_int16(path: Path) -> Tuple[Optional[Any], int, Optional[str]]:
    """Load 16-bit PCM WAV as mono int16 ndarray. Returns (samples, sample_rate, error)."""
    try:
        import numpy as np
    except Exception as exc:
        return None, 0, str(exc)
    try:
        with wave.open(str(path), "rb") as wf:
            ch = wf.getnchannels()
            sw = wf.getsampwidth()
            rate = wf.getframerate() or 16000
            nframes = wf.getnframes()
            raw = wf.readframes(nframes)
    except Exception as exc:
        return None, 0, str(exc)
    if sw != 2:
        return None, 0, "Whisper preprocessing needs 16-bit PCM WAV."
    data = np.frombuffer(raw, dtype=np.int16)
    if ch == 1:
        mono = data
    elif ch >= 2:
        flat = data.reshape(-1, ch).astype(np.float32)
        mono = np.mean(flat, axis=1).astype(np.int16)
    else:
        return None, 0, "Invalid WAV channel count."
    return mono, int(rate), None


def _rms_per_frame_int16(samples: Any, frame: int) -> Any:
    import numpy as np

    n = (int(samples.shape[0]) // frame) * frame
    if n <= 0:
        return np.array([], dtype=np.float64)
    blocks = samples[:n].reshape(-1, frame).astype(np.float64)
    return np.sqrt(np.mean(blocks * blocks, axis=1))


def preprocess_wav_for_whisper(samples: Any, sample_rate: int) -> Tuple[Any, Optional[str]]:
    """Trim edge silence and shorten long internal silences. Returns (mono int16 ndarray, error)."""
    try:
        import numpy as np
    except Exception as exc:
        return None, str(exc)
    if samples is None or int(samples.shape[0]) < 1:
        return None, "Empty audio."
    rate = max(1, int(sample_rate))
    frame = max(int(rate * (WHISPER_PRE_FRAME_MS / 1000.0)), 1)
    thr = float(WHISPER_PRE_SILENCE_RMS)
    rms = _rms_per_frame_int16(samples, frame)
    if rms.size < 1:
        return samples, None
    voiced = rms > thr
    if not bool(np.any(voiced)):
        return samples[:0], "No speech detected (audio is mostly silence)."
    first = int(np.argmax(voiced))
    last = int(rms.shape[0] - 1 - np.argmax(voiced[::-1]))
    pad = int(rate * (WHISPER_PRE_EDGE_PAD_MS / 1000.0))
    min_sp = int(rate * (WHISPER_PRE_MIN_SPEECH_MS / 1000.0))
    start = max(0, first * frame - pad)
    end = min(int(samples.shape[0]), (last + 1) * frame + pad)
    if end - start < min_sp:
        start = max(0, first * frame)
        end = min(int(samples.shape[0]), (last + 1) * frame)
    trimmed = samples[start:end].copy()
    if int(trimmed.shape[0]) < min_sp:
        return trimmed[:0], "No speech detected (audio is mostly silence)."

    rms2 = _rms_per_frame_int16(trimmed, frame)
    if rms2.size < 1:
        return trimmed, None
    v2 = rms2 > thr
    max_gap = int((rate * WHISPER_PRE_MAX_INTERNAL_SILENCE_SEC) / frame)
    keep_gap = max(1, int((rate * WHISPER_PRE_KEEP_INTERNAL_SILENCE_SEC) / frame))
    keep_samples = keep_gap * frame

    out_parts: List[Any] = []
    f = 0
    nfr = int(v2.shape[0])
    while f < nfr:
        while f < nfr and not bool(v2[f]):
            f += 1
        if f >= nfr:
            break
        t = f
        while t < nfr and bool(v2[t]):
            t += 1
        out_parts.append(trimmed[f * frame : t * frame])
        if t >= nfr:
            break
        u = t
        while u < nfr and not bool(v2[u]):
            u += 1
        silence_frames = u - t
        if silence_frames > max_gap:
            out_parts.append(np.zeros(keep_samples, dtype=np.int16))
        else:
            out_parts.append(trimmed[t * frame : u * frame])
        f = u

    if not out_parts:
        return trimmed, None
    merged = np.concatenate(out_parts, axis=0)
    n_sample_full = (int(trimmed.shape[0]) // frame) * frame
    if n_sample_full < int(trimmed.shape[0]):
        merged = np.concatenate([merged, trimmed[n_sample_full:]], axis=0)
    if int(merged.shape[0]) < 1:
        return merged, "No speech detected after preprocessing."
    return merged, None


def prepare_wav_path_for_whisper(source: Path) -> Tuple[Path, Optional[str], Optional[Path]]:
    """Pick WAV bytes to upload: trimmed/collapsed copy when possible.

    Returns (upload_path, fatal_error_string_or_none, temp_path_to_delete_or_none).
    """
    mono, rate, err = _read_wav_mono_int16(source)
    if err is not None or mono is None:
        return source, None, None
    processed, perr = preprocess_wav_for_whisper(mono, rate)
    if perr is not None:
        return source, perr, None
    try:
        import numpy as np
    except Exception:
        return source, None, None
    if processed is None or int(processed.shape[0]) < 1:
        return source, "No speech detected (audio is mostly silence).", None
    if int(processed.shape[0]) == int(mono.shape[0]) and bool(np.array_equal(processed, mono)):
        return source, None, None
    fd, tmp_name = tempfile.mkstemp(suffix=".wav", prefix="whisper_pre_")
    os.close(fd)
    tmp = Path(tmp_name)
    werr = write_mono_int16_wav(tmp, processed, rate)
    if werr is not None:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        return source, None, None
    return tmp, None, tmp




def _transcribe_audio_openai_single(
    file_path: Path,
    language: Optional[str],
    *,
    prompt: Optional[str] = None,
    temperature: float = 0.0,
    progress: Optional[Callable[[int], None]] = None,
) -> str:
    """Single-request Whisper upload. Caller handles retries/fallback strategy."""

    def _pg(p: int) -> None:
        if progress is not None:
            try:
                progress(min(100, max(0, int(p))))
            except Exception:
                pass

    _pg(10)
    boundary = uuid.uuid4().hex.encode("ascii")
    crlf = b"\r\n"
    body_chunks: List[bytes] = []

    def add_field(name: str, value: str) -> None:
        body_chunks.append(b"--" + boundary + crlf)
        body_chunks.append(
            f'Content-Disposition: form-data; name="{name}"'.encode("utf-8") + crlf + crlf
        )
        body_chunks.append(value.encode("utf-8") + crlf)

    api_key = get_openai_api_key()
    if not api_key:
        return "OPENAI_API_KEY is not set. Use TOKEN ADD in the main menu or set the environment variable."

    _pg(14)
    add_field("model", "whisper-1")
    if language:
        add_field("language", language)
    if prompt and prompt.strip():
        add_field("prompt", prompt.strip()[:WHISPER_TRANSCRIBE_PROMPT_CHAR_LIMIT])
    add_field("temperature", str(temperature))

    _pg(22)
    filename = file_path.name
    try:
        audio_bytes = file_path.read_bytes()
    except OSError as exc:
        return f"Could not read audio file: {exc}"

    _pg(30)
    body_chunks.append(b"--" + boundary + crlf)
    body_chunks.append(
        f'Content-Disposition: form-data; name="file"; filename="{filename}"'.encode("utf-8")
        + crlf
    )
    body_chunks.append(b"Content-Type: audio/wav" + crlf + crlf)
    body_chunks.append(audio_bytes + crlf)
    body_chunks.append(b"--" + boundary + b"--" + crlf)
    body = b"".join(body_chunks)

    _pg(38)
    req = request.Request(
        OPENAI_TRANSCRIPTION_URL,
        data=body,
        method="POST",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": f"multipart/form-data; boundary={boundary.decode('ascii')}",
        },
    )

    try:
        with request.urlopen(req, timeout=120) as response:
            raw_bytes = response.read()
        _pg(76)
        raw = raw_bytes.decode("utf-8")
    except error.HTTPError as exc:
        _pg(50)
        try:
            details = exc.read().decode("utf-8")
        except Exception:
            details = str(exc)
        result = f"Whisper API error ({exc.code}): {details}"
    except Exception as exc:
        _pg(50)
        result = f"Whisper request failed: {exc}"
    else:
        try:
            parsed = json.loads(raw)
            text = parsed.get("text")
            if isinstance(text, str):
                result = text.strip()
                _pg(94)
            else:
                result = "Whisper returned an unexpected response format."
                _pg(88)
        except json.JSONDecodeError:
            result = "Whisper returned invalid JSON."
            _pg(88)

    return result


def _whisper_context_too_long_error(text: str) -> bool:
    needle = text.strip().lower()
    if not needle:
        return False
    markers = (
        "maximum context length",
        "context length exceeded",
        "prompt is too long",
        "too many tokens",
        "reduce the length",
        "request too large",
        "payload too large",
        "content size limit",
        "maximum content size",
        "26214400",
        "413",
        "entity too large",
    )
    return any(m in needle for m in markers)


def _is_likely_api_error_message_global(text: str) -> bool:
    """Module-level variant used by non-UI helpers (UI also defines its own)."""
    t = (text or "").strip()
    if not t:
        return False
    prefixes = (
        "OPENAI_API_KEY",
        "ChatGPT API error",
        "Failed to contact ChatGPT",
        "ChatGPT returned",
        "No response received",
        "Whisper API error",
        "Whisper request failed",
        "Whisper returned",
        "Could not read audio file",
        "Recording needs optional packages",
        "No speech detected",
        "Empty audio.",
    )
    return any(t.startswith(p) for p in prefixes)


def whisper_chunk_duration_sec(sample_rate: int) -> int:
    """Seconds of mono int16 audio per chunk so WAV uploads stay under Whisper size limits."""
    rate = max(1, int(sample_rate))
    pcm_bps = 2 * rate
    max_sec_budget = int(WHISPER_SAFE_CHUNK_PCM_BYTES // pcm_bps)
    return max(45, min(WHISPER_TRANSCRIBE_CHUNK_SEC, max_sec_budget))


def _transcribe_audio_openai_chunked(
    file_path: Path,
    language: Optional[str],
    *,
    prompt: Optional[str] = None,
    temperature: float = 0.0,
    progress: Optional[Callable[[int], None]] = None,
) -> str:
    """Fallback for oversized uploads/context: split WAV and merge partial transcripts."""

    def _pg(p: int) -> None:
        if progress is not None:
            try:
                progress(min(100, max(0, int(p))))
            except Exception:
                pass

    _pg(12)
    mono, rate, read_err = _read_wav_mono_int16(file_path)
    if read_err is not None or mono is None:
        return _transcribe_audio_openai_single(
            file_path,
            language,
            prompt=prompt,
            temperature=temperature,
            progress=progress,
        )
    chunk_sec = whisper_chunk_duration_sec(int(rate))
    chunk_samples = max(int(rate * chunk_sec), 1)
    if int(mono.shape[0]) <= chunk_samples:
        return _transcribe_audio_openai_single(
            file_path,
            language,
            prompt=prompt,
            temperature=temperature,
            progress=progress,
        )

    transcripts: List[str] = []
    sample_count = int(mono.shape[0])
    chunk_starts = list(range(0, sample_count, chunk_samples))
    n_chunks = max(len(chunk_starts), 1)
    _pg(18)
    for ci, start in enumerate(chunk_starts):
        end = min(start + chunk_samples, sample_count)
        part = mono[start:end]
        if int(part.shape[0]) < 1:
            continue
        fd, tmp_name = tempfile.mkstemp(suffix=".wav", prefix="whisper_chunk_")
        os.close(fd)
        tmp = Path(tmp_name)
        try:
            werr = write_mono_int16_wav(tmp, part, rate)
            if werr is not None:
                return f"Could not write chunked audio: {werr}"
            chunk_result = _transcribe_audio_openai_single(
                tmp,
                language,
                prompt=prompt,
                temperature=temperature,
                progress=None,
            ).strip()
        finally:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
        if _is_likely_api_error_message_global(chunk_result):
            return chunk_result
        if chunk_result:
            transcripts.append(chunk_result)
        _pg(22 + int(72 * (ci + 1) / n_chunks))
    merged = " ".join(t for t in transcripts if t.strip()).strip()
    _pg(97)
    if merged:
        return merged
    return "Whisper returned empty text."


def transcribe_audio_openai(
    file_path: Path,
    language: Optional[str],
    *,
    prompt: Optional[str] = None,
    temperature: float = 0.0,
    progress: Optional[Callable[[int], None]] = None,
) -> str:
    """Send local audio to Whisper with fallback for long context/uploads."""
    last_pct = [0]

    def _p(pct: int) -> None:
        v = min(100, max(0, int(pct)))
        if v < last_pct[0]:
            v = last_pct[0]
        else:
            last_pct[0] = v
        if progress is not None:
            try:
                progress(v)
            except Exception:
                pass

    _p(2)
    upload_path, prep_err, temp_upload = prepare_wav_path_for_whisper(file_path)
    if prep_err is not None:
        return prep_err
    _p(6)
    try:
        upl_sz = 0
        try:
            upl_sz = int(upload_path.stat().st_size)
        except OSError:
            pass
        if upl_sz >= WHISPER_SKIP_SINGLE_FILE_BYTES:
            _p(10)
            return _transcribe_audio_openai_chunked(
                upload_path,
                language,
                prompt=prompt,
                temperature=temperature,
                progress=_p,
            )
        first_try = _transcribe_audio_openai_single(
            upload_path,
            language,
            prompt=prompt,
            temperature=temperature,
            progress=_p,
        )
        if _whisper_context_too_long_error(first_try):
            _p(92)
            return _transcribe_audio_openai_chunked(
                upload_path,
                language,
                prompt=prompt,
                temperature=temperature,
                progress=_p,
            )
        _p(99)
        return first_try
    finally:
        if temp_upload is not None:
            try:
                temp_upload.unlink(missing_ok=True)
            except OSError:
                pass


def archive_journal_recording(wav_path: Path) -> Optional[Path]:
    """Copy a session WAV into RECORDING_DIR.

    Files are named rcdYYYYMMDD.wav, then rcdYYYYMMDD1.wav, rcdYYYYMMDD2.wav, … for the same day.
    """
    try:
        RECORDING_DIR.mkdir(parents=True, exist_ok=True)
        day = datetime.now().strftime("%Y%m%d")
        base = f"rcd{day}"
        for n in range(0, 10000):
            name = f"{base}.wav" if n == 0 else f"{base}{n}.wav"
            dest = RECORDING_DIR / name
            if dest.exists():
                continue
            shutil.copy2(wav_path, dest)
            return dest.resolve()
    except OSError:
        return None


def latest_archived_journal_wav() -> Optional[Path]:
    """Newest journal clip in ``RECORDING_DIR`` (``rcd*.wav`` by modification time), or ``None``."""
    try:
        if not RECORDING_DIR.is_dir():
            return None
        best_mtime: float = -1.0
        best_path: Optional[Path] = None
        for p in RECORDING_DIR.iterdir():
            if not p.is_file() or p.suffix.lower() != ".wav":
                continue
            if not p.stem.lower().startswith("rcd"):
                continue
            try:
                mtime = float(p.stat().st_mtime)
            except OSError:
                continue
            if mtime > best_mtime:
                best_mtime = mtime
                best_path = p.resolve()
        return best_path
    except OSError:
        return None


def wav_mono_duration_seconds(path: Path) -> float:
    """Return duration in seconds for a readable mono WAV, or 0.0 on error."""
    try:
        with wave.open(str(path), "rb") as wf:
            rate = wf.getframerate() or 16000
            return wf.getnframes() / float(rate)
    except Exception:
        return 0.0


def estimate_whisper_cost_usd(wav_path: Path) -> Tuple[float, float]:
    """Return (duration_sec, approximate_usd) using WHISPER_USD_PER_MIN."""
    dur = wav_mono_duration_seconds(wav_path)
    usd = (dur / 60.0) * WHISPER_USD_PER_MIN if dur > 0 else 0.0
    return dur, usd


def bind_hover_tooltip(widget: Any, text_callable: Callable[[], str]) -> None:
    """Show a tooltip only for this widget; place inside its toplevel, hugging edges when clipped."""
    if tk is None:
        return
    tip: Dict[str, Optional[Any]] = {"w": None}

    def hide(_evt: Optional[Any] = None) -> None:
        tw = tip["w"]
        if tw is not None:
            try:
                tw.destroy()
            except tk.TclError:
                pass
            tip["w"] = None

    def show(evt: Any) -> None:
        hide()
        msg = (text_callable() or "").strip()
        if not msg:
            return
        tw = tk.Toplevel(widget)
        tip["w"] = tw
        tw.wm_overrideredirect(True)
        tw.wm_attributes("-topmost", True)
        lbl = tk.Label(
            tw,
            text=msg,
            justify="left",
            background="#ffffe0",
            foreground="#000000",
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 9),
            wraplength=TOOLTIP_WRAP_PX,
        )
        lbl.pack(ipadx=4, ipady=2)
        m = 8
        try:
            top = widget.winfo_toplevel()
            top.update_idletasks()
            win_x = int(top.winfo_rootx())
            win_y = int(top.winfo_rooty())
            win_w = max(int(top.winfo_width()), 160)
            win_h = max(int(top.winfo_height()), 120)
        except tk.TclError:
            win_x, win_y = 0, 0
            win_w, win_h = 800, 600
        win_r = win_x + win_w
        win_b = win_y + win_h
        cx = int(evt.x_root)
        cy = int(evt.y_root)
        pref_x = cx + 12
        pref_y = cy + 12
        space_right = max(0, win_r - m - pref_x)
        space_left = max(0, pref_x - win_x - m)
        if space_right >= space_left:
            wrap = max(100, min(TOOLTIP_WRAP_PX_MAX, space_right - 8))
        else:
            wrap = max(100, min(TOOLTIP_WRAP_PX_MAX, space_left - 8))
        max_inner = max(100, win_w - 2 * m - 16)
        lbl.config(wraplength=min(wrap, max_inner, TOOLTIP_WRAP_PX_MAX))
        tw.update_idletasks()
        tip_w = int(tw.winfo_reqwidth())
        tip_h = int(tw.winfo_reqheight())
        if tip_w > win_w - 2 * m:
            lbl.config(wraplength=max_inner)
            tw.update_idletasks()
            tip_w = int(tw.winfo_reqwidth())
            tip_h = int(tw.winfo_reqheight())
        x = pref_x
        if x + tip_w > win_r - m:
            x = win_r - m - tip_w
        if x < win_x + m:
            x = win_x + m
        y = pref_y
        if y + tip_h > win_b - m:
            y = win_b - m - tip_h
        if y < win_y + m:
            y = win_y + m
        x = max(win_x + m, min(x, win_r - tip_w - m))
        y = max(win_y + m, min(y, win_b - tip_h - m))
        tw.wm_geometry(f"+{x}+{y}")

    widget.bind("<Enter>", show, add="+")
    widget.bind("<Leave>", hide, add="+")
    widget.bind("<ButtonPress>", hide, add="+")


def bind_button_hover_if_enabled(
    widget: Any,
    get_rest_style: Callable[[], Tuple[str, str, str, str, str]],
    hover_bg: Union[str, Callable[[], str]],
    hover_fg: Union[str, Callable[[], str]],
) -> None:
    """Apply hover colors on <Enter> only when state is normal; <Leave> restores idle look.

    get_rest_style returns (state, bg, fg, activebackground, activeforeground) for the
    non-hover appearance; state should match widget.cget('state') logic for that moment.
    hover_bg / hover_fg may be callables (e.g. lambda: theme.hover_primary) so themes can
    change without rebinding.
    """
    if tk is None:
        return

    def _hover_color(spec: Union[str, Callable[[], str]]) -> str:
        return spec() if callable(spec) else spec

    def on_leave(_evt: Optional[Any] = None) -> None:
        try:
            st, bg, fg, abg, afg = get_rest_style()
        except tk.TclError:
            return
        kw: Dict[str, Any] = {
            "bg": bg,
            "fg": fg,
            "activebackground": abg,
            "activeforeground": afg,
        }
        if str(st) == "disabled":
            kw["disabledforeground"] = fg
        try:
            widget.config(**kw)
        except tk.TclError:
            pass

    def on_enter(_evt: Optional[Any] = None) -> None:
        try:
            st, _b, _f, _ab, _af = get_rest_style()
        except tk.TclError:
            return
        if str(st) != "normal":
            return
        hb = _hover_color(hover_bg)
        hf = _hover_color(hover_fg)
        try:
            widget.config(
                bg=hb,
                fg=hf,
                activebackground=hb,
                activeforeground=hf,
            )
        except tk.TclError:
            pass

    widget.bind("<Enter>", on_enter, add="+")
    widget.bind("<Leave>", on_leave, add="+")


def write_mono_int16_wav(path: Path, samples: object, sample_rate: int) -> Optional[str]:
    """Write mono int16 PCM to WAV. Returns error string or None on success."""
    try:
        import numpy as np
    except Exception as exc:
        return str(exc)
    if not isinstance(samples, np.ndarray):
        return "Internal error: audio must be a numpy array."
    arr = np.atleast_1d(samples.squeeze())
    if arr.dtype != np.int16:
        arr = arr.astype(np.int16)
    if arr.size == 0:
        return "Empty audio buffer."
    try:
        with wave.open(str(path), "wb") as wf:
            wf.setnchannels(1)
            wf.setsampwidth(2)
            wf.setframerate(sample_rate)
            wf.writeframes(arr.tobytes())
    except OSError as exc:
        return str(exc)
    return None


def record_microphone_session_wav(
    output_path: Path,
    stop_event: threading.Event,
    *,
    sample_rate: int = 16000,
    chunk_interval_sec: float = LIVE_STT_CHUNK_INTERVAL_SEC,
    on_audio_chunk: Optional[Callable[[Path], None]] = None,
    on_pcm_block: Optional[Callable[[Any], None]] = None,
    pause_event: Optional[threading.Event] = None,
) -> Optional[str]:
    """Record mono WAV until stop_event.

    Optional on_audio_chunk(path): periodic temp WAV paths for live STT (legacy).
    Optional on_pcm_block(block): each captured block as int16 numpy array (mono); runs in the record thread.
    Optional pause_event: while set, input is still read (to avoid device overrun) but not written to the
    output buffer and on_pcm_block is not called so metering/waveform can stay frozen until resumed.
    """
    try:
        import numpy as np
        import sounddevice as sd
    except Exception as exc:
        return (
            "Recording needs optional packages. Install with:\n"
            f"  {sys.executable} -m pip install sounddevice numpy\n"
            f"Details: {exc}"
        )

    frames: List[object] = []
    last_flushed_samples = 0
    next_chunk_at = time.monotonic() + chunk_interval_sec
    block_samples = (
        WAVEFORM_INPUT_BLOCK_SAMPLES if on_pcm_block is not None else 4096
    )
    try:
        with sd.InputStream(
            samplerate=sample_rate,
            channels=1,
            dtype=np.int16,
            blocksize=block_samples,
        ) as stream:
            while not stop_event.is_set():
                data, _overflowed = stream.read(block_samples)
                if not data.size:
                    continue
                if pause_event is not None and pause_event.is_set():
                    continue
                frames.append(data.copy())
                if on_pcm_block is not None:
                    try:
                        on_pcm_block(data.copy())
                    except Exception:
                        pass
                if on_audio_chunk is not None:
                    now = time.monotonic()
                    if now >= next_chunk_at and frames:
                        big = np.concatenate(frames, axis=0)
                        delta = big[last_flushed_samples:]
                        if delta.size >= LIVE_STT_MIN_CHUNK_SAMPLES:
                            fd, tmp_name = tempfile.mkstemp(suffix=".wav", prefix="stt_chunk_")
                            os.close(fd)
                            chunk_path = Path(tmp_name)
                            werr = write_mono_int16_wav(chunk_path, delta, sample_rate)
                            if werr is None:
                                on_audio_chunk(chunk_path)
                                last_flushed_samples = int(big.shape[0])
                            else:
                                try:
                                    chunk_path.unlink(missing_ok=True)
                                except OSError:
                                    pass
                        next_chunk_at = now + chunk_interval_sec
    except Exception as exc:
        return str(exc)

    if not frames:
        return "No audio captured."

    audio = np.concatenate(frames, axis=0)
    werr = write_mono_int16_wav(output_path, audio, sample_rate)
    if werr is not None:
        return werr
    return None


def generate_journal_report_from_sources(journal_text: str, speech_transcript: str) -> str:
    system_message = (
        "You produce clear, professional summaries of daily work notes. "
        "Highlight key activities, decisions, blockers, and suggested follow-ups. "
        "Use short sections with bullets where appropriate."
    )
    user_content = (
        "### Journal text\n"
        + (journal_text.strip() or "(empty)")
        + "\n\n### Speech-to-text transcript\n"
        + (speech_transcript.strip() or "(none)")
    )
    messages: List[Dict[str, object]] = [
        {"role": "system", "content": system_message},
        {"role": "user", "content": user_content},
    ]
    return chat_completion(
        messages,
        model=OPENAI_THINKING_MODEL,
        reasoning_effort="high",
    )


def open_journal_window_editor(draft_data: Optional[Dict[str, object]] = None) -> bool:
    if tk is None or messagebox is None:
        print("Window mode is not available on this Python setup.")
        return False

    now = datetime.now()
    default_date = now.strftime("%m/%d/%Y")
    default_time = now.strftime("%I:%M%p").lstrip("0")
    draft_text = ""
    draft_speech = ""
    draft_report = ""
    draft_date = default_date
    draft_time = default_time
    edit_target_sheet = ""
    edit_target_row = 0
    if draft_data:
        draft_text = str(draft_data.get("text", "") or "")
        draft_speech = str(draft_data.get("speech_transcript", "") or "")
        draft_report = str(draft_data.get("ai_report", "") or "")
        draft_date = str(draft_data.get("date", default_date) or default_date)
        draft_time = str(draft_data.get("time", default_time) or default_time)
        edit_target_sheet = str(draft_data.get("edit_target_sheet", "") or "")
        try:
            edit_target_row = int(draft_data.get("edit_target_row", 0) or 0)
        except (TypeError, ValueError):
            edit_target_row = 0

    root = tk.Tk()
    root_prefs = load_preferences()
    ui_lang_holder: List[str] = [
        normalize_ui_language(str(root_prefs.get(UI_LANGUAGE_PREF_KEY, "en")))
    ]

    def tr(key: str, **kwargs: object) -> str:
        return ui_translate(ui_lang_holder[0], key, **kwargs)

    window_app_name = root_prefs.get("app_name", "Daily Logger").strip() or "Daily Logger"
    root.title(window_app_name)
    root.geometry("1360x720")
    root.minsize(1020, 620)
    theme_holder: List[JournalWindowThemeSpec] = [load_journal_window_theme_spec()]

    def th() -> JournalWindowThemeSpec:
        return theme_holder[0]


    t_init = th()
    root.configure(bg=t_init.surface)
    startup_total_steps = 6
    startup_progress = {"value": 0}
    startup_overlay = tk.Frame(root, bg=t_init.surface, bd=0, highlightthickness=0)
    startup_overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
    startup_box = tk.Frame(startup_overlay, bg=t_init.surface, bd=0, highlightthickness=0)
    startup_box.place(relx=0.5, rely=0.5, anchor="center")
    splash_w = 460
    splash_title = tk.Label(
        startup_box,
        text=tr("splash.title", app=window_app_name),
        bg=t_init.surface,
        fg=t_init.text,
        font=("Segoe UI", 11, "bold"),
        anchor="w",
    )
    splash_title.pack(fill="x", padx=16, pady=(14, 6))
    splash_detail = tk.Label(
        startup_box,
        text=tr("splash.detail.theme"),
        bg=t_init.surface,
        fg=t_init.muted,
        font=("Segoe UI", 9),
        anchor="w",
    )
    splash_detail.pack(fill="x", padx=16, pady=(0, 10))
    startup_bar: Any
    startup_canvas: Optional[Any] = None
    startup_fill: Optional[Any] = None
    if ttk is not None:
        startup_bar = ttk.Progressbar(
            startup_box,
            orient="horizontal",
            mode="determinate",
            maximum=float(startup_total_steps),
            length=splash_w - 32,
        )
        startup_bar.pack(padx=16, pady=(0, 14))
    else:
        startup_canvas = tk.Canvas(
            startup_box,
            width=splash_w - 32,
            height=16,
            bg=t_init.field,
            highlightthickness=1,
            highlightbackground=t_init.border,
        )
        startup_canvas.pack(padx=16, pady=(0, 14))
        startup_fill = startup_canvas.create_rectangle(0, 0, 0, 16, fill=t_init.accent, width=0)
        startup_bar = None

    def _startup_step(detail_key: str) -> None:
        startup_progress["value"] = min(startup_total_steps, startup_progress["value"] + 1)
        splash_detail.config(text=tr(detail_key))
        if startup_bar is not None:
            startup_bar["value"] = float(startup_progress["value"])
        elif startup_canvas is not None and startup_fill is not None:
            bar_w = int((splash_w - 32) * (startup_progress["value"] / startup_total_steps))
            startup_canvas.coords(startup_fill, 0, 0, bar_w, 16)
        startup_overlay.lift()
        root.update_idletasks()

    _startup_step("splash.detail.theme")
    _jw_style: Any = None
    if ttk is not None:
        _jw_style = ttk.Style(root)
        try:
            _jw_style.theme_use("clam")
        except tk.TclError:
            pass
        _jw_style.configure("Journal.TCombobox", **t_init.ttk_combobox_kwargs())
        if t_init.is_dark:
            _jw_style.map(
                "Journal.TCombobox",
                fieldbackground=[
                    ("readonly", t_init.field),
                    ("disabled", t_init.btn_disabled),
                ],
                selectbackground=[("readonly", t_init.accent)],
                selectforeground=[("readonly", "white")],
            )
        else:
            _jw_style.map(
                "Journal.TCombobox",
                fieldbackground=[
                    ("readonly", t_init.field),
                    ("disabled", t_init.btn_disabled),
                ],
            )
    # Bring the journal window to front so it does not hide behind the console.
    root.lift()
    root.attributes("-topmost", True)
    root.after(250, lambda: root.attributes("-topmost", False))
    root.focus_force()
    # Mutable so console commands (like editprev) can toggle edit mode in-place.
    is_edit_mode = {"v": bool(edit_target_sheet and edit_target_row > 0)}

    shell = tk.Frame(root, bg=t_init.surface, bd=0, highlightthickness=0)
    shell.pack(fill="both", expand=True)
    shell.grid_rowconfigure(0, weight=1)
    shell.grid_columnconfigure(1, weight=1)
    shell.grid_columnconfigure(0, minsize=170)

    nav_rail = tk.Frame(shell, bg=t_init.panel, width=170, bd=0, highlightthickness=0)
    nav_rail.grid(row=0, column=0, sticky="nsw")
    nav_rail.grid_rowconfigure(100, weight=1)
    nav_rail.grid_columnconfigure(0, weight=1)
    nav_rail.grid_propagate(False)

    content_host = tk.Frame(shell, bg=t_init.surface, bd=0, highlightthickness=0)
    content_host.grid(row=0, column=1, sticky="nsew")
    content_host.grid_rowconfigure(0, weight=1)
    content_host.grid_rowconfigure(1, weight=0)
    content_host.grid_columnconfigure(0, weight=1)
    console_input_holder: Dict[str, Any] = {"row": None}

    journal_page = tk.Frame(content_host, bg=t_init.surface, bd=0, highlightthickness=0)
    ai_recap_page = tk.Frame(content_host, bg=t_init.surface, bd=0, highlightthickness=0)
    chatbot_page = tk.Frame(content_host, bg=t_init.surface, bd=0, highlightthickness=0)
    console_page = tk.Frame(content_host, bg=t_init.surface, bd=0, highlightthickness=0)
    settings_page = tk.Frame(content_host, bg=t_init.surface, bd=0, highlightthickness=0)
    for _p in (journal_page, ai_recap_page, chatbot_page, console_page, settings_page):
        _p.grid(row=0, column=0, sticky="nsew")
    # Ensure first paint shows Journal instead of last-created stacked page.
    journal_page.tkraise()

    nav_collapsed = {"value": False}
    nav_animating = {"value": False}
    nav_full_width = 170
    nav_restore_page = {"key": "journal"}
    nav_title = tk.Label(
        nav_rail,
        text=window_app_name,
        bg=t_init.panel,
        fg=t_init.muted,
        font=("Segoe UI", 10, "bold"),
    )
    nav_title.grid(row=0, column=0, sticky="w", padx=(12, 0), pady=(14, 10))

    nav_buttons: Dict[str, Any] = {}
    active_page = {"key": "journal"}
    active_page_frame: Dict[str, Any] = {"frame": None}
    page_leave_reset_handlers: Dict[str, Callable[[], None]] = {}

    def _layout_console_row(frame: Any) -> None:
        frame.update_idletasks()
        fw = frame.winfo_width()
        reveal_width = nav_summon_btn.winfo_width() if nav_collapsed["value"] else 0
        left_margin = 20 + (reveal_width + 8 if nav_collapsed["value"] else 0)
        right_margin = 20
        if save_entry_btn_holder.get("btn") is not None and frame is journal_page:
            save_x = save_entry_btn.winfo_x()
            save_w = save_entry_btn.winfo_width()
            if save_x > 0 and save_w > 0:
                right_margin = max(right_margin, fw - save_x + 8)
        row_w = max(280, fw - left_margin - right_margin)
        console_row = console_input_holder.get("row")
        if console_row is not None:
            console_row.place(
                in_=frame,
                x=left_margin,
                rely=1.0,
                y=-12,
                anchor="sw",
                width=row_w,
            )
            console_row.lift()

    def show_page(page_key: str) -> None:
        page_map = {
            "journal": journal_page,
            "ai_recap": ai_recap_page,
            "chatbot": chatbot_page,
            "console": console_page,
            "settings": settings_page,
        }
        prev_key = active_page["key"]
        if page_key == "console":
            _clear_console_hint()
        frame = page_map.get(page_key, journal_page)
        frame.tkraise()
        active_page["key"] = page_key
        if prev_key != page_key:
            reset_fn = page_leave_reset_handlers.get(prev_key)
            if reset_fn is not None:
                reset_fn()
        active_page_frame["frame"] = frame
        console_row = console_input_holder.get("row")
        if console_row is not None:
            _layout_console_row(frame)
        for key, btn in nav_buttons.items():
            if key == page_key:
                btn.config(bg=th().accent, fg="white")
            else:
                btn.config(bg=th().btn_secondary, fg=th().text)

    page_toggle_buttons: List[Any] = []
    nav_summon_btn = tk.Button(
        content_host,
        text="▶",
        bg=t_init.toolbar_btn_config()[0],
        fg=t_init.toolbar_btn_config()[1],
        activebackground=t_init.toolbar_btn_config()[2],
        activeforeground=t_init.toolbar_btn_config()[3],
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        padx=2,
        pady=12,
        cursor="hand2",
        bd=0,
        highlightthickness=0,
        width=1,
    )

    def _place_page_toggle(btn: Any) -> None:
        if nav_animating["value"]:
            return
        if nav_collapsed["value"]:
            btn.place_forget()
        else:
            nav_rail.update_idletasks()
            ph = nav_rail.winfo_height()
            y = max(56, (ph // 2) - 14)
            # Place on the right seam of the Pages rail.
            x = max(0, nav_rail.winfo_width() - 12)
            btn.place(x=x, y=y)

    def _place_nav_summon() -> None:
        if nav_animating["value"]:
            return
        content_host.update_idletasks()
        jh = content_host.winfo_height()
        y = max(56, (jh // 2) - 14)
        nav_summon_btn.place(x=0, y=y)
        nav_summon_btn.lift()

    def _register_page_toggle(parent: Any) -> Any:
        if page_toggle_buttons:
            return page_toggle_buttons[0]
        btn = tk.Button(
            nav_rail,
            text="◀",
            bg=t_init.toolbar_btn_config()[0],
            fg=t_init.toolbar_btn_config()[1],
            activebackground=t_init.toolbar_btn_config()[2],
            activeforeground=t_init.toolbar_btn_config()[3],
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=2,
            pady=12,
            cursor="hand2",
            bd=0,
            highlightthickness=0,
            width=1,
        )
        btn.config(command=lambda b=btn: set_nav_visible(False))
        bind_button_hover_if_enabled(
            btn,
            lambda: (
                "normal",
                th().toolbar_btn_config()[0],
                th().toolbar_btn_config()[1],
                th().toolbar_btn_config()[2],
                th().toolbar_btn_config()[3],
            ),
            lambda: th().toolbar_hover()[0],
            lambda: th().toolbar_hover()[1],
        )
        page_toggle_buttons.append(btn)
        _place_page_toggle(btn)
        nav_rail.bind(
            "<Configure>",
            lambda _e, b=btn: _place_page_toggle(b) if not nav_animating["value"] else None,
            add="+",
        )
        return btn

    def set_nav_visible(visible: bool) -> None:
        if nav_animating["value"]:
            return
        target_collapsed = not visible
        if nav_collapsed["value"] == target_collapsed:
            return

        nav_animating["value"] = True
        nav_collapsed["value"] = not visible

        def _animate_width(start: int, target: int, done: Callable[[], None]) -> None:
            duration_ms = 220.0
            t0 = time.perf_counter()

            def _tick() -> None:
                elapsed = (time.perf_counter() - t0) * 1000.0
                p = min(1.0, elapsed / duration_ms)
                eased = 1.0 - ((1.0 - p) ** 3)
                nxt = int(round(start + (target - start) * eased))
                shell.grid_columnconfigure(0, minsize=nxt)
                nav_rail.config(width=nxt)
                if p >= 1.0:
                    done()
                else:
                    root.after(16, _tick)

            _tick()

        if visible:
            nav_rail.grid()
            nav_summon_btn.place_forget()
            shell.grid_columnconfigure(0, minsize=0)
            nav_rail.config(width=0)

            def _on_expand_done() -> None:
                nav_animating["value"] = False
                for btn in page_toggle_buttons:
                    _place_page_toggle(btn)
                restore_key = nav_restore_page.get("key", "journal")
                if restore_key in ("journal", "ai_recap", "chatbot", "console", "settings"):
                    show_page(restore_key)

            _animate_width(0, nav_full_width, _on_expand_done)
        else:
            for btn in page_toggle_buttons:
                btn.place_forget()
            nav_restore_page["key"] = active_page["key"]

            def _on_collapse_done() -> None:
                nav_animating["value"] = False
                shell.grid_columnconfigure(0, minsize=0)
                nav_rail.config(width=0)
                nav_rail.grid_remove()
                _place_nav_summon()

            _animate_width(nav_full_width, 0, _on_collapse_done)

    top = tk.Frame(journal_page, bg=t_init.panel, bd=0, highlightthickness=0)
    top.pack(fill="x", padx=t_init.pad_outer, pady=t_init.pad_top_y)
    top.grid_columnconfigure(5, weight=1)
    top.grid_columnconfigure(6, weight=0)
    _register_page_toggle(journal_page)
    date_lbl = tk.Label(
        top,
        text="Date (mm/dd/yyyy):",
        bg=t_init.panel,
        fg=t_init.muted,
        font=t_init.date_label_font,
    )
    date_lbl.grid(row=0, column=0, sticky="w", padx=(12, 0), pady=12)
    date_entry: object
    if DateEntry is not None:
        date_entry = DateEntry(
            top,
            width=14,
            date_pattern="mm/dd/yyyy",
            state="normal",  # Keep typing enabled while allowing popup calendar selection.
            background=t_init.field,
            foreground=t_init.text,
            borderwidth=1,
        )
        date_entry.grid(row=0, column=1, padx=(8, 20), pady=12, sticky="w")
        try:
            date_entry.set_date(draft_date)
        except Exception:
            date_entry.delete(0, "end")
            date_entry.insert(0, draft_date)
    else:
        date_entry = tk.Entry(
            top,
            width=16,
            bg=t_init.field,
            fg=t_init.text,
            insertbackground=t_init.text,
            relief="flat",
            highlightthickness=1,
            highlightbackground=t_init.border,
            highlightcolor=t_init.accent,
            font=("Segoe UI", 10),
        )
        date_entry.grid(row=0, column=1, padx=(8, 20), pady=12, sticky="w")
        date_entry.insert(0, draft_date)
    time_lbl = tk.Label(
        top,
        text="Time (hh:mmAM/PM or rn):",
        bg=t_init.panel,
        fg=t_init.muted,
        font=t_init.date_label_font,
    )
    time_lbl.grid(row=0, column=2, sticky="w", pady=12)
    time_entry = tk.Entry(
        top,
        width=16,
        bg=t_init.field,
        fg=t_init.text,
        insertbackground=t_init.text,
        relief="flat",
        highlightthickness=1,
        highlightbackground=t_init.border,
        highlightcolor=t_init.accent,
        font=("Segoe UI", 10),
    )
    time_entry.grid(row=0, column=3, padx=(8, 0), pady=12)
    time_entry.insert(0, draft_time)
    def update_date_time_to_now() -> None:
        current_now = datetime.now()
        date_entry.delete(0, "end")
        date_entry.insert(0, current_now.strftime("%m/%d/%Y"))
        time_entry.delete(0, "end")
        time_entry.insert(0, current_now.strftime("%I:%M%p").lstrip("0"))
        save_draft()
    _ut_bg, _ut_fg, _ut_abg, _ut_afg = t_init.toolbar_btn_config()
    update_time_btn = tk.Button(
        top,
        text="Update Time",
        command=update_date_time_to_now,
        bg=_ut_bg,
        fg=_ut_fg,
        activebackground=_ut_abg,
        activeforeground=_ut_afg,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=12,
        pady=6,
        cursor="hand2",
    )
    update_time_btn.grid(row=0, column=4, padx=(12, 12), sticky="w")
    bind_button_hover_if_enabled(
        update_time_btn,
        lambda: th().toolbar_bind_rest(),
        lambda: th().toolbar_hover()[0],
        lambda: th().toolbar_hover()[1],
    )
    find_row = tk.Frame(journal_page, bg=t_init.panel, bd=0, highlightthickness=0)
    find_row.pack(fill="x", padx=t_init.pad_outer, pady=(0, 6))
    find_row.grid_columnconfigure(8, weight=1)
    find_lbl = tk.Label(
        find_row,
        text="Find:",
        bg=t_init.panel,
        fg=t_init.muted,
        font=("Segoe UI", 9, "bold"),
    )
    find_lbl.grid(row=0, column=0, sticky="w", padx=(12, 6), pady=8)
    find_scope_var = tk.StringVar(value="all")
    find_scope_all_rb = tk.Radiobutton(
        find_row,
        text="All",
        value="all",
        variable=find_scope_var,
        bg=t_init.panel,
        fg=t_init.muted,
        activebackground=t_init.panel,
        activeforeground=t_init.text,
        selectcolor=t_init.panel,
        font=("Segoe UI", 9),
        highlightthickness=0,
        bd=0,
        padx=4,
    )
    find_scope_all_rb.grid(row=0, column=1, sticky="w", padx=(2, 4), pady=8)
    find_scope_one_rb = tk.Radiobutton(
        find_row,
        text="Current box",
        value="one",
        variable=find_scope_var,
        bg=t_init.panel,
        fg=t_init.muted,
        activebackground=t_init.panel,
        activeforeground=t_init.text,
        selectcolor=t_init.panel,
        font=("Segoe UI", 9),
        highlightthickness=0,
        bd=0,
        padx=4,
    )
    find_scope_one_rb.grid(row=0, column=2, sticky="w", padx=(0, 8), pady=8)
    find_var = tk.StringVar(value="")
    find_entry = tk.Entry(
        find_row,
        textvariable=find_var,
        width=28,
        bg=t_init.field,
        fg=t_init.text,
        insertbackground=t_init.text,
        relief="flat",
        highlightthickness=1,
        highlightbackground=t_init.border,
        highlightcolor=t_init.accent,
        font=("Segoe UI", 10),
    )
    find_entry.grid(row=0, column=3, sticky="w", pady=8)
    find_case_var = tk.BooleanVar(value=False)
    find_case_chk = tk.Checkbutton(
        find_row,
        text="Case",
        variable=find_case_var,
        bg=t_init.panel,
        fg=t_init.muted,
        activebackground=t_init.panel,
        activeforeground=t_init.text,
        selectcolor=t_init.panel,
        font=("Segoe UI", 9),
        highlightthickness=0,
        bd=0,
        padx=4,
    )
    find_case_chk.grid(row=0, column=4, sticky="w", padx=(8, 0), pady=8)
    find_word_var = tk.BooleanVar(value=False)
    find_word_chk = tk.Checkbutton(
        find_row,
        text="Word",
        variable=find_word_var,
        bg=t_init.panel,
        fg=t_init.muted,
        activebackground=t_init.panel,
        activeforeground=t_init.text,
        selectcolor=t_init.panel,
        font=("Segoe UI", 9),
        highlightthickness=0,
        bd=0,
        padx=4,
    )
    find_word_chk.grid(row=0, column=5, sticky="w", padx=(4, 0), pady=8)
    bind_hover_tooltip(
        find_scope_all_rb,
        lambda: tr("tip.find_all"),
    )
    bind_hover_tooltip(
        find_scope_one_rb,
        lambda: tr("tip.find_one"),
    )
    bind_hover_tooltip(
        find_case_chk,
        lambda: tr("tip.find_case"),
    )
    bind_hover_tooltip(
        find_word_chk,
        lambda: tr("tip.find_word"),
    )
    find_status = tk.Label(
        find_row,
        text="",
        bg=t_init.panel,
        fg=t_init.muted,
        font=("Segoe UI", 9),
    )
    find_status.grid(row=0, column=6, sticky="w", padx=(8, 0), pady=8)
    find_prev_btn = tk.Button(
        find_row,
        text="Prev",
        bg=_ut_bg,
        fg=_ut_fg,
        activebackground=_ut_abg,
        activeforeground=_ut_afg,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=10,
        pady=4,
        cursor="hand2",
    )
    find_prev_btn.grid(row=0, column=7, sticky="e", padx=(10, 6), pady=8)
    find_next_btn = tk.Button(
        find_row,
        text="Next",
        bg=_ut_bg,
        fg=_ut_fg,
        activebackground=_ut_abg,
        activeforeground=_ut_afg,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=10,
        pady=4,
        cursor="hand2",
    )
    find_next_btn.grid(row=0, column=8, sticky="e", padx=6, pady=8)
    find_close_btn = tk.Button(
        find_row,
        text="Close",
        bg=_ut_bg,
        fg=_ut_fg,
        activebackground=_ut_abg,
        activeforeground=_ut_afg,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=10,
        pady=4,
        cursor="hand2",
    )
    find_close_btn.grid(row=0, column=9, sticky="e", padx=(6, 12), pady=8)
    find_row.pack_forget()

    center = tk.Frame(journal_page, bg=t_init.surface)
    center.pack(
        fill="both",
        expand=True,
        padx=t_init.pad_outer,
        pady=(0, t_init.pad_center_y),
    )
    center.grid_columnconfigure(0, weight=2)
    center.grid_columnconfigure(1, weight=2)
    center.grid_rowconfigure(0, weight=1)

    left_col = tk.Frame(center, bg=t_init.surface)
    left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
    left_col.grid_columnconfigure(0, weight=1)
    left_col.grid_rowconfigure(1, weight=1)
    journal_title_lbl = tk.Label(
        left_col,
        text="Journal Text",
        bg=t_init.surface,
        fg=t_init.muted,
        font=t_init.section_label_font,
    )
    journal_title_lbl.grid(row=0, column=0, sticky="w", pady=(0, 6))
    editor_frame = tk.Frame(left_col, bg=t_init.panel, bd=0, highlightthickness=0)
    editor_frame.grid(row=1, column=0, sticky="nsew")
    editor_frame.grid_rowconfigure(0, weight=1)
    editor_frame.grid_columnconfigure(0, weight=1)
    text_box = tk.Text(
        editor_frame,
        wrap="word",
        height=12,
        undo=True,
        autoseparators=True,
        maxundo=-1,
        bg=t_init.field,
        fg=t_init.text,
        insertbackground=t_init.text,
        relief="flat",
        padx=12,
        pady=12,
        font=("Segoe UI", 11),
        highlightthickness=1,
        highlightbackground=t_init.border,
        highlightcolor=t_init.accent,
    )
    scroll_bar = tk.Scrollbar(
        editor_frame,
        command=text_box.yview,
        bg=t_init.panel,
        troughcolor=t_init.field,
        activebackground=t_init.accent,
        bd=0,
        highlightthickness=0,
        width=11,
    )
    text_box.configure(yscrollcommand=scroll_bar.set)
    text_box.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
    scroll_bar.grid(row=0, column=1, sticky="ns", padx=(0, 12), pady=12)
    text_box.insert("1.0", draft_text)
    text_box.focus_set()
    root.after(50, text_box.focus_set)

    right_col = tk.Frame(center, bg=t_init.surface)
    right_col.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
    right_col.grid_rowconfigure(0, weight=1)
    right_col.grid_rowconfigure(1, weight=1)
    right_col.grid_columnconfigure(0, weight=1)

    stt_outer = tk.Frame(right_col, bg=t_init.surface)
    stt_outer.grid(row=0, column=0, sticky="nsew", pady=(0, 6))
    stt_outer.grid_columnconfigure(0, weight=1)
    stt_header = tk.Frame(stt_outer, bg=t_init.surface)
    stt_header.grid(row=0, column=0, sticky="ew", pady=(0, 4))
    stt_header.grid_columnconfigure(1, weight=1)
    stt_title_lbl = tk.Label(
        stt_header,
        text="Speech to text",
        bg=t_init.surface,
        fg=t_init.muted,
        font=t_init.section_label_font,
    )
    stt_title_lbl.grid(row=0, column=0, sticky="w")
    stt_saved_path_var = tk.StringVar(value="")
    stt_saved_path_entry = tk.Entry(
        stt_header,
        textvariable=stt_saved_path_var,
        state="readonly",
        readonlybackground=t_init.surface,
        fg=t_init.muted,
        font=("Segoe UI", 8),
        relief="flat",
        borderwidth=0,
        highlightthickness=0,
        highlightbackground=t_init.surface,
        insertwidth=0,
        justify="right",
        takefocus=1,
        selectbackground=t_init.accent,
        selectforeground="white",
        cursor="xterm",
    )
    stt_saved_path_entry.grid(row=0, column=1, sticky="ew", padx=(14, 8))

    def _set_stt_saved_path_display(text: str) -> None:
        stt_saved_path_entry.config(state="normal")
        stt_saved_path_var.set(text)
        stt_saved_path_entry.config(state="readonly")

    def open_journal_recording_folder() -> None:
        try:
            RECORDING_DIR.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass
        open_path_with_default_app(RECORDING_DIR)

    open_recording_btn = tk.Button(
        stt_header,
        text="Open",
        command=open_journal_recording_folder,
        bg=_ut_bg,
        fg=_ut_fg,
        activebackground=_ut_abg,
        activeforeground=_ut_afg,
        relief="flat",
        font=("Segoe UI", 8, "bold"),
        padx=6,
        pady=2,
        cursor="hand2",
    )
    open_recording_btn.grid(row=0, column=2, sticky="e")

    def open_recording_tooltip_text() -> str:
        return tr("tip.open_recordings")

    bind_hover_tooltip(open_recording_btn, open_recording_tooltip_text)
    bind_button_hover_if_enabled(
        open_recording_btn,
        lambda: th().toolbar_bind_rest(),
        lambda: th().toolbar_hover()[0],
        lambda: th().toolbar_hover()[1],
    )

    stt_top = tk.Frame(stt_outer, bg=t_init.panel, bd=0, highlightthickness=0)
    stt_top.grid(row=1, column=0, sticky="ew", pady=(0, 6))
    stt_top.grid_columnconfigure(5, weight=1)
    lang_var = tk.StringVar(value="Auto")

    stt_status = tk.Label(
        stt_top,
        text="",
        bg=t_init.panel,
        fg=t_init.muted,
        font=("Segoe UI", 9),
        anchor="w",
        justify="left",
        wraplength=420,
    )

    if ttk is not None:
        lang_combo = ttk.Combobox(
            stt_top,
            textvariable=lang_var,
            values=("Auto", "English", "简体中文"),
            state="readonly",
            width=11,
            style="Journal.TCombobox",
        )
    else:
        lang_combo = tk.OptionMenu(stt_top, lang_var, "Auto", "English", "简体中文")
        lang_combo.config(bg=t_init.panel, fg=t_init.text, highlightthickness=0)

    stt_frame = tk.Frame(stt_outer, bg=t_init.panel, bd=0, highlightthickness=0)
    stt_frame.grid(row=2, column=0, sticky="nsew")
    stt_frame.grid_rowconfigure(0, weight=0)
    stt_frame.grid_rowconfigure(1, weight=1)
    stt_frame.grid_columnconfigure(0, weight=1)
    stt_frame.grid_columnconfigure(2, minsize=JOURNAL_SIDE_ACTION_GRID_MINSIZE)
    stt_outer.grid_rowconfigure(2, weight=1)

    wave_canvas = tk.Canvas(
        stt_frame,
        height=52,
        bg=t_init.field,
        highlightthickness=1,
        highlightbackground=t_init.border,
        highlightcolor=t_init.accent,
    )
    wave_canvas.grid(row=0, column=0, columnspan=3, sticky="ew", padx=10, pady=(10, 4))

    stt_box = tk.Text(
        stt_frame,
        wrap="word",
        height=8,
        undo=True,
        autoseparators=True,
        maxundo=-1,
        bg=t_init.field,
        fg=t_init.text,
        insertbackground=t_init.text,
        relief="flat",
        padx=10,
        pady=10,
        font=("Segoe UI", 10),
        highlightthickness=1,
        highlightbackground=t_init.border,
        highlightcolor=t_init.accent,
    )
    stt_scroll = tk.Scrollbar(
        stt_frame,
        command=stt_box.yview,
        bg=t_init.panel,
        troughcolor=t_init.field,
        activebackground=t_init.accent,
        bd=0,
        highlightthickness=0,
        width=11,
    )
    stt_box.configure(yscrollcommand=stt_scroll.set)
    stt_box.grid(row=1, column=0, sticky="nsew", padx=(10, 0), pady=(4, 10))
    stt_scroll.grid(row=1, column=1, sticky="ns", padx=(0, 2), pady=(4, 10))
    transcribe_hover = tk.Frame(stt_frame, bg=t_init.panel)
    transcribe_hover.grid(row=1, column=2, sticky="ns", padx=(2, 10), pady=(4, 10))
    _tid = t_init.transcribe_idle_disabled_config()
    transcribe_btn = tk.Button(
        transcribe_hover,
        text="Transcribe",
        state="disabled",
        width=JOURNAL_SIDE_ACTION_BTN_WIDTH_CH,
        bg=_tid[0],
        fg=_tid[1],
        activebackground=_tid[2],
        activeforeground=_tid[3],
        disabledforeground=_tid[4],
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=10,
        pady=8,
        cursor="hand2",
    )
    transcribe_btn.pack()
    stt_box.insert("1.0", draft_speech)

    report_outer = tk.Frame(right_col, bg=t_init.surface)
    report_outer.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
    report_outer.grid_rowconfigure(1, weight=1)
    report_outer.grid_columnconfigure(0, weight=1)
    report_header = tk.Frame(report_outer, bg=t_init.surface)
    report_header.grid(row=0, column=0, sticky="ew", pady=(0, 4))
    report_header.grid_columnconfigure(1, weight=1)
    report_title_lbl = tk.Label(
        report_header,
        text="AI report",
        bg=t_init.surface,
        fg=t_init.muted,
        font=t_init.section_label_font,
    )
    report_title_lbl.grid(row=0, column=0, sticky="w")
    report_status = tk.Label(
        report_header,
        text="",
        bg=t_init.surface,
        fg=t_init.muted,
        font=("Segoe UI", 9),
        anchor="e",
    )
    report_status.grid(row=0, column=1, sticky="e", padx=(8, 0))

    report_frame = tk.Frame(report_outer, bg=t_init.panel, bd=0, highlightthickness=0)
    report_frame.grid(row=1, column=0, sticky="nsew")
    report_frame.grid_rowconfigure(0, weight=1)
    report_frame.grid_columnconfigure(0, weight=1)
    report_frame.grid_columnconfigure(2, minsize=JOURNAL_SIDE_ACTION_GRID_MINSIZE)
    report_box = tk.Text(
        report_frame,
        wrap="word",
        height=8,
        undo=True,
        autoseparators=True,
        maxundo=-1,
        bg=t_init.field,
        fg=t_init.text,
        insertbackground=t_init.text,
        relief="flat",
        padx=10,
        pady=10,
        font=("Segoe UI", 10),
        highlightthickness=1,
        highlightbackground=t_init.border,
        highlightcolor=t_init.accent,
    )
    report_scroll = tk.Scrollbar(
        report_frame,
        command=report_box.yview,
        bg=t_init.panel,
        troughcolor=t_init.field,
        activebackground=t_init.accent,
        bd=0,
        highlightthickness=0,
        width=11,
    )
    report_box.configure(yscrollcommand=report_scroll.set)
    report_box.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=(4, 10))
    report_scroll.grid(row=0, column=1, sticky="ns", padx=(0, 2), pady=(4, 10))
    gen_report_hover = tk.Frame(report_frame, bg=t_init.panel)
    gen_report_hover.grid(row=0, column=2, sticky="ns", padx=(2, 10), pady=(4, 10))
    _, _gn, _gf, _gab, _gaf = t_init.gen_bind_rest()
    gen_button = tk.Button(
        gen_report_hover,
        text="Generate report",
        width=JOURNAL_SIDE_ACTION_BTN_WIDTH_CH,
        bg=_gn,
        fg=_gf,
        activebackground=_gab,
        activeforeground=_gaf,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=10,
        pady=8,
        cursor="hand2",
    )
    gen_button.pack()
    report_box.insert("1.0", draft_report)

    placeholder_frames: List[Any] = []
    placeholder_title_labels: List[Any] = []
    placeholder_body_labels: List[Any] = []

    placeholder_frames: List[Any] = []
    placeholder_title_labels: List[Any] = []
    placeholder_body_labels: List[Any] = []

    api_key_prompt_hooks: Dict[str, Callable[[], None]] = {}

    def build_ai_recap_and_chatbot_pages() -> None:
        _register_page_toggle(ai_recap_page)
        _register_page_toggle(chatbot_page)

        t0 = t_init
        _tb, _tf, _tab, _taf = t0.toolbar_btn_config()

        # --- Shared: append styled lines to a read-only transcript ---
        def _append_transcript(box: Any, role: str, body: str) -> None:
            box.config(state="normal")
            if role == "user":
                box.insert("end", tr("chat.you") + "\n", ("t_meta",))
                box.insert("end", (body or "").strip() + "\n\n", ("t_user",))
            else:
                box.insert("end", tr("chat.assistant") + "\n", ("t_meta",))
                box.insert("end", (body or "").strip() + "\n\n", ("t_bot",))
            box.config(state="disabled")
            box.see("end")

        # ========== AI Recap ==========
        recap_wrap = tk.Frame(ai_recap_page, bg=t0.surface)
        recap_wrap.pack(
            fill="both",
            expand=True,
            padx=t0.pad_outer,
            pady=(0, t0.pad_center_y + JOURNAL_WINDOW_CONSOLE_RESERVE_BOTTOM),
        )
        recap_wrap.grid_columnconfigure(0, weight=1)
        recap_wrap.grid_rowconfigure(3, weight=1)

        recap_title = tk.Label(
            recap_wrap,
            text=tr("recap.title"),
            bg=t0.surface,
            fg=t0.text,
            font=("Segoe UI", 15, "bold"),
            anchor="w",
        )
        recap_title.grid(row=0, column=0, sticky="w", pady=(0, 8))

        recap_top = tk.Frame(recap_wrap, bg=t0.panel, highlightthickness=1, highlightbackground=t0.border)
        recap_top.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        recap_top.grid_columnconfigure(4, weight=1)

        recap_thinking_var = tk.BooleanVar(value=False)
        recap_thinking_chk = tk.Checkbutton(
            recap_top,
            text=tr("recap.thinking"),
            variable=recap_thinking_var,
            bg=t0.panel,
            fg=t0.muted,
            activebackground=t0.panel,
            activeforeground=t0.text,
            selectcolor=t0.field,
            font=("Segoe UI", 9),
        )
        recap_thinking_chk.grid(row=0, column=0, padx=(10, 8), pady=8, sticky="w")

        recap_from_fr = tk.Frame(recap_top, bg=t0.panel)
        recap_from_fr.grid(row=0, column=1, padx=(0, 10), pady=8, sticky="w")
        recap_from_lbl = tk.Label(
            recap_from_fr,
            text=tr("recap.from"),
            bg=t0.panel,
            fg=t0.muted,
            font=t0.date_label_font,
        )
        recap_from_lbl.pack(side="left", padx=(0, 8))
        recap_from_de: Any = None
        recap_to_de: Any = None
        _today = datetime.now().date()
        if DateEntry is not None:
            recap_from_de = DateEntry(
                recap_from_fr,
                width=14,
                date_pattern="mm/dd/yyyy",
                state="normal",
                background=t0.field,
                foreground=t0.text,
                borderwidth=1,
            )
            recap_from_de.pack(side="left")
            try:
                recap_from_de.set_date(_today)
            except Exception:
                pass
        else:
            tk.Label(
                recap_from_fr,
                text=tr("recap.install_dates"),
                bg=t0.panel,
                fg=t0.muted,
                font=("Segoe UI", 9),
            ).pack(side="left")

        recap_to_var = tk.BooleanVar(value=False)
        recap_all_journal_var = tk.BooleanVar(value=False)
        recap_to_wrap = tk.Frame(recap_top, bg=t0.panel)
        recap_to_wrap.grid(row=0, column=2, padx=(0, 8), pady=8, sticky="w")
        recap_to_chk = tk.Checkbutton(
            recap_to_wrap,
            text=tr("recap.to_chk"),
            variable=recap_to_var,
            bg=t0.panel,
            fg=t0.muted,
            activebackground=t0.panel,
            activeforeground=t0.text,
            selectcolor=t0.field,
            font=("Segoe UI", 9),
        )
        recap_to_chk.pack(side="left")
        recap_all_journal_chk = tk.Checkbutton(
            recap_top,
            text=tr("recap.all_journal"),
            variable=recap_all_journal_var,
            bg=t0.panel,
            fg=t0.muted,
            activebackground=t0.panel,
            activeforeground=t0.text,
            selectcolor=t0.field,
            font=("Segoe UI", 9),
        )
        recap_all_journal_chk.grid(row=1, column=0, columnspan=8, sticky="w", padx=(10, 0), pady=(0, 6))
        bind_hover_tooltip(recap_all_journal_chk, lambda: tr("tip.recap_all_journal"))
        recap_through_fr = tk.Frame(recap_top, bg=t0.panel)
        recap_through_lbl = tk.Label(
            recap_through_fr,
            text=tr("recap.through"),
            bg=t0.panel,
            fg=t0.muted,
            font=t0.date_label_font,
        )
        recap_through_lbl.pack(side="left", padx=(0, 8))
        if DateEntry is not None:
            recap_to_de = DateEntry(
                recap_through_fr,
                width=14,
                date_pattern="mm/dd/yyyy",
                state="normal",
                background=t0.field,
                foreground=t0.text,
                borderwidth=1,
            )
            recap_to_de.pack(side="left")
            try:
                recap_to_de.set_date(_today)
            except Exception:
                pass
        else:
            tk.Label(
                recap_through_fr,
                text=tr("recap.through_placeholder"),
                bg=t0.panel,
                fg=t0.muted,
                font=("Segoe UI", 9),
            ).pack(side="left")

        recap_cal_row = tk.Frame(recap_wrap, bg=t0.surface)
        recap_cal_row.grid_columnconfigure(1, weight=1)
        recap_selected_dates: set = set()
        recap_calendar: Any = None
        if Calendar is not None:
            recap_calendar = Calendar(
                recap_cal_row,
                selectmode="day",
                showweeknumbers=False,
                background=t0.field,
                foreground=t0.text,
                headersbackground=t0.panel,
                headersforeground=t0.text,
                weekendbackground=t0.field,
                weekendforeground=t0.muted,
                normalbackground=t0.field,
                normalforeground=t0.text,
                othermonthbackground=t0.field,
                othermonthforeground=t0.muted,
                selectbackground=t0.accent,
                selectforeground="white",
                bordercolor=t0.border,
                font=("Segoe UI", 9),
            )
            recap_calendar.grid(row=0, column=0, sticky="nw", padx=(0, 12), pady=(0, 4))
        else:
            tk.Label(
                recap_cal_row,
                text=tr("recap.install_calendar"),
                bg=t0.surface,
                fg=t0.muted,
                font=("Segoe UI", 9),
                wraplength=400,
                justify="left",
            ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        recap_sel_lbl = tk.Label(
            recap_cal_row,
            text=tr("recap.selected.none"),
            bg=t0.surface,
            fg=t0.muted,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
        )
        recap_sel_lbl.grid(row=0, column=1, sticky="nw", pady=(0, 4))

        recap_cal_marks_tag = "recap_sel"

        def recap_refresh_cal_marks() -> None:
            if recap_calendar is None:
                return
            try:
                recap_calendar.calevent_remove("all")
            except Exception:
                pass
            t = th()
            for d in recap_selected_dates:
                try:
                    recap_calendar.calevent_create(d, "", recap_cal_marks_tag)
                except Exception:
                    pass
            try:
                recap_calendar.tag_config(recap_cal_marks_tag, background=t.accent, foreground="white")
            except Exception:
                pass

        def recap_update_sel_label() -> None:
            if recap_all_journal_var.get():
                recap_sel_lbl.config(text=tr("recap.all_journal_active"))
                return
            if not recap_selected_dates:
                recap_sel_lbl.config(text=tr("recap.selected.none"))
                return
            ordered = sorted(recap_selected_dates)
            parts = [x.strftime("%m/%d/%Y") for x in ordered]
            recap_sel_lbl.config(text=tr("recap.selected.prefix") + ", ".join(parts))

        def recap_sync_to_checkbox() -> None:
            if recap_all_journal_var.get():
                recap_to_chk.config(state="disabled")
                return
            if recap_to_var.get():
                return
            if len(recap_selected_dates) > 1:
                recap_to_chk.config(state="disabled")
            else:
                recap_to_chk.config(state="normal")

        def recap_to_tooltip() -> str:
            if str(recap_to_chk.cget("state")) == "disabled":
                return tr("tip.recap_to_disabled")
            return tr("tip.recap_to")

        bind_hover_tooltip(recap_to_wrap, recap_to_tooltip)

        def on_recap_calendar_toggle(_evt: Optional[Any] = None) -> None:
            if recap_all_journal_var.get() or recap_to_var.get() or recap_calendar is None:
                return
            try:
                picked = recap_calendar.selection_get()
            except Exception:
                return
            if picked in recap_selected_dates:
                recap_selected_dates.remove(picked)
            else:
                recap_selected_dates.add(picked)
            recap_refresh_cal_marks()
            recap_update_sel_label()
            recap_sync_to_checkbox()

        if recap_calendar is not None:
            recap_calendar.bind("<<CalendarSelected>>", on_recap_calendar_toggle)

        recap_session: Dict[str, Any] = {"messages": [], "bootstrapped": False, "busy": False}
        recap_pending_images: List[Path] = []
        recap_pending_files: List[Path] = []

        def recap_refresh_date_controls() -> None:
            busy = bool(recap_session.get("busy"))
            if recap_all_journal_var.get():
                try:
                    recap_cal_row.grid_remove()
                except tk.TclError:
                    pass
                if recap_from_de is not None:
                    try:
                        recap_from_de.config(state="disabled")
                    except tk.TclError:
                        pass
                if recap_to_de is not None:
                    try:
                        recap_to_de.config(state="disabled")
                    except tk.TclError:
                        pass
                recap_to_chk.config(state="disabled")
                if recap_calendar is not None:
                    try:
                        recap_calendar.config(state="disabled")
                    except tk.TclError:
                        pass
                recap_all_journal_chk.config(state=("disabled" if busy else "normal"))
                recap_update_sel_label()
                return
            if recap_from_de is not None:
                try:
                    recap_from_de.config(state=("disabled" if busy else "normal"))
                except tk.TclError:
                    pass
            if recap_to_de is not None:
                try:
                    recap_to_de.config(state=("disabled" if busy else "normal"))
                except tk.TclError:
                    pass
            if recap_calendar is not None:
                try:
                    recap_calendar.config(state=("disabled" if busy else "normal"))
                except tk.TclError:
                    pass
            recap_all_journal_chk.config(state=("disabled" if busy else "normal"))
            if not recap_to_var.get():
                try:
                    recap_cal_row.grid(row=2, column=0, sticky="ew", pady=(0, 8))
                except tk.TclError:
                    pass
            else:
                try:
                    recap_cal_row.grid_remove()
                except tk.TclError:
                    pass
            recap_sync_to_checkbox()
            recap_update_sel_label()

        def on_recap_all_journal_toggle(*_a: Any) -> None:
            if recap_all_journal_var.get():
                recap_to_var.set(False)
            recap_refresh_date_controls()

        recap_all_journal_var.trace_add("write", on_recap_all_journal_toggle)

        def on_recap_to_mode(*_a: Any) -> None:
            if recap_to_var.get():
                recap_through_fr.grid(row=0, column=3, padx=(0, 8), pady=8, sticky="w")
                recap_cal_row.grid_remove()
                if len(recap_selected_dates) == 1 and DateEntry is not None:
                    only = next(iter(recap_selected_dates))
                    if recap_from_de is not None and recap_to_de is not None:
                        try:
                            recap_from_de.set_date(only)
                            recap_to_de.set_date(only)
                        except Exception:
                            pass
            else:
                recap_through_fr.grid_remove()
                recap_cal_row.grid(row=2, column=0, sticky="ew", pady=(0, 8))
                recap_refresh_cal_marks()
                recap_update_sel_label()
                recap_sync_to_checkbox()
            recap_refresh_date_controls()

        recap_to_var.trace_add("write", lambda *_: on_recap_to_mode())
        recap_through_fr.grid_remove()
        recap_cal_row.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        recap_refresh_date_controls()

        recap_mid = tk.Frame(recap_wrap, bg=t0.surface)
        recap_mid.grid(row=3, column=0, sticky="nsew", pady=(0, 8))
        recap_mid.grid_rowconfigure(0, weight=1)
        recap_mid.grid_columnconfigure(0, weight=1)

        recap_transcript = tk.Text(
            recap_mid,
            wrap="word",
            state="disabled",
            font=("Segoe UI", 10),
            bg=t0.field,
            fg=t0.text,
            insertbackground=t0.text,
            relief="flat",
            padx=12,
            pady=12,
            highlightthickness=1,
            highlightbackground=t0.border,
            highlightcolor=t0.accent,
        )
        recap_ts = tk.Scrollbar(
            recap_mid,
            command=recap_transcript.yview,
            bg=t0.panel,
            troughcolor=t0.field,
            activebackground=t0.accent,
            bd=0,
            highlightthickness=0,
            width=11,
        )
        recap_transcript.configure(yscrollcommand=recap_ts.set)
        recap_transcript.grid(row=0, column=0, sticky="nsew")
        recap_ts.grid(row=0, column=1, sticky="ns")
        recap_transcript.tag_configure("t_meta", foreground=t0.muted, font=("Segoe UI", 9, "bold"))
        recap_transcript.tag_configure("t_user", foreground=t0.text, font=("Segoe UI", 10))
        recap_transcript.tag_configure("t_bot", foreground=t0.text, font=("Segoe UI", 10))

        recap_attach_row = tk.Frame(recap_wrap, bg=t0.surface)
        recap_attach_row.grid(row=4, column=0, sticky="ew", pady=(0, 4))
        recap_pending_lbl = tk.Label(
            recap_attach_row,
            text=tr("recap.attachments", what=tr("recap.attachments_none")),
            bg=t0.surface,
            fg=t0.muted,
            font=("Segoe UI", 9),
            anchor="w",
        )
        recap_pending_lbl.pack(side="left", fill="x", expand=True)

        def recap_refresh_pending_lbl() -> None:
            bits = []
            if recap_pending_images:
                bits.append(tr("recap.n_images", n=len(recap_pending_images)))
            if recap_pending_files:
                bits.append(tr("recap.n_files", n=len(recap_pending_files)))
            what = ", ".join(bits) if bits else tr("recap.attachments_none")
            recap_pending_lbl.config(text=tr("recap.attachments", what=what))

        def recap_pick_image() -> None:
            p = filedialog.askopenfilename(
                title="Attach image",
                filetypes=[
                    ("Images", "*.png *.jpg *.jpeg *.gif *.webp"),
                    ("All files", "*.*"),
                ],
            )
            if p:
                recap_pending_images.append(Path(p))
                recap_refresh_pending_lbl()

        def recap_pick_file() -> None:
            p = filedialog.askopenfilename(title="Attach file", filetypes=[("Text / data", "*.*")])
            if p:
                recap_pending_files.append(Path(p))
                recap_refresh_pending_lbl()

        recap_img_btn = tk.Button(
            recap_attach_row,
            text=tr("recap.image"),
            command=recap_pick_image,
            bg=_tb,
            fg=_tf,
            activebackground=_tab,
            activeforeground=_taf,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=4,
            cursor="hand2",
        )
        recap_img_btn.pack(side="right", padx=(6, 0))
        recap_file_btn = tk.Button(
            recap_attach_row,
            text=tr("recap.file"),
            command=recap_pick_file,
            bg=_tb,
            fg=_tf,
            activebackground=_tab,
            activeforeground=_taf,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=4,
            cursor="hand2",
        )
        recap_file_btn.pack(side="right", padx=(6, 0))

        recap_bottom = tk.Frame(recap_wrap, bg=t0.panel, highlightthickness=1, highlightbackground=t0.border)
        recap_bottom.grid(row=5, column=0, sticky="ew", pady=(0, 0))
        recap_bottom.grid_columnconfigure(0, weight=1)

        recap_input = tk.Text(
            recap_bottom,
            height=3,
            wrap="word",
            font=("Segoe UI", 10),
            bg=t0.field,
            fg=t0.text,
            insertbackground=t0.text,
            relief="flat",
            padx=10,
            pady=8,
            highlightthickness=0,
        )
        recap_input.grid(row=0, column=0, sticky="ew", padx=(8, 4), pady=8)

        recap_btn_fr = tk.Frame(recap_bottom, bg=t0.panel)
        recap_btn_fr.grid(row=0, column=1, sticky="ns", padx=(4, 8), pady=8)

        recap_send_btn = tk.Button(
            recap_btn_fr,
            text=tr("recap.send"),
            bg=t0.accent,
            fg="white",
            activebackground=t0.hover_primary,
            activeforeground="white",
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=14,
            pady=8,
            cursor="hand2",
        )
        recap_send_btn.pack(fill="x", pady=(0, 6))
        recap_new_btn = tk.Button(
            recap_btn_fr,
            text=tr("recap.new_chat"),
            bg=_tb,
            fg=_tf,
            activebackground=_tab,
            activeforeground=_taf,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=6,
            cursor="hand2",
        )
        recap_new_btn.pack(fill="x")

        def _recap_send_rest_style() -> Tuple[str, str, str, str, str]:
            t = th()
            if str(recap_send_btn.cget("state")) != "normal":
                ds, gb, gf, dab, daf = t.gen_bind_disabled()
                return ds, gb, gf, dab, daf
            return ("normal", t.accent, "white", t.hover_primary, "white")

        bind_button_hover_if_enabled(
            recap_img_btn,
            lambda: th().toolbar_bind_rest(),
            lambda: th().toolbar_hover()[0],
            lambda: th().toolbar_hover()[1],
        )
        bind_button_hover_if_enabled(
            recap_file_btn,
            lambda: th().toolbar_bind_rest(),
            lambda: th().toolbar_hover()[0],
            lambda: th().toolbar_hover()[1],
        )
        bind_button_hover_if_enabled(
            recap_new_btn,
            lambda: th().toolbar_bind_rest(),
            lambda: th().toolbar_hover()[0],
            lambda: th().toolbar_hover()[1],
        )
        bind_button_hover_if_enabled(
            recap_send_btn,
            _recap_send_rest_style,
            lambda: th().hover_primary,
            lambda: "white",
        )

        _AI_SEND_SPIN = ("-", "/", "|", "\\")
        recap_send_spin: Dict[str, Any] = {"after_id": None, "i": 0}

        def _stop_recap_send_spinner() -> None:
            aid = recap_send_spin.get("after_id")
            if aid is not None:
                try:
                    root.after_cancel(aid)
                except (tk.TclError, ValueError):
                    pass
                recap_send_spin["after_id"] = None

        def _start_recap_send_spinner() -> None:
            _stop_recap_send_spinner()

            def _tick() -> None:
                if not recap_session.get("busy"):
                    recap_send_spin["after_id"] = None
                    return
                try:
                    i = recap_send_spin["i"] % len(_AI_SEND_SPIN)
                    recap_send_btn.config(text=tr("ai.send_busy_prefix") + _AI_SEND_SPIN[i])
                    recap_send_spin["i"] = recap_send_spin["i"] + 1
                    recap_send_spin["after_id"] = root.after(130, _tick)
                except tk.TclError:
                    recap_send_spin["after_id"] = None

            recap_send_spin["i"] = 0
            _tick()

        def recap_set_sending(sending: bool) -> None:
            recap_session["busy"] = sending
            st = "disabled" if sending else "normal"
            recap_send_btn.config(state=st)
            recap_new_btn.config(state=st)
            recap_img_btn.config(state=st)
            recap_file_btn.config(state=st)
            recap_input.config(state=st)
            recap_thinking_chk.config(state=st)
            if sending:
                _start_recap_send_spinner()
            else:
                _stop_recap_send_spinner()
                try:
                    recap_send_btn.config(text=tr("recap.send"))
                except tk.TclError:
                    pass
            recap_refresh_date_controls()

        def reset_recap_session(*_a: Any) -> None:
            recap_session["messages"].clear()
            recap_session["bootstrapped"] = False
            recap_session["busy"] = False
            recap_all_journal_var.set(False)
            recap_pending_images.clear()
            recap_pending_files.clear()
            recap_refresh_pending_lbl()
            recap_transcript.config(state="normal")
            recap_transcript.delete("1.0", "end")
            recap_transcript.config(state="disabled")
            recap_input.delete("1.0", "end")
            recap_selected_dates.clear()
            recap_update_sel_label()
            recap_refresh_cal_marks()
            recap_set_sending(False)

        def reset_recap_on_page_leave() -> None:
            reset_recap_session()
            recap_to_var.set(False)
            try:
                td = datetime.now().date()
                if recap_from_de is not None:
                    recap_from_de.set_date(td)
                if recap_to_de is not None:
                    recap_to_de.set_date(td)
            except Exception:
                pass
            recap_sync_to_checkbox()
            on_recap_to_mode()

        page_leave_reset_handlers["ai_recap"] = reset_recap_on_page_leave

        def recap_new_chat() -> None:
            if recap_session.get("busy"):
                return
            reset_recap_session()

        recap_new_btn.config(command=recap_new_chat)

        def recap_build_context() -> Optional[str]:
            if recap_all_journal_var.get():
                return build_journal_context()
            if recap_to_var.get():
                if DateEntry is None or recap_from_de is None or recap_to_de is None:
                    messagebox.showerror(tr("msg.ai_recap"), tr("recap.err.tkcal_range"))
                    return None
                try:
                    d0 = recap_from_de.get_date()
                    d1 = recap_to_de.get_date()
                except Exception as exc:
                    messagebox.showerror(tr("msg.ai_recap"), tr("recap.err.read_range", err=str(exc)))
                    return None
                start = datetime.combine(d0, datetime.min.time())
                end = datetime.combine(d1, datetime.min.time())
                if end < start:
                    start, end = end, start
                return build_journal_context_for_range((start, end))
            if recap_selected_dates:
                return build_journal_context_for_date_set(recap_selected_dates)
            if DateEntry is None or recap_from_de is None:
                messagebox.showerror(tr("msg.ai_recap"), tr("recap.err.from_tkcal"))
                return None
            try:
                only = recap_from_de.get_date()
            except Exception as exc:
                messagebox.showerror(tr("msg.ai_recap"), tr("recap.err.read_from", err=str(exc)))
                return None
            return build_journal_context_for_date_set({only})

        def recap_send() -> None:
            if recap_session["busy"]:
                return
            if not get_openai_api_key():
                go_settings = messagebox.askyesno(
                    tr("msg.no_api_key_use_ai_title"),
                    tr("msg.no_api_key_use_ai_body"),
                )
                if go_settings:
                    goto_tok = api_key_prompt_hooks.get("goto_token")
                    if callable(goto_tok):
                        goto_tok()
                return
            text = recap_input.get("1.0", "end-1c").strip()
            if not text and not recap_pending_images and not recap_pending_files:
                return
            ctx = recap_build_context()
            if ctx is None:
                return
            model_name = OPENAI_THINKING_MODEL if recap_thinking_var.get() else OPENAI_MODEL
            effort = "high" if recap_thinking_var.get() else None
            imgs = list(recap_pending_images)
            files = list(recap_pending_files)
            recap_set_sending(True)

            def kickoff() -> None:
                try:
                    if not recap_session["bootstrapped"]:
                        recap_session["messages"] = [
                            {
                                "role": "system",
                                "content": (
                                    "You answer questions only using the user's journal context. "
                                    "If the answer is not in the journal, say you do not know based on the journal."
                                ),
                            },
                            {"role": "system", "content": f"Journal context:\n{ctx}"},
                        ]
                        recap_session["bootstrapped"] = True
                    user_msg = build_user_message_with_attachments(text, imgs, files)
                    recap_session["messages"].append(user_msg)
                    answer = chat_completion(
                        recap_session["messages"],
                        model=model_name,
                        reasoning_effort=effort,
                    )

                    def done() -> None:
                        recap_set_sending(False)
                        if _is_likely_api_error_message(answer):
                            messagebox.showerror(tr("msg.ai_recap"), answer[:4000])
                            if recap_session["messages"] and recap_session["messages"][-1].get("role") == "user":
                                recap_session["messages"].pop()
                            return
                        recap_session["messages"].append({"role": "assistant", "content": answer})
                        _append_transcript(recap_transcript, "user", text or tr("chat.attachment_only"))
                        _append_transcript(recap_transcript, "assistant", answer)
                        recap_input.delete("1.0", "end")
                        recap_pending_images.clear()
                        recap_pending_files.clear()
                        recap_refresh_pending_lbl()

                    root.after(0, done)
                except Exception as exc:

                    def fail() -> None:
                        recap_set_sending(False)
                        messagebox.showerror(tr("msg.ai_recap"), str(exc))
                        if recap_session["messages"] and recap_session["messages"][-1].get("role") == "user":
                            recap_session["messages"].pop()

                    root.after(0, fail)

            threading.Thread(target=kickoff, daemon=True).start()

        recap_send_btn.config(command=recap_send)

        def recap_on_enter_key(event: Any) -> Optional[str]:
            # Text widget: <Return> alone may not consume the key; use KeyPress-Return / KP_Enter.
            if (getattr(event, "state", 0) or 0) & 0x0001:
                return None
            if (getattr(event, "state", 0) or 0) & 0x0004:
                return None
            recap_send()
            return "break"

        recap_input.bind("<KeyPress-Return>", recap_on_enter_key, add="+")
        recap_input.bind("<KeyPress-KP_Enter>", recap_on_enter_key, add="+")

        # ========== Chatbot ==========
        cb_wrap = tk.Frame(chatbot_page, bg=t0.surface)
        cb_wrap.pack(
            fill="both",
            expand=True,
            padx=t0.pad_outer,
            pady=(0, t0.pad_center_y + JOURNAL_WINDOW_CONSOLE_RESERVE_BOTTOM),
        )
        cb_wrap.grid_columnconfigure(0, weight=1)
        cb_wrap.grid_rowconfigure(1, weight=1)

        cb_title = tk.Label(
            cb_wrap,
            text=tr("chatbot.title"),
            bg=t0.surface,
            fg=t0.text,
            font=("Segoe UI", 15, "bold"),
            anchor="w",
        )
        cb_title.grid(row=0, column=0, sticky="w", pady=(0, 8))

        cb_top = tk.Frame(cb_wrap, bg=t0.panel, highlightthickness=1, highlightbackground=t0.border)
        cb_top.grid(row=1, column=0, sticky="nsew", pady=(0, 8))
        cb_top.grid_rowconfigure(0, weight=1)
        cb_top.grid_columnconfigure(0, weight=1)

        cb_transcript = tk.Text(
            cb_top,
            wrap="word",
            state="disabled",
            font=("Segoe UI", 10),
            bg=t0.field,
            fg=t0.text,
            insertbackground=t0.text,
            relief="flat",
            padx=12,
            pady=12,
            highlightthickness=1,
            highlightbackground=t0.border,
            highlightcolor=t0.accent,
        )
        cb_ts = tk.Scrollbar(
            cb_top,
            command=cb_transcript.yview,
            bg=t0.panel,
            troughcolor=t0.field,
            activebackground=t0.accent,
            bd=0,
            highlightthickness=0,
            width=11,
        )
        cb_transcript.configure(yscrollcommand=cb_ts.set)
        cb_transcript.grid(row=0, column=0, sticky="nsew")
        cb_ts.grid(row=0, column=1, sticky="ns")
        cb_transcript.tag_configure("t_meta", foreground=t0.muted, font=("Segoe UI", 9, "bold"))
        cb_transcript.tag_configure("t_user", foreground=t0.text, font=("Segoe UI", 10))
        cb_transcript.tag_configure("t_bot", foreground=t0.text, font=("Segoe UI", 10))

        cb_attach = tk.Frame(cb_wrap, bg=t0.surface)
        cb_attach.grid(row=2, column=0, sticky="ew", pady=(0, 4))
        cb_thinking_var = tk.BooleanVar(value=False)
        cb_thinking_chk = tk.Checkbutton(
            cb_attach,
            text=tr("chatbot.thinking"),
            variable=cb_thinking_var,
            bg=t0.surface,
            fg=t0.muted,
            activebackground=t0.surface,
            activeforeground=t0.text,
            selectcolor=t0.field,
            font=("Segoe UI", 9),
        )
        cb_thinking_chk.pack(side="left")
        cb_pending_lbl = tk.Label(
            cb_attach,
            text=tr("recap.attachments", what=tr("recap.attachments_none")),
            bg=t0.surface,
            fg=t0.muted,
            font=("Segoe UI", 9),
            anchor="w",
        )
        cb_pending_lbl.pack(side="left", fill="x", expand=True, padx=(12, 0))

        cb_session: Dict[str, Any] = {
            "messages": [{"role": "system", "content": "You are a helpful assistant."}],
            "busy": False,
        }
        cb_pending_images: List[Path] = []
        cb_pending_files: List[Path] = []

        def cb_refresh_pending() -> None:
            bits = []
            if cb_pending_images:
                bits.append(tr("recap.n_images", n=len(cb_pending_images)))
            if cb_pending_files:
                bits.append(tr("recap.n_files", n=len(cb_pending_files)))
            what = ", ".join(bits) if bits else tr("recap.attachments_none")
            cb_pending_lbl.config(text=tr("recap.attachments", what=what))

        def cb_pick_image() -> None:
            p = filedialog.askopenfilename(
                title="Attach image",
                filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.webp"), ("All files", "*.*")],
            )
            if p:
                cb_pending_images.append(Path(p))
                cb_refresh_pending()

        def cb_pick_file() -> None:
            p = filedialog.askopenfilename(title="Attach file", filetypes=[("All files", "*.*")])
            if p:
                cb_pending_files.append(Path(p))
                cb_refresh_pending()

        cb_img_btn = tk.Button(
            cb_attach,
            text=tr("recap.image"),
            command=cb_pick_image,
            bg=_tb,
            fg=_tf,
            activebackground=_tab,
            activeforeground=_taf,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=4,
            cursor="hand2",
        )
        cb_img_btn.pack(side="right", padx=(6, 0))
        cb_file_btn = tk.Button(
            cb_attach,
            text=tr("recap.file"),
            command=cb_pick_file,
            bg=_tb,
            fg=_tf,
            activebackground=_tab,
            activeforeground=_taf,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=4,
            cursor="hand2",
        )
        cb_file_btn.pack(side="right", padx=(6, 0))

        cb_bottom = tk.Frame(cb_wrap, bg=t0.panel, highlightthickness=1, highlightbackground=t0.border)
        cb_bottom.grid(row=3, column=0, sticky="ew")
        cb_bottom.grid_columnconfigure(0, weight=1)

        cb_input = tk.Text(
            cb_bottom,
            height=3,
            wrap="word",
            font=("Segoe UI", 10),
            bg=t0.field,
            fg=t0.text,
            insertbackground=t0.text,
            relief="flat",
            padx=10,
            pady=8,
            highlightthickness=0,
        )
        cb_input.grid(row=0, column=0, sticky="ew", padx=(8, 4), pady=8)

        cb_btn_fr = tk.Frame(cb_bottom, bg=t0.panel)
        cb_btn_fr.grid(row=0, column=1, sticky="ns", padx=(4, 8), pady=8)

        cb_send_btn = tk.Button(
            cb_btn_fr,
            text=tr("recap.send"),
            bg=t0.accent,
            fg="white",
            activebackground=t0.hover_primary,
            activeforeground="white",
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=14,
            pady=8,
            cursor="hand2",
        )
        cb_send_btn.pack(fill="x", pady=(0, 6))
        cb_new_btn = tk.Button(
            cb_btn_fr,
            text=tr("recap.new_chat"),
            bg=_tb,
            fg=_tf,
            activebackground=_tab,
            activeforeground=_taf,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=6,
            cursor="hand2",
        )
        cb_new_btn.pack(fill="x")

        def _cb_send_rest_style() -> Tuple[str, str, str, str, str]:
            t = th()
            if str(cb_send_btn.cget("state")) != "normal":
                ds, gb, gf, dab, daf = t.gen_bind_disabled()
                return ds, gb, gf, dab, daf
            return ("normal", t.accent, "white", t.hover_primary, "white")

        bind_button_hover_if_enabled(
            cb_img_btn,
            lambda: th().toolbar_bind_rest(),
            lambda: th().toolbar_hover()[0],
            lambda: th().toolbar_hover()[1],
        )
        bind_button_hover_if_enabled(
            cb_file_btn,
            lambda: th().toolbar_bind_rest(),
            lambda: th().toolbar_hover()[0],
            lambda: th().toolbar_hover()[1],
        )
        bind_button_hover_if_enabled(
            cb_new_btn,
            lambda: th().toolbar_bind_rest(),
            lambda: th().toolbar_hover()[0],
            lambda: th().toolbar_hover()[1],
        )
        bind_button_hover_if_enabled(
            cb_send_btn,
            _cb_send_rest_style,
            lambda: th().hover_primary,
            lambda: "white",
        )

        cb_send_spin: Dict[str, Any] = {"after_id": None, "i": 0}

        def _stop_cb_send_spinner() -> None:
            aid = cb_send_spin.get("after_id")
            if aid is not None:
                try:
                    root.after_cancel(aid)
                except (tk.TclError, ValueError):
                    pass
                cb_send_spin["after_id"] = None

        def _start_cb_send_spinner() -> None:
            _stop_cb_send_spinner()

            def _tick() -> None:
                if not cb_session.get("busy"):
                    cb_send_spin["after_id"] = None
                    return
                try:
                    i = cb_send_spin["i"] % len(_AI_SEND_SPIN)
                    cb_send_btn.config(text=tr("ai.send_busy_prefix") + _AI_SEND_SPIN[i])
                    cb_send_spin["i"] = cb_send_spin["i"] + 1
                    cb_send_spin["after_id"] = root.after(130, _tick)
                except tk.TclError:
                    cb_send_spin["after_id"] = None

            cb_send_spin["i"] = 0
            _tick()

        def cb_set_sending(sending: bool) -> None:
            cb_session["busy"] = sending
            st = "disabled" if sending else "normal"
            cb_send_btn.config(state=st)
            cb_new_btn.config(state=st)
            cb_img_btn.config(state=st)
            cb_file_btn.config(state=st)
            cb_input.config(state=st)
            cb_thinking_chk.config(state=st)
            if sending:
                _start_cb_send_spinner()
            else:
                _stop_cb_send_spinner()
                try:
                    cb_send_btn.config(text=tr("recap.send"))
                except tk.TclError:
                    pass

        def reset_chatbot_session(*_a: Any) -> None:
            cb_session["messages"] = [{"role": "system", "content": "You are a helpful assistant."}]
            cb_session["busy"] = False
            cb_pending_images.clear()
            cb_pending_files.clear()
            cb_refresh_pending()
            cb_transcript.config(state="normal")
            cb_transcript.delete("1.0", "end")
            cb_transcript.config(state="disabled")
            cb_input.delete("1.0", "end")
            cb_set_sending(False)

        page_leave_reset_handlers["chatbot"] = reset_chatbot_session

        def cb_new_chat() -> None:
            if cb_session.get("busy"):
                return
            reset_chatbot_session()

        cb_new_btn.config(command=cb_new_chat)

        def cb_send() -> None:
            if cb_session["busy"]:
                return
            if not get_openai_api_key():
                if messagebox.askyesno(
                    tr("msg.no_api_key_use_ai_title"),
                    tr("msg.no_api_key_use_ai_body"),
                ):
                    goto_tok = api_key_prompt_hooks.get("goto_token")
                    if callable(goto_tok):
                        goto_tok()
                return
            text = cb_input.get("1.0", "end-1c").strip()
            if not text and not cb_pending_images and not cb_pending_files:
                return
            model_name = OPENAI_THINKING_MODEL if cb_thinking_var.get() else OPENAI_MODEL
            effort = "high" if cb_thinking_var.get() else None
            imgs = list(cb_pending_images)
            files = list(cb_pending_files)
            cb_set_sending(True)

            def kickoff() -> None:
                try:
                    user_msg = build_user_message_with_attachments(text, imgs, files)
                    cb_session["messages"].append(user_msg)
                    answer = chat_completion(
                        cb_session["messages"],
                        model=model_name,
                        reasoning_effort=effort,
                    )

                    def done() -> None:
                        cb_set_sending(False)
                        if _is_likely_api_error_message(answer):
                            messagebox.showerror(tr("msg.chatbot"), answer[:4000])
                            if cb_session["messages"] and cb_session["messages"][-1].get("role") == "user":
                                cb_session["messages"].pop()
                            return
                        cb_session["messages"].append({"role": "assistant", "content": answer})
                        _append_transcript(cb_transcript, "user", text or tr("chat.attachment_only"))
                        _append_transcript(cb_transcript, "assistant", answer)
                        cb_input.delete("1.0", "end")
                        cb_pending_images.clear()
                        cb_pending_files.clear()
                        cb_refresh_pending()

                    root.after(0, done)
                except Exception as exc:

                    def fail() -> None:
                        cb_set_sending(False)
                        messagebox.showerror(tr("msg.chatbot"), str(exc))
                        if cb_session["messages"] and cb_session["messages"][-1].get("role") == "user":
                            cb_session["messages"].pop()

                    root.after(0, fail)

            threading.Thread(target=kickoff, daemon=True).start()

        cb_send_btn.config(command=cb_send)

        def cb_on_enter_key(event: Any) -> Optional[str]:
            if (getattr(event, "state", 0) or 0) & 0x0001:
                return None
            if (getattr(event, "state", 0) or 0) & 0x0004:
                return None
            cb_send()
            return "break"

        cb_input.bind("<KeyPress-Return>", cb_on_enter_key, add="+")
        cb_input.bind("<KeyPress-KP_Enter>", cb_on_enter_key, add="+")

        def apply_ai_recap_chatbot_theme() -> None:
            t = th()
            tb, tf, tab, taf = t.toolbar_btn_config()
            for fr in (
                recap_wrap,
                recap_title,
                recap_top,
                recap_thinking_chk,
                recap_all_journal_chk,
                recap_from_fr,
                recap_to_wrap,
                recap_to_chk,
                recap_through_fr,
                recap_cal_row,
                recap_sel_lbl,
                recap_mid,
                recap_attach_row,
                recap_pending_lbl,
                recap_bottom,
                recap_btn_fr,
                cb_wrap,
                cb_title,
                cb_attach,
                cb_pending_lbl,
                cb_bottom,
                cb_btn_fr,
            ):
                try:
                    fr.configure(bg=t.surface)
                except Exception:
                    try:
                        fr.configure(bg=t.panel)
                    except Exception:
                        pass
            recap_title.configure(bg=t.surface, fg=t.text)
            recap_top.configure(bg=t.panel, highlightbackground=t.border)
            recap_thinking_chk.configure(
                bg=t.panel,
                fg=t.muted,
                activebackground=t.panel,
                activeforeground=t.text,
                selectcolor=t.field,
            )
            recap_all_journal_chk.configure(
                bg=t.panel,
                fg=t.muted,
                activebackground=t.panel,
                activeforeground=t.text,
                selectcolor=t.field,
            )
            recap_to_wrap.configure(bg=t.panel)
            recap_to_chk.configure(
                bg=t.panel,
                fg=t.muted,
                activebackground=t.panel,
                activeforeground=t.text,
                selectcolor=t.field,
            )
            recap_from_fr.configure(bg=t.panel)
            for _w in recap_from_fr.winfo_children():
                if isinstance(_w, tk.Label):
                    _w.configure(bg=t.panel, fg=t.muted)
            recap_through_fr.configure(bg=t.panel)
            for _w in recap_through_fr.winfo_children():
                if isinstance(_w, tk.Label):
                    _w.configure(bg=t.panel, fg=t.muted)
            recap_cal_row.configure(bg=t.surface)
            recap_sel_lbl.configure(bg=t.surface, fg=t.muted)
            recap_mid.configure(bg=t.surface)
            recap_transcript.config(
                bg=t.field,
                fg=t.text,
                insertbackground=t.text,
                highlightbackground=t.border,
                highlightcolor=t.accent,
            )
            recap_ts.config(bg=t.panel, troughcolor=t.field, activebackground=t.accent)
            recap_transcript.tag_configure("t_meta", foreground=t.muted)
            recap_transcript.tag_configure("t_user", foreground=t.text)
            recap_transcript.tag_configure("t_bot", foreground=t.text)
            recap_attach_row.configure(bg=t.surface)
            recap_pending_lbl.configure(bg=t.surface, fg=t.muted)
            recap_bottom.configure(bg=t.panel, highlightbackground=t.border)
            recap_input.config(bg=t.field, fg=t.text, insertbackground=t.text)
            recap_btn_fr.configure(bg=t.panel)
            recap_send_btn.configure(
                bg=t.accent,
                fg="white",
                activebackground=t.hover_primary,
                activeforeground="white",
            )
            recap_new_btn.configure(bg=tb, fg=tf, activebackground=tab, activeforeground=taf)
            recap_img_btn.configure(bg=tb, fg=tf, activebackground=tab, activeforeground=taf)
            recap_file_btn.configure(bg=tb, fg=tf, activebackground=tab, activeforeground=taf)
            if recap_calendar is not None:
                try:
                    recap_calendar.config(
                        background=t.field,
                        foreground=t.text,
                        headersbackground=t.panel,
                        headersforeground=t.text,
                        weekendbackground=t.field,
                        weekendforeground=t.muted,
                        normalbackground=t.field,
                        normalforeground=t.text,
                        othermonthbackground=t.field,
                        othermonthforeground=t.muted,
                        selectbackground=t.accent,
                        selectforeground="white",
                        bordercolor=t.border,
                    )
                except Exception:
                    pass
                recap_refresh_cal_marks()
            if DateEntry is not None and recap_from_de is not None:
                try:
                    recap_from_de.config(background=t.field, foreground=t.text)
                except tk.TclError:
                    pass
            if DateEntry is not None and recap_to_de is not None:
                try:
                    recap_to_de.config(background=t.field, foreground=t.text)
                except tk.TclError:
                    pass
            cb_wrap.configure(bg=t.surface)
            cb_title.configure(bg=t.surface, fg=t.text)
            cb_top.configure(bg=t.panel, highlightbackground=t.border)
            cb_transcript.config(
                bg=t.field,
                fg=t.text,
                insertbackground=t.text,
                highlightbackground=t.border,
                highlightcolor=t.accent,
            )
            cb_ts.config(bg=t.panel, troughcolor=t.field, activebackground=t.accent)
            cb_transcript.tag_configure("t_meta", foreground=t.muted)
            cb_transcript.tag_configure("t_user", foreground=t.text)
            cb_transcript.tag_configure("t_bot", foreground=t.text)
            cb_attach.configure(bg=t.surface)
            cb_thinking_chk.configure(
                bg=t.surface,
                fg=t.muted,
                activebackground=t.surface,
                activeforeground=t.text,
                selectcolor=t.field,
            )
            cb_pending_lbl.configure(bg=t.surface, fg=t.muted)
            cb_bottom.configure(bg=t.panel, highlightbackground=t.border)
            cb_input.config(bg=t.field, fg=t.text, insertbackground=t.text)
            cb_btn_fr.configure(bg=t.panel)
            cb_send_btn.configure(
                bg=t.accent,
                fg="white",
                activebackground=t.hover_primary,
                activeforeground="white",
            )
            cb_new_btn.configure(bg=tb, fg=tf, activebackground=tab, activeforeground=taf)
            cb_img_btn.configure(bg=tb, fg=tf, activebackground=tab, activeforeground=taf)
            cb_file_btn.configure(bg=tb, fg=tf, activebackground=tab, activeforeground=taf)

        def refresh_recap_chat_i18n() -> None:
            try:
                _has_key = bool(get_openai_api_key())
                recap_title.config(
                    text=tr("recap.title" if _has_key else "recap.title_no_key")
                )
                recap_thinking_chk.config(text=tr("recap.thinking"))
                recap_all_journal_chk.config(text=tr("recap.all_journal"))
                recap_from_lbl.config(text=tr("recap.from"))
                recap_to_chk.config(text=tr("recap.to_chk"))
                recap_through_lbl.config(text=tr("recap.through"))
                recap_img_btn.config(text=tr("recap.image"))
                recap_file_btn.config(text=tr("recap.file"))
                if not recap_session.get("busy"):
                    recap_send_btn.config(text=tr("recap.send"))
                recap_new_btn.config(text=tr("recap.new_chat"))
                recap_update_sel_label()
                recap_refresh_pending_lbl()
                cb_title.config(
                    text=tr("chatbot.title" if _has_key else "chatbot.title_no_key")
                )
                cb_thinking_chk.config(text=tr("chatbot.thinking"))
                cb_img_btn.config(text=tr("recap.image"))
                cb_file_btn.config(text=tr("recap.file"))
                if not cb_session.get("busy"):
                    cb_send_btn.config(text=tr("recap.send"))
                cb_new_btn.config(text=tr("recap.new_chat"))
                cb_refresh_pending()
            except tk.TclError:
                pass

        build_ai_recap_and_chatbot_pages._i18n = refresh_recap_chat_i18n  # type: ignore[attr-defined]
        build_ai_recap_and_chatbot_pages._apply_theme = apply_ai_recap_chatbot_theme  # type: ignore[attr-defined]

    build_ai_recap_and_chatbot_pages()
    settings_wrap = tk.Frame(settings_page, bg=t_init.surface)
    settings_wrap.pack(fill="both", expand=True, padx=20, pady=20)
    _register_page_toggle(settings_page)
    settings_title = tk.Label(
        settings_wrap,
        text=tr("settings.title"),
        bg=t_init.surface,
        fg=t_init.text,
        font=("Segoe UI", 16, "bold"),
        anchor="w",
    )
    settings_title.pack(anchor="w", pady=(0, 12))
    settings_status_var = tk.StringVar(value="")
    settings_status_lbl = tk.Label(
        settings_wrap,
        textvariable=settings_status_var,
        bg=t_init.surface,
        fg=t_init.muted,
        font=("Segoe UI", 9),
        anchor="w",
        justify="left",
    )
    settings_status_lbl.pack(fill="x", pady=(0, 10))

    settings_rows: List[Any] = []
    settings_labels: List[Any] = []
    settings_label_keys: List[Tuple[Any, str]] = []

    def _make_settings_row(label_key: str) -> Tuple[Any, Any]:
        row = tk.Frame(settings_wrap, bg=t_init.surface)
        row.pack(fill="x", pady=(0, 10))
        lbl = tk.Label(
            row,
            text=tr(label_key),
            bg=t_init.surface,
            fg=t_init.muted,
            font=("Segoe UI", 10, "bold"),
            width=18,
            anchor="w",
        )
        lbl.pack(side="left")
        settings_rows.append(row)
        settings_labels.append(lbl)
        settings_label_keys.append((lbl, label_key))
        return row, lbl

    settings_prefs = load_preferences()
    settings_app_name = {"value": settings_prefs.get("app_name", "Daily Logger") or "Daily Logger"}
    console_hint_state: Dict[str, Any] = {"text": "", "apply": None, "reset_after_id": None}

    def _clear_console_hint() -> None:
        console_hint_state["text"] = ""
        _id = console_hint_state.get("reset_after_id")
        if _id is not None:
            try:
                root.after_cancel(_id)
            except Exception:
                pass
            console_hint_state["reset_after_id"] = None
        apply_hint = console_hint_state.get("apply")
        if callable(apply_hint):
            apply_hint()

    def _set_settings_status(msg: str) -> None:
        settings_status_var.set("")
        console_hint_state["text"] = msg.strip()
        _id = console_hint_state.get("reset_after_id")
        if _id is not None:
            try:
                root.after_cancel(_id)
            except Exception:
                pass
        console_hint_state["reset_after_id"] = root.after(10000, _clear_console_hint)
        apply_hint = console_hint_state.get("apply")
        if callable(apply_hint):
            apply_hint()

    lang_row, _ = _make_settings_row("settings.language")
    ui_lang_var = tk.StringVar(
        value=tr("settings.lang.chinese")
        if ui_lang_holder[0] == "zh"
        else tr("settings.lang.english")
    )
    lang_ui_combo: Any = None
    if ttk is not None:
        lang_ui_combo = ttk.Combobox(
            lang_row,
            textvariable=ui_lang_var,
            values=(tr("settings.lang.english"), tr("settings.lang.chinese")),
            state="readonly",
            width=14,
            style="Journal.TCombobox",
        )
        lang_ui_combo.pack(side="left", fill="x", expand=True, padx=(0, 8))
    else:
        lang_ui_combo = tk.OptionMenu(
            lang_row,
            ui_lang_var,
            tr("settings.lang.english"),
            tr("settings.lang.chinese"),
        )
        lang_ui_combo.pack(side="left", fill="x", expand=True, padx=(0, 8))

    def _on_ui_language_selected(_evt: object | None = None) -> None:
        raw = ui_lang_var.get().strip()
        new_lang = "zh" if raw == tr("settings.lang.chinese") else "en"
        if new_lang == ui_lang_holder[0]:
            return
        prefs = load_preferences()
        prefs[UI_LANGUAGE_PREF_KEY] = new_lang
        save_preferences(prefs)
        ui_lang_holder[0] = new_lang
        apply_journal_window_colors()

    rename_row, _ = _make_settings_row("settings.rename")
    rename_entry = tk.Entry(
        rename_row,
        bg=t_init.field,
        fg=t_init.text,
        insertbackground=t_init.text,
        relief="flat",
        highlightthickness=1,
        highlightbackground=t_init.border,
        highlightcolor=t_init.accent,
        font=("Segoe UI", 10),
    )
    rename_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
    rename_entry.insert(0, settings_app_name["value"])
    rename_btn = tk.Button(
        rename_row,
        text=tr("settings.rename_btn"),
        bg=t_init.btn_secondary,
        fg=t_init.text,
        activebackground=t_init.secondary_hover,
        activeforeground=t_init.text,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=10,
        pady=6,
        cursor="hand2",
    )
    rename_btn.pack(side="left")

    startup_row, _ = _make_settings_row("settings.startup")
    startup_state = {"enabled": is_startup_enabled()}
    startup_toggle_btn = tk.Button(
        startup_row,
        text=tr("settings.on") if startup_state["enabled"] else tr("settings.off"),
        bg=t_init.btn_secondary,
        fg=t_init.text,
        activebackground=t_init.secondary_hover,
        activeforeground=t_init.text,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=12,
        pady=6,
        cursor="hand2",
        width=7,
    )
    startup_toggle_btn.pack(side="left")

    theme_row, _ = _make_settings_row("settings.theme")
    settings_theme_btn = tk.Button(
        theme_row,
        text=t_init.toggle_label,
        command=lambda: toggle_journal_window_theme(),
        bg=t_init.btn_secondary,
        fg=t_init.text,
        activebackground=t_init.secondary_hover,
        activeforeground=t_init.text,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=12,
        pady=6,
        cursor="hand2",
    )
    settings_theme_btn.pack(side="left")

    def _backup_mode_btn_label(mode_val: str) -> str:
        return tr(
            {"On": "backup.on", "Off": "backup.off", "Limited": "backup.limited"}.get(
                mode_val, "backup.off"
            )
        )

    backup_row, _ = _make_settings_row("settings.backup")
    backup_mode = {"value": "On"}
    if _is_pref_true(settings_prefs.get("backup_limited", "false")):
        backup_mode["value"] = "Limited"
    elif not _is_pref_true(settings_prefs.get("backup_enabled", "true")):
        backup_mode["value"] = "Off"
    backup_mode_btn = tk.Button(
        backup_row,
        text=_backup_mode_btn_label(backup_mode["value"]),
        bg=t_init.btn_secondary,
        fg=t_init.text,
        activebackground=t_init.secondary_hover,
        activeforeground=t_init.text,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=12,
        pady=6,
        cursor="hand2",
        width=9,
    )
    backup_mode_btn.pack(side="left", padx=(0, 8))
    backup_manual_btn = tk.Button(
        backup_row,
        text=tr("settings.manual"),
        bg=t_init.btn_secondary,
        fg=t_init.text,
        activebackground=t_init.secondary_hover,
        activeforeground=t_init.text,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=12,
        pady=6,
        cursor="hand2",
        width=9,
    )
    backup_manual_btn.pack(side="left")
    bind_hover_tooltip(
        backup_mode_btn,
        lambda: tr("tip.backup_mode"),
    )
    bind_hover_tooltip(
        backup_manual_btn,
        lambda: tr("tip.backup_manual"),
    )

    token_row, _ = _make_settings_row("settings.token")
    token_saved = {"value": get_openai_api_key() or ""}
    token_entry = tk.Entry(
        token_row,
        bg=t_init.field,
        fg=t_init.text,
        insertbackground=t_init.text,
        relief="flat",
        highlightthickness=1,
        highlightbackground=t_init.border,
        highlightcolor=t_init.accent,
        font=("Consolas", 10),
    )
    token_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
    if token_saved["value"]:
        token_entry.insert(0, "*" * max(32, len(token_saved["value"])))
    token_save_btn = tk.Button(
        token_row,
        text=tr("settings.save"),
        bg=t_init.btn_secondary,
        fg=t_init.text,
        activebackground=t_init.secondary_hover,
        activeforeground=t_init.text,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=10,
        pady=6,
        cursor="hand2",
        width=7,
    )
    token_save_btn.pack(side="left", padx=(0, 8))
    token_copy_btn = tk.Button(
        token_row,
        text=tr("settings.copy"),
        bg=t_init.btn_secondary,
        fg=t_init.text,
        activebackground=t_init.secondary_hover,
        activeforeground=t_init.text,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=10,
        pady=6,
        cursor="hand2",
        width=7,
    )
    token_copy_btn.pack(side="left")

    start_menu_row, _ = _make_settings_row("settings.start_menu")
    start_menu_app_btn = tk.Button(
        start_menu_row,
        bg=t_init.btn_secondary,
        fg=t_init.text,
        activebackground=t_init.secondary_hover,
        activeforeground=t_init.text,
        relief="flat",
        text=tr("settings.start_menu_app"),
        font=("Segoe UI", 9, "bold"),
        padx=12,
        pady=6,
        cursor="hand2",
    )
    start_menu_app_btn.pack(side="left", padx=(0, 8))
    start_menu_journal_btn = tk.Button(
        start_menu_row,
        text=tr("settings.start_menu_journal"),
        bg=t_init.btn_secondary,
        fg=t_init.text,
        activebackground=t_init.secondary_hover,
        activeforeground=t_init.text,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=12,
        pady=6,
        cursor="hand2",
    )
    start_menu_journal_btn.pack(side="left")
    bind_hover_tooltip(
        start_menu_app_btn,
        lambda: tr("tip.start_menu_app"),
    )
    bind_hover_tooltip(
        start_menu_journal_btn,
        lambda: tr("tip.start_menu_journal"),
    )

    def _refresh_token_entry_mask() -> None:
        token_entry.delete(0, "end")
        if token_saved["value"]:
            token_entry.insert(0, "*" * max(32, len(token_saved["value"])))

    def _is_token_mask(value: str) -> bool:
        return bool(value) and all(ch == "*" for ch in value)

    def _on_rename_apply() -> None:
        new_name = rename_entry.get().strip() or "Daily Logger"
        updated = rename_app_name_to(new_name)
        settings_app_name["value"] = updated
        console_app_name["value"] = updated
        root.title(updated)
        nav_title.config(text=updated)
        rename_entry.delete(0, "end")
        rename_entry.insert(0, updated)
        _set_settings_status(tr("status.rename_ok", name=updated))

    def _on_toggle_startup() -> None:
        should_enable = not startup_state["enabled"]
        ok = create_startup_shortcut() if should_enable else remove_startup_shortcut()
        if not ok:
            _set_settings_status(tr("status.startup_fail"))
            return
        startup_state["enabled"] = should_enable
        startup_toggle_btn.config(
            text=tr("settings.on") if should_enable else tr("settings.off")
        )
        prefs = load_preferences()
        prefs["startup_enabled"] = "true" if should_enable else "false"
        save_preferences(prefs)
        _set_settings_status(
            tr("status.startup_on") if should_enable else tr("status.startup_off")
        )

    def _persist_backup_mode(mode: str) -> None:
        prefs = load_preferences()
        if mode == "On":
            prefs["backup_enabled"] = "true"
            prefs["backup_limited"] = "false"
        elif mode == "Off":
            prefs["backup_enabled"] = "false"
            prefs["backup_limited"] = "false"
        else:
            prefs["backup_enabled"] = "true"
            prefs["backup_limited"] = "true"
        if save_preferences(prefs):
            _set_settings_status(tr("status.backup_mode", mode=_backup_mode_btn_label(mode)))
        else:
            _set_settings_status(tr("status.backup_save_fail"))

    def _on_cycle_backup_mode() -> None:
        order = ("On", "Off", "Limited")
        idx = order.index(backup_mode["value"])
        next_mode = order[(idx + 1) % len(order)]
        backup_mode["value"] = next_mode
        backup_mode_btn.config(text=_backup_mode_btn_label(next_mode))
        _persist_backup_mode(next_mode)

    def _on_manual_backup() -> None:
        prefs = load_preferences()
        evict_oldest_backup_if_limited_full(prefs)
        backup_path = run_backup_now()
        if backup_path is None:
            _set_settings_status(tr("status.backup_skip"))
            return
        trim_backups_if_limited(prefs)
        _set_settings_status(tr("status.backup_ok", name=backup_path.name))

    def _on_token_focus_in(_evt: Optional[Any] = None) -> None:
        if _is_token_mask(token_entry.get()):
            token_entry.delete(0, "end")

    def _on_token_save() -> None:
        typed = token_entry.get().strip()
        if _is_token_mask(typed):
            _set_settings_status(tr("status.token_same"))
            return
        if not typed:
            if delete_openai_api_key():
                token_saved["value"] = ""
                _refresh_token_entry_mask()
                _set_settings_status(tr("status.token_removed"))
                _ai_i18n = getattr(build_ai_recap_and_chatbot_pages, "_i18n", None)
                if callable(_ai_i18n):
                    _ai_i18n()
            else:
                _set_settings_status(tr("status.token_remove_fail"))
            return
        if save_openai_api_key(typed):
            token_saved["value"] = typed
            _refresh_token_entry_mask()
            _set_settings_status(tr("status.token_saved"))
            _ai_i18n = getattr(build_ai_recap_and_chatbot_pages, "_i18n", None)
            if callable(_ai_i18n):
                _ai_i18n()
        else:
            _set_settings_status(tr("status.token_save_fail"))

    def _on_token_copy() -> None:
        current = get_openai_api_key() or ""
        if not current:
            _set_settings_status(tr("status.token_no_copy"))
            return
        if copy_text_to_clipboard(current):
            _set_settings_status(tr("status.token_copied"))
        else:
            _set_settings_status(tr("status.token_copy_fail"))

    def _on_start_menu_button(selected: str) -> None:
        if selected == "journal":
            ok = sb_create_journal_search_shortcut()
        else:
            ok = sb_create_bat_search_shortcut()
        if ok:
            _set_settings_status(tr("status.start_menu_ok", which=selected))
        else:
            _set_settings_status(tr("status.start_menu_fail", which=selected))

    rename_btn.config(command=_on_rename_apply)
    startup_toggle_btn.config(command=_on_toggle_startup)
    backup_mode_btn.config(command=_on_cycle_backup_mode)
    backup_manual_btn.config(command=_on_manual_backup)
    token_save_btn.config(command=_on_token_save)
    token_copy_btn.config(command=_on_token_copy)
    start_menu_app_btn.config(command=lambda: _on_start_menu_button("app"))
    start_menu_journal_btn.config(command=lambda: _on_start_menu_button("journal"))
    token_entry.bind("<FocusIn>", _on_token_focus_in, add="+")

    def _goto_settings_token_field() -> None:
        show_page("settings")

        def _focus_token() -> None:
            try:
                token_entry.focus_set()
            except tk.TclError:
                return
            try:
                token_entry.selection_range(0, "end")
            except tk.TclError:
                pass

        root.after(100, _focus_token)

    api_key_prompt_hooks["goto_token"] = _goto_settings_token_field

    for _btn in (
        rename_btn,
        startup_toggle_btn,
        settings_theme_btn,
        backup_mode_btn,
        backup_manual_btn,
        token_save_btn,
        token_copy_btn,
        start_menu_app_btn,
        start_menu_journal_btn,
    ):
        bind_button_hover_if_enabled(
            _btn,
            lambda b=_btn: (
                str(b.cget("state")),
                th().btn_secondary,
                th().text,
                th().secondary_hover,
                th().text,
            ),
            lambda: th().secondary_hover,
            lambda: th().text,
        )

    console_wrap = tk.Frame(console_page, bg=t_init.surface)
    console_wrap.pack(fill="both", expand=True, padx=20, pady=20)
    _register_page_toggle(console_page)
    console_title = tk.Label(
        console_wrap,
        text="Console",
        bg=t_init.surface,
        fg=t_init.text,
        font=("Segoe UI", 16, "bold"),
        anchor="w",
    )
    console_title.pack(anchor="w", pady=(0, 8))
    console_output = tk.Text(
        console_wrap,
        wrap="word",
        height=20,
        bg=t_init.field,
        fg=t_init.text,
        insertbackground=t_init.text,
        relief="flat",
        padx=12,
        pady=12,
        font=("Consolas", 10),
        state="disabled",
        highlightthickness=1,
        highlightbackground=t_init.border,
        highlightcolor=t_init.accent,
    )
    console_output.pack(fill="both", expand=True, side="left")
    console_scroll = tk.Scrollbar(
        console_wrap,
        command=console_output.yview,
        bg=t_init.panel,
        troughcolor=t_init.field,
        activebackground=t_init.accent,
        bd=0,
        highlightthickness=0,
        width=11,
    )
    console_scroll.pack(fill="y", side="right")
    console_output.configure(yscrollcommand=console_scroll.set)

    console_input_row = tk.Frame(content_host, bg=t_init.surface)
    console_input_row.place_forget()
    console_input_holder["row"] = console_input_row
    console_prompt = tk.Label(
        console_input_row,
        text="> ",
        bg=t_init.surface,
        fg=t_init.muted,
        font=("Consolas", 11, "bold"),
        cursor="hand2",
    )
    console_prompt.pack(side="left")
    console_entry = tk.Entry(
        console_input_row,
        bg=t_init.field,
        fg=t_init.text,
        insertbackground=t_init.text,
        relief="flat",
        highlightthickness=1,
        highlightbackground=t_init.border,
        highlightcolor=t_init.accent,
        font=("Consolas", 10),
        width=1,
    )
    console_entry.pack(side="left", fill="x", expand=True, padx=(0, 0))
    console_insertwidth_normal = int(console_entry.cget("insertwidth") or 1)
    console_entry_state: Dict[str, bool] = {"placeholder": False}

    def _set_console_placeholder() -> None:
        if console_entry.get():
            return
        console_entry_state["placeholder"] = True
        console_entry.config(
            fg=t_init.muted,
            font=("Consolas", 10, "italic"),
            insertwidth=0,
        )
        hint = str(console_hint_state.get("text", "")).strip()
        console_entry.insert(0, hint or tr("console.placeholder"))

    def _clear_console_placeholder() -> None:
        if not console_entry_state["placeholder"]:
            return
        console_entry.delete(0, "end")
        console_entry_state["placeholder"] = False
        console_entry.config(
            fg=th().text,
            font=("Consolas", 10),
            insertwidth=console_insertwidth_normal,
        )
        console_hint_state["text"] = ""

    def _show_console_hint_placeholder() -> None:
        if console_entry_state["placeholder"] or not console_entry.get().strip():
            console_entry.delete(0, "end")
            console_entry_state["placeholder"] = False
            _set_console_placeholder()

    console_hint_state["apply"] = _show_console_hint_placeholder
    _set_console_placeholder()
    _journal_find_state: Dict[str, Any] = {"widget": text_box}

    def _active_journal_text_widget() -> tk.Text:
        w = root.focus_get()
        if isinstance(w, tk.Text) and w in (text_box, stt_box, report_box):
            _journal_find_state["widget"] = w
            return w
        saved = _journal_find_state.get("widget")
        if isinstance(saved, tk.Text):
            return saved
        return text_box

    def _search_scope_widgets() -> Tuple[tk.Text, ...]:
        if find_scope_var.get() == "one":
            return (_active_journal_text_widget(),)
        return (text_box, stt_box, report_box)

    def _journal_select_all(_evt: Optional[Any] = None) -> str:
        w = _active_journal_text_widget()
        try:
            w.tag_add("sel", "1.0", "end-1c")
            w.mark_set("insert", "end-1c")
            w.see("insert")
        except tk.TclError:
            pass
        return "break"

    def _find_close() -> None:
        find_row.pack_forget()
        find_status.config(text="")
        _active_journal_text_widget().focus_set()

    def _find_all_ranges(w: tk.Text, query: str, case_sensitive: bool, whole_word: bool) -> List[Tuple[str, str]]:
        pattern = query if not whole_word else rf"\m{re.escape(query)}\M"
        ranges: List[Tuple[str, str]] = []
        idx = "1.0"
        while True:
            pos = w.search(
                pattern,
                idx,
                stopindex="end-1c",
                nocase=not case_sensitive,
                regexp=whole_word,
            )
            if not pos:
                break
            end = f"{pos}+{len(query)}c"
            ranges.append((pos, end))
            idx = end
        return ranges

    def _find_update_status_for_selection(
        w: tk.Text, query: str, sel_start: str, sel_widget: Optional[tk.Text] = None
    ) -> None:
        if not query:
            find_status.config(text="")
            return
        widgets = _search_scope_widgets()
        all_ranges: List[Tuple[tk.Text, str, str]] = []
        for _w in widgets:
            _ranges = _find_all_ranges(
                _w,
                query,
                case_sensitive=find_case_var.get(),
                whole_word=find_word_var.get(),
            )
            all_ranges.extend([(_w, _s, _e) for _s, _e in _ranges])
        if not all_ranges:
            find_status.config(text="No matches")
            return
        current = 1
        sw = sel_widget or w
        for i, (rw, s, _e) in enumerate(all_ranges, start=1):
            if rw is sw and s == sel_start:
                current = i
                break
        find_status.config(text=f"{current}/{len(all_ranges)}")

    def _find_next(direction: int = 1) -> str:
        query = find_var.get()
        if not query:
            find_status.config(text="Type text to find")
            return "break"
        widgets = _search_scope_widgets()
        ranges_by_widget: Dict[tk.Text, List[Tuple[str, str]]] = {}
        total_matches = 0
        for _w in widgets:
            _ranges = _find_all_ranges(
                _w,
                query,
                case_sensitive=find_case_var.get(),
                whole_word=find_word_var.get(),
            )
            ranges_by_widget[_w] = _ranges
            total_matches += len(_ranges)
        if total_matches == 0:
            find_status.config(text="No matches")
            return "break"
        w = _active_journal_text_widget()
        if w not in widgets:
            w = widgets[0]
        start = w.index("insert")
        if w.tag_ranges("sel"):
            start = w.index("sel.last") if direction > 0 else w.index("sel.first")
        start_widget_idx = widgets.index(w)

        def _pick_forward() -> Optional[Tuple[tk.Text, str, str]]:
            for wi in range(start_widget_idx, len(widgets)):
                _w = widgets[wi]
                _ranges = ranges_by_widget.get(_w, [])
                if not _ranges:
                    continue
                if wi == start_widget_idx:
                    for s, e in _ranges:
                        if _w.compare(s, ">", start):
                            return (_w, s, e)
                else:
                    return (_w, _ranges[0][0], _ranges[0][1])
            for wi in range(0, start_widget_idx + 1):
                _w = widgets[wi]
                _ranges = ranges_by_widget.get(_w, [])
                if not _ranges:
                    continue
                if wi == start_widget_idx:
                    return (_w, _ranges[0][0], _ranges[0][1])
                return (_w, _ranges[0][0], _ranges[0][1])
            return None

        def _pick_backward() -> Optional[Tuple[tk.Text, str, str]]:
            for wi in range(start_widget_idx, -1, -1):
                _w = widgets[wi]
                _ranges = ranges_by_widget.get(_w, [])
                if not _ranges:
                    continue
                if wi == start_widget_idx:
                    for s, e in reversed(_ranges):
                        if _w.compare(s, "<", start):
                            return (_w, s, e)
                else:
                    s, e = _ranges[-1]
                    return (_w, s, e)
            for wi in range(len(widgets) - 1, start_widget_idx - 1, -1):
                _w = widgets[wi]
                _ranges = ranges_by_widget.get(_w, [])
                if not _ranges:
                    continue
                if wi == start_widget_idx:
                    s, e = _ranges[-1]
                    return (_w, s, e)
                s, e = _ranges[-1]
                return (_w, s, e)
            return None

        picked = _pick_forward() if direction > 0 else _pick_backward()
        if not picked:
            find_status.config(text="No matches")
            return "break"
        pw, pos, end = picked
        for _w in (text_box, stt_box, report_box):
            _w.tag_remove("sel", "1.0", "end")
        pw.tag_add("sel", pos, end)
        pw.mark_set("insert", end if direction > 0 else pos)
        pw.see(pos)
        pw.focus_set()
        _journal_find_state["widget"] = pw
        _find_update_status_for_selection(pw, query, pos, sel_widget=pw)
        return "break"

    def _find_prev(_evt: Optional[Any] = None) -> str:
        return _find_next(-1)

    def _find_open(_evt: Optional[Any] = None) -> str:
        if str(find_row.winfo_manager()) == "pack":
            _find_close()
            return "break"
        if active_page["key"] != "journal":
            show_page("journal")
        w = _active_journal_text_widget()
        find_row.pack(fill="x", padx=t_init.pad_outer, pady=(0, 6), before=center)
        selected = ""
        try:
            selected = w.get("sel.first", "sel.last")
        except tk.TclError:
            selected = ""
        if selected.strip():
            find_var.set(selected)
        _find_update_status_for_selection(w, find_var.get(), "", sel_widget=w)
        find_entry.focus_set()
        find_entry.selection_range(0, "end")
        return "break"

    find_next_btn.config(command=lambda: _find_next(1))
    find_prev_btn.config(command=lambda: _find_next(-1))
    find_close_btn.config(command=_find_close)
    find_scope_all_rb.config(
        command=lambda: _find_update_status_for_selection(
            _active_journal_text_widget(), find_var.get(), "", sel_widget=_active_journal_text_widget()
        )
    )
    find_scope_one_rb.config(
        command=lambda: _find_update_status_for_selection(
            _active_journal_text_widget(), find_var.get(), "", sel_widget=_active_journal_text_widget()
        )
    )
    find_case_chk.config(
        command=lambda: _find_update_status_for_selection(
            _active_journal_text_widget(), find_var.get(), "", sel_widget=_active_journal_text_widget()
        )
    )
    find_word_chk.config(
        command=lambda: _find_update_status_for_selection(
            _active_journal_text_widget(), find_var.get(), "", sel_widget=_active_journal_text_widget()
        )
    )
    find_entry.bind("<Return>", lambda _e: _find_next(1), add="+")
    find_entry.bind("<Shift-Return>", _find_prev, add="+")
    find_entry.bind(
        "<KeyRelease>",
        lambda _e: _find_update_status_for_selection(_active_journal_text_widget(), find_var.get(), ""),
        add="+",
    )
    for _b in (find_prev_btn, find_next_btn, find_close_btn):
        bind_button_hover_if_enabled(
            _b,
            lambda: th().toolbar_bind_rest(),
            lambda: th().toolbar_hover()[0],
            lambda: th().toolbar_hover()[1],
        )

    def gen_rest_style() -> Tuple[str, str, str, str, str]:
        if str(gen_button.cget("state")) != "normal":
            return th().gen_bind_disabled()
        return th().gen_bind_rest()

    bind_button_hover_if_enabled(
        gen_button,
        gen_rest_style,
        lambda: th().hover_primary,
        lambda: "white",
    )

    save_entry_btn_holder: Dict[str, Any] = {"btn": None}

    def refresh_save_entry_state() -> None:
        btn = save_entry_btn_holder.get("btn")
        if btn is None:
            return
        has_any = bool(
            text_box.get("1.0", "end-1c").strip()
            or stt_box.get("1.0", "end-1c").strip()
            or report_box.get("1.0", "end-1c").strip()
        )
        t = th()
        if has_any:
            btn.config(
                state="normal",
                bg=t.accent,
                fg="white",
                activebackground=t.hover_save,
                activeforeground="white",
                cursor="hand2",
            )
        else:
            btn.config(
                state="disabled",
                bg=t.btn_disabled,
                fg=t.disabled_fg,
                disabledforeground=t.disabled_fg,
                cursor="arrow",
            )

    def load_latest_entry_into_current_journal(values: Dict[str, object]) -> None:
        """
        Load an existing journal record into the currently visible journal text boxes.

        Used by the GUI console command `JS -> EDITPREV` so we don't open a new journal editor window.
        """
        nonlocal edit_target_sheet, edit_target_row
        # Journal console helpers pass keys from `get_latest_journal_entry_for_edit()`.
        edit_target_sheet = str(values.get("sheet_name", "") or "")
        try:
            edit_target_row = int(values.get("row_index", 0) or 0)
        except (TypeError, ValueError):
            edit_target_row = 0
        is_edit_mode["v"] = bool(edit_target_sheet and edit_target_row > 0)

        new_text = str(values.get("text", "") or "")
        new_speech = str(values.get("speech_transcript", "") or "")
        new_report = str(values.get("ai_report", "") or "")
        new_date = str(values.get("date", "") or "")
        new_time = str(values.get("time", "") or "")

        text_box.delete("1.0", "end")
        text_box.insert("1.0", new_text)
        stt_box.delete("1.0", "end")
        stt_box.insert("1.0", new_speech)
        report_box.delete("1.0", "end")
        report_box.insert("1.0", new_report)

        # DateEntry may be `DateEntry` or `tk.Entry`; try both.
        try:
            date_entry.delete(0, "end")
            date_entry.insert(0, new_date)
        except Exception:
            try:
                if new_date:
                    date_entry.set_date(new_date)  # type: ignore[attr-defined]
            except Exception:
                pass
        try:
            time_entry.delete(0, "end")
            time_entry.insert(0, new_time)
        except Exception:
            pass

        last_journal_wav["path"] = None
        _set_stt_saved_path_display("")
        stt_status.config(text="")
        report_status.config(text="")
        update_transcribe_ui()
        refresh_save_entry_state()
        try:
            text_box.focus_set()
        except tk.TclError:
            pass

    def load_draft_into_current_journal() -> bool:
        """Load saved journal window draft into the current journal editor widgets."""
        draft = load_journal_window_draft()
        if not draft:
            return False
        # Drafts should not be treated as "edit existing row".
        nonlocal edit_target_sheet, edit_target_row
        edit_target_sheet = ""
        edit_target_row = 0
        is_edit_mode["v"] = False

        text_box.delete("1.0", "end")
        text_box.insert("1.0", str(draft.get("text", "") or ""))
        stt_box.delete("1.0", "end")
        stt_box.insert("1.0", str(draft.get("speech_transcript", "") or ""))
        report_box.delete("1.0", "end")
        report_box.insert("1.0", str(draft.get("ai_report", "") or ""))

        draft_date = str(draft.get("date", "") or "").strip()
        draft_time = str(draft.get("time", "") or "").strip()
        try:
            date_entry.delete(0, "end")
            date_entry.insert(0, draft_date)
        except Exception:
            try:
                if draft_date:
                    date_entry.set_date(draft_date)  # type: ignore[attr-defined]
            except Exception:
                pass
        try:
            time_entry.delete(0, "end")
            time_entry.insert(0, draft_time)
        except Exception:
            pass

        last_journal_wav["path"] = None
        _set_stt_saved_path_display("")
        stt_status.config(text="")
        report_status.config(text="")
        update_transcribe_ui()
        refresh_save_entry_state()
        try:
            text_box.focus_set()
        except tk.TclError:
            pass
        return True

    def start_new_journal(discard_without_confirm: bool = False) -> bool:
        """
        Clear the current journal editor widgets to start a fresh page.
        If there is unsaved content, ask for confirmation unless `discard_without_confirm` is True.
        """
        if recording_ui_busy["v"] or transcribing_busy["v"]:
            messagebox.showinfo(
                "New Journal",
                "Finish recording/transcribing before starting a new journal.",
            )
            return False

        has_content = any(
            [
                text_box.get("1.0", "end-1c").strip(),
                stt_box.get("1.0", "end-1c").strip(),
                report_box.get("1.0", "end-1c").strip(),
            ]
        )

        should_discard = True
        if has_content and not discard_without_confirm:
            should_discard = messagebox.askyesno(
                "Start new journal",
                "Discard the current journal editor content and clear the saved draft? ",
            )
        if not should_discard:
            return False

        # Clear stored draft so we don't immediately bring it back from disk.
        try:
            clear_journal_window_draft()
        except Exception:
            pass

        # Reset editor state.
        edit_target_sheet = ""
        edit_target_row = 0
        is_edit_mode["v"] = False
        last_journal_wav["path"] = None
        _set_stt_saved_path_display("")
        stt_status.config(text="")
        report_status.config(text="")
        transcribing_progress["v"] = 0

        text_box.delete("1.0", "end")
        stt_box.delete("1.0", "end")
        report_box.delete("1.0", "end")

        # Reset date/time to now.
        now = datetime.now()
        current_now_date = now.strftime("%m/%d/%Y")
        current_now_time = now.strftime("%I:%M%p").lstrip("0")
        try:
            date_entry.delete(0, "end")
            date_entry.insert(0, current_now_date)
        except Exception:
            try:
                date_entry.set_date(current_now_date)  # type: ignore[attr-defined]
            except Exception:
                pass
        try:
            time_entry.delete(0, "end")
            time_entry.insert(0, current_now_time)
        except Exception:
            pass

        update_transcribe_ui()
        refresh_save_entry_state()

        try:
            text_box.focus_set()
        except tk.TclError:
            pass
        return True

    record_stop = threading.Event()
    record_pause = threading.Event()
    record_thread_holder: Dict[str, object] = {"thread": None}
    record_path_holder: Dict[str, object] = {"path": None}
    recording_ui_busy = {"v": False}
    last_journal_wav: Dict[str, Optional[Path]] = {"path": None}
    transcribing_busy = {"v": False}
    transcribing_progress: Dict[str, int] = {"v": 0}
    wave_lock = threading.Lock()
    wave_holder: Dict[str, List[float]] = {"levels": []}
    wave_gate: Dict[str, Any] = {"rms": 0.0}
    wave_after: Dict[str, Optional[Any]] = {"id": None}

    def cancel_wave_tick() -> None:
        wid = wave_after["id"]
        if wid is not None:
            try:
                root.after_cancel(wid)
            except tk.TclError:
                pass
            wave_after["id"] = None

    def redraw_waveform_canvas() -> None:
        with wave_lock:
            pts = list(wave_holder["levels"])
        wave_canvas.update_idletasks()
        wpx = max(40, int(wave_canvas.winfo_width()))
        hpx = max(30, int(wave_canvas.winfo_height()))
        wave_canvas.delete("all")
        mid = hpx * 0.5
        t = th()
        base_color = t.muted if isinstance(t.muted, str) else "#888888"
        if len(pts) < 1:
            wave_canvas.create_line(4, mid, wpx - 4, mid, fill=base_color, width=1)
            return
        if len(pts) == 1:
            v = float(pts[0])
            y0 = mid - v * (hpx * 0.38)
            wave_canvas.create_line(4, y0, wpx - 4, y0, fill=t.waveform, width=1)
            return
        n = len(pts)
        coords: List[float] = []
        for i, v in enumerate(pts):
            x = 4.0 + (wpx - 8.0) * (i / max(n - 1, 1))
            y = mid - float(v) * (hpx * 0.38)
            coords.extend([x, y])
        wave_canvas.create_line(*coords, fill=t.waveform, width=1)

    def wave_tick() -> None:
        if not recording_ui_busy["v"]:
            wave_after["id"] = None
            return
        redraw_waveform_canvas()
        wave_after["id"] = root.after(33, wave_tick)

    def start_wave_tick() -> None:
        cancel_wave_tick()
        wave_after["id"] = root.after(33, wave_tick)

    def reset_waveform_session() -> None:
        cancel_wave_tick()
        with wave_lock:
            wave_holder["levels"].clear()
        wave_gate["rms"] = 0.0
        redraw_waveform_canvas()

    def on_pcm_block_journal(block: Any) -> None:
        try:
            import numpy as np

            flat = np.asarray(block, dtype=np.float64).reshape(-1)
            if flat.size == 0:
                return
            rms = float(np.sqrt(np.mean(flat * flat)))
            wave_gate["rms"] = rms
            adj = max(0.0, rms - WAVEFORM_RMS_NOISE_FLOOR)
            denom = max(WAVEFORM_RMS_NORM - WAVEFORM_RMS_NOISE_FLOOR, 1.0)
            peak = min(1.0, adj / denom)
            with wave_lock:
                levels = wave_holder["levels"]
                levels.append(peak)
                over = len(levels) - WAVEFORM_MAX_DRAW_SAMPLES
                if over > 0:
                    del levels[:over]
        except Exception:
            pass

    def update_transcribe_ui() -> None:
        t = th()
        if transcribing_busy["v"]:
            tb = t.transcribe_busy_config()
            transcribe_btn.config(
                state="disabled",
                width=JOURNAL_SIDE_ACTION_BTN_WIDTH_CH,
                bg=tb[0],
                fg=tb[1],
                activebackground=tb[2],
                activeforeground=tb[3],
                disabledforeground=tb[4],
            )
            return
        if recording_ui_busy["v"]:
            tb = t.transcribe_idle_disabled_config()
            transcribe_btn.config(
                state="disabled",
                width=JOURNAL_SIDE_ACTION_BTN_WIDTH_CH,
                bg=tb[0],
                fg=tb[1],
                activebackground=tb[2],
                activeforeground=tb[3],
                disabledforeground=tb[4],
            )
            return
        p = last_journal_wav.get("path")
        has_session = p is not None and isinstance(p, Path) and p.exists()
        has_archived = latest_archived_journal_wav() is not None
        if has_session or has_archived:
            bg, fg, abg, afg = t.side_action_config()
            transcribe_btn.config(
                state="normal",
                width=JOURNAL_SIDE_ACTION_BTN_WIDTH_CH,
                bg=bg,
                fg=fg,
                activebackground=abg,
                activeforeground=afg,
            )
        else:
            tb = t.transcribe_idle_disabled_config()
            transcribe_btn.config(
                state="disabled",
                width=JOURNAL_SIDE_ACTION_BTN_WIDTH_CH,
                bg=tb[0],
                fg=tb[1],
                activebackground=tb[2],
                activeforeground=tb[3],
                disabledforeground=tb[4],
            )

    def transcribe_tooltip_text() -> str:
        if transcribing_busy["v"]:
            pct = int(transcribing_progress.get("v", 0))
            return tr("journal.transcribe_tooltip_busy_full").format(pct=pct)
        if recording_ui_busy["v"]:
            return tr("journal.transcribe_tooltip_wait_recording")
        p = last_journal_wav.get("path")
        if p is not None and isinstance(p, Path) and p.exists():
            return tr("journal.transcribe_tooltip_prev_session")
        if latest_archived_journal_wav() is not None:
            return tr("journal.transcribe_tooltip_archived")
        return tr("journal.transcribe_tooltip_no_recording").format(dir=str(RECORDING_DIR))

    def run_transcribe() -> None:
        if transcribing_busy["v"]:
            return
        p = last_journal_wav.get("path")
        cleared_stale_cache = False
        if p is not None and isinstance(p, Path) and not p.exists():
            last_journal_wav["path"] = None
            cleared_stale_cache = True
            update_transcribe_ui()
            p = None
        if p is not None and isinstance(p, Path) and p.exists():
            use_path = p
        else:
            archived = latest_archived_journal_wav()
            if archived is None:
                messagebox.showinfo(
                    "Speech to text",
                    "No recording is available. Record audio first, or save a journal recording "
                    f"to your Recording folder:\n{RECORDING_DIR}",
                )
                return
            if cleared_stale_cache:
                use_path = archived
                last_journal_wav["path"] = archived
            elif not messagebox.askyesno(
                "Speech to text",
                "There is no recording from this session.\n\n"
                "Would you like to transcribe the most recent saved file in your Recording folder?\n\n"
                f"{archived.name}",
            ):
                return
            else:
                use_path = archived
                last_journal_wav["path"] = archived
        if not use_path.exists():
            last_journal_wav["path"] = None
            alt = latest_archived_journal_wav()
            if alt is None:
                messagebox.showinfo(
                    "Speech to text",
                    "That recording file is no longer on disk. There are no other saved "
                    f"recordings in:\n{RECORDING_DIR}",
                )
                update_transcribe_ui()
                return
            use_path = alt
            last_journal_wav["path"] = alt
        if not get_openai_api_key():
            messagebox.showerror(
                "Speech to text",
                "No OpenAI API key. Use TOKEN ADD in the main menu or set OPENAI_API_KEY.",
            )
            return
        transcribing_progress["v"] = 0

        def schedule_progress(pct: int) -> None:
            p = min(100, max(0, int(pct)))
            transcribing_progress["v"] = p

            def _ui() -> None:
                try:
                    stt_status.config(text=f"Transcribing… ({p}%)")
                except tk.TclError:
                    pass

            root.after(0, _ui)

        transcribing_busy["v"] = True
        update_transcribe_ui()
        schedule_progress(0)
        lang_snap = _language_code_for_whisper()

        def work() -> None:
            result = ""
            try:
                result = transcribe_audio_openai(
                    use_path,
                    lang_snap,
                    temperature=0.0,
                    progress=schedule_progress,
                )
            except BaseException as _tw_exc:
                result = f"Whisper request failed: {_tw_exc}"
            finally:
                try:
                    schedule_progress(100)
                except Exception:
                    pass

            def done() -> None:
                transcribing_busy["v"] = False
                transcribing_progress["v"] = 0
                update_transcribe_ui()
                stt_status.config(text="")
                if _is_likely_api_error_message(result):
                    messagebox.showerror("Speech to text", result[:4000])
                    return
                final_text = result.strip()
                if final_text:
                    if stt_box.get("1.0", "end-1c").strip():
                        stt_box.insert("end", " ")
                    stt_box.insert("end", final_text)
                save_draft()
                refresh_save_entry_state()

            root.after(0, done)

        threading.Thread(target=work, daemon=True).start()

    transcribe_btn.config(command=run_transcribe)
    bind_hover_tooltip(transcribe_btn, transcribe_tooltip_text)

    def transcribe_rest_style() -> Tuple[str, str, str, str, str]:
        t = th()
        if transcribing_busy["v"]:
            b0, b1, b2, b3, b4 = t.transcribe_busy_config()
            return ("disabled", b0, b1, b2, b3)
        if recording_ui_busy["v"]:
            b0, b1, b2, b3, b4 = t.transcribe_idle_disabled_config()
            return ("disabled", b0, b1, b2, b3)
        p = last_journal_wav.get("path")
        if (p is not None and isinstance(p, Path) and p.exists()) or (
            latest_archived_journal_wav() is not None
        ):
            return t.side_action_bind_rest()
        b0, b1, b2, b3, b4 = t.transcribe_idle_disabled_config()
        return ("disabled", b0, b1, b2, b3)

    bind_button_hover_if_enabled(
        transcribe_btn,
        transcribe_rest_style,
        lambda: th().hover_primary,
        lambda: "white",
    )
    update_transcribe_ui()
    wave_canvas.bind("<Configure>", lambda _e: redraw_waveform_canvas())
    wave_canvas.after(80, redraw_waveform_canvas)

    def _editor_has_meaningful_body_content() -> bool:
        return bool(
            text_box.get("1.0", "end-1c").strip()
            or stt_box.get("1.0", "end-1c").strip()
            or report_box.get("1.0", "end-1c").strip()
        )

    def _draft_file_has_restorable_content(d: Dict[str, object]) -> bool:
        if str(d.get("text", "") or "").strip():
            return True
        if str(d.get("speech_transcript", "") or "").strip():
            return True
        if str(d.get("ai_report", "") or "").strip():
            return True
        wraw = d.get("journal_recording_wav")
        if not wraw:
            return False
        try:
            return Path(str(wraw)).is_file()
        except OSError:
            return False

    def apply_draft_dict_to_ui(d: Dict[str, object]) -> None:
        nonlocal edit_target_sheet, edit_target_row
        _txt = str(d.get("text", "") or "")
        _sp = str(d.get("speech_transcript", "") or "")
        _rp = str(d.get("ai_report", "") or "")
        text_box.delete("1.0", "end")
        text_box.insert("1.0", _txt)
        stt_box.delete("1.0", "end")
        stt_box.insert("1.0", _sp)
        report_box.delete("1.0", "end")
        report_box.insert("1.0", _rp)
        _dt = str(d.get("date", "") or "").strip()
        if _dt:
            if DateEntry is not None and isinstance(date_entry, DateEntry):  # type: ignore[arg-type]
                try:
                    date_entry.set_date(_dt)  # type: ignore[attr-defined]
                except Exception:
                    try:
                        date_entry.delete(0, "end")  # type: ignore[attr-defined]
                    except Exception:
                        pass
                    date_entry.insert(0, _dt)  # type: ignore[attr-defined]
            else:
                date_entry.delete(0, "end")
                date_entry.insert(0, _dt)
        _tm = str(d.get("time", "") or "").strip()
        if _tm:
            time_entry.delete(0, "end")
            time_entry.insert(0, _tm)
        edit_target_sheet = str(d.get("edit_target_sheet", "") or "")
        try:
            edit_target_row = int(d.get("edit_target_row", 0) or 0)
        except (TypeError, ValueError):
            edit_target_row = 0
        is_edit_mode["v"] = bool(edit_target_sheet and edit_target_row > 0)
        _wav = d.get("journal_recording_wav")
        last_journal_wav["path"] = None
        if _wav:
            try:
                _wp = Path(str(_wav))
                if _wp.is_file():
                    last_journal_wav["path"] = _wp.resolve()
                    _set_stt_saved_path_display(tr("journal.saved_path", path=str(_wp)))
                else:
                    _set_stt_saved_path_display("")
            except OSError:
                _set_stt_saved_path_display("")
        else:
            _set_stt_saved_path_display("")
        stt_status.config(text="")
        report_status.config(text="")
        update_transcribe_ui()
        refresh_save_entry_state()
        try:
            text_box.focus_set()
        except tk.TclError:
            pass

    def on_restore_draft_click() -> None:
        d = load_journal_window_draft()
        if not isinstance(d, dict) or not _draft_file_has_restorable_content(d):
            messagebox.showinfo(
                tr("msg.journal_window"),
                tr("msg.no_draft_to_restore"),
            )
            return
        if _editor_has_meaningful_body_content():
            if not messagebox.askyesno(
                tr("journal.restore_confirm_title"),
                tr("journal.restore_confirm"),
            ):
                return
        apply_draft_dict_to_ui(d)
        save_draft()

    journal_top_actions = tk.Frame(top, bg=t_init.panel)
    journal_top_actions.grid(row=0, column=5, sticky="w", padx=(4, 8), pady=12)
    restore_draft_btn = tk.Button(
        journal_top_actions,
        text=tr("journal.restore_draft"),
        command=on_restore_draft_click,
        bg=_ut_bg,
        fg=_ut_fg,
        activebackground=_ut_abg,
        activeforeground=_ut_afg,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=10,
        pady=6,
        cursor="hand2",
    )
    restore_draft_btn.pack(side="left")
    bind_button_hover_if_enabled(
        restore_draft_btn,
        lambda: th().toolbar_bind_rest(),
        lambda: th().toolbar_hover()[0],
        lambda: th().toolbar_hover()[1],
    )
    bind_hover_tooltip(restore_draft_btn, lambda: tr("tip.restore_draft"))

    def _language_code_for_whisper() -> Optional[str]:
        choice = lang_var.get().strip()
        if choice == "English":
            return "en"
        if choice in ("简体中文", "中文", "Chinese"):
            return "zh"
        return None

    def _is_likely_api_error_message(text: str) -> bool:
        t = text.strip()
        if not t:
            return False
        prefixes = (
            "OPENAI_API_KEY",
            "ChatGPT API error",
            "Failed to contact ChatGPT",
            "ChatGPT returned",
            "No response received",
            "Whisper API error",
            "Whisper request failed",
            "Whisper returned",
            "Could not read audio file",
            "Recording needs optional packages",
            "No speech detected",
            "Empty audio.",
        )
        return any(t.startswith(p) for p in prefixes)

    def _journal_rec_btn_set(btn: Any, enabled: bool) -> None:
        t = th()
        if enabled:
            bg, fg, abg, afg = t.side_action_config()
            btn.config(
                state="normal",
                bg=bg,
                fg=fg,
                activebackground=abg,
                activeforeground=afg,
                cursor="hand2",
            )
        else:
            btn.config(
                state="disabled",
                bg=t.btn_disabled,
                fg=t.disabled_fg,
                disabledforeground=t.disabled_fg,
                cursor="arrow",
            )

    def on_record_worker_finished(err: Optional[str], wav_path: Optional[Path]) -> None:
        record_thread_holder["thread"] = None
        record_pause.clear()
        cancel_wave_tick()
        if err:
            recording_ui_busy["v"] = False
            last_journal_wav["path"] = None
            update_transcribe_ui()
            stt_status.config(text="")
            _set_stt_saved_path_display("")
            _journal_rec_btn_set(start_rec_button, True)
            _journal_rec_btn_set(pause_rec_button, False)
            _journal_rec_btn_set(stop_rec_button, False)
            pause_rec_button.config(text="Pause recording")
            if ttk is not None:
                lang_combo.config(state="readonly")
            else:
                lang_combo.config(state="normal")
            if wav_path is not None:
                try:
                    wav_path.unlink(missing_ok=True)
                except OSError:
                    pass
            reset_waveform_session()
            messagebox.showerror("Speech to text", err[:4000])
            return
        if wav_path is None or not wav_path.exists():
            recording_ui_busy["v"] = False
            last_journal_wav["path"] = None
            update_transcribe_ui()
            stt_status.config(text="")
            _set_stt_saved_path_display("")
            _journal_rec_btn_set(start_rec_button, True)
            _journal_rec_btn_set(pause_rec_button, False)
            _journal_rec_btn_set(stop_rec_button, False)
            pause_rec_button.config(text="Pause recording")
            if ttk is not None:
                lang_combo.config(state="readonly")
            else:
                lang_combo.config(state="normal")
            reset_waveform_session()
            return
        dest = archive_journal_recording(wav_path)
        try:
            wav_path.unlink(missing_ok=True)
        except OSError:
            pass
        recording_ui_busy["v"] = False
        _journal_rec_btn_set(start_rec_button, True)
        _journal_rec_btn_set(pause_rec_button, False)
        _journal_rec_btn_set(stop_rec_button, False)
        pause_rec_button.config(text="Pause recording")
        if ttk is not None:
            lang_combo.config(state="readonly")
        else:
            lang_combo.config(state="normal")
        reset_waveform_session()
        if dest is not None:
            last_journal_wav["path"] = dest
            stt_status.config(text="")
            _set_stt_saved_path_display(tr("journal.saved_path", path=str(dest)))
        else:
            last_journal_wav["path"] = None
            _set_stt_saved_path_display("")
            stt_status.config(
                text=f"Recording finished (could not copy to {RECORDING_DIR}).",
            )
        update_transcribe_ui()

    def record_worker_main(wav_path: Path) -> None:
        err = record_microphone_session_wav(
            wav_path,
            record_stop,
            chunk_interval_sec=LIVE_STT_CHUNK_INTERVAL_SEC,
            on_audio_chunk=None,
            on_pcm_block=on_pcm_block_journal,
            pause_event=record_pause,
        )
        root.after(0, lambda: on_record_worker_finished(err, wav_path))

    def start_recording() -> None:
        if recording_ui_busy["v"]:
            return
        fd, tmp_name = tempfile.mkstemp(suffix=".wav", prefix="journal_mic_")
        os.close(fd)
        tmp = Path(tmp_name)
        record_path_holder["path"] = tmp
        last_journal_wav["path"] = None
        update_transcribe_ui()
        _set_stt_saved_path_display("")
        record_stop.clear()
        record_pause.clear()
        recording_ui_busy["v"] = True
        reset_waveform_session()
        start_wave_tick()
        _journal_rec_btn_set(start_rec_button, False)
        _journal_rec_btn_set(pause_rec_button, True)
        _journal_rec_btn_set(stop_rec_button, True)
        pause_rec_button.config(text="Pause recording")
        if ttk is not None:
            lang_combo.config(state="disabled")
        else:
            lang_combo.config(state="disabled")
        stt_status.config(text="Recording…")
        th = threading.Thread(target=record_worker_main, args=(tmp,), daemon=True)
        record_thread_holder["thread"] = th
        th.start()

    def stop_recording() -> None:
        th = record_thread_holder["thread"]
        if not (
            recording_ui_busy["v"]
            and isinstance(th, threading.Thread)
            and th.is_alive()
        ):
            return
        stt_status.config(text="Stopping…")
        record_stop.set()
        record_pause.clear()
        _journal_rec_btn_set(start_rec_button, False)
        _journal_rec_btn_set(pause_rec_button, False)
        _journal_rec_btn_set(stop_rec_button, False)

    def toggle_pause_recording() -> None:
        th = record_thread_holder["thread"]
        if not (
            recording_ui_busy["v"]
            and isinstance(th, threading.Thread)
            and th.is_alive()
        ):
            return
        if record_pause.is_set():
            record_pause.clear()
            pause_rec_button.config(text="Pause recording")
            stt_status.config(text="Recording…")
        else:
            record_pause.set()
            pause_rec_button.config(text="Resume recording")
            stt_status.config(text="Recording paused")

    _sr_bg, _sr_fg, _sr_abg, _sr_afg = t_init.side_action_config()
    start_rec_button = tk.Button(
        stt_top,
        text="Start recording",
        command=start_recording,
        bg=_sr_bg,
        fg=_sr_fg,
        activebackground=_sr_abg,
        activeforeground=_sr_afg,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=10,
        pady=6,
        cursor="hand2",
    )
    pause_rec_button = tk.Button(
        stt_top,
        text="Pause recording",
        command=toggle_pause_recording,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=8,
        pady=6,
    )
    stop_rec_button = tk.Button(
        stt_top,
        text="Stop recording",
        command=stop_recording,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=8,
        pady=6,
    )
    _journal_rec_btn_set(pause_rec_button, False)
    _journal_rec_btn_set(stop_rec_button, False)
    start_rec_button.grid(row=0, column=0, sticky="w", padx=(12, 4), pady=8)
    pause_rec_button.grid(row=0, column=1, sticky="w", padx=(0, 4), pady=8)
    stop_rec_button.grid(row=0, column=2, sticky="w", padx=(0, 8), pady=8)

    def rec_primary_rest(btn: Any) -> Tuple[str, str, str, str, str]:
        t = th()
        if str(btn.cget("state")) == "normal":
            return t.side_action_bind_rest()
        return t.side_action_disabled()

    bind_button_hover_if_enabled(
        start_rec_button,
        lambda b=start_rec_button: rec_primary_rest(b),
        lambda: th().hover_primary,
        lambda: "white",
    )
    bind_button_hover_if_enabled(
        pause_rec_button,
        lambda b=pause_rec_button: rec_primary_rest(b),
        lambda: th().hover_primary,
        lambda: "white",
    )
    bind_button_hover_if_enabled(
        stop_rec_button,
        lambda b=stop_rec_button: rec_primary_rest(b),
        lambda: th().hover_primary,
        lambda: "white",
    )
    stt_status.grid(row=0, column=3, sticky="ew", padx=(4, 12), pady=8)
    stt_lang_lbl = tk.Label(
        stt_top,
        text="Language:",
        bg=t_init.panel,
        fg=t_init.muted,
        font=("Segoe UI", 9),
    )
    stt_lang_lbl.grid(row=0, column=4, sticky="w", padx=(4, 0), pady=8)
    lang_combo.grid(row=0, column=5, sticky="ew", padx=(8, 12), pady=8)

    def run_generate_report() -> None:
        if not get_openai_api_key():
            messagebox.showerror(
                "Journal Window",
                "No OpenAI API key. Use TOKEN ADD in the main menu or set OPENAI_API_KEY.",
            )
            return
        t = th()
        gen_button.config(
            state="disabled",
            bg=t.btn_disabled,
            fg=t.disabled_fg,
            disabledforeground=t.disabled_fg,
            cursor="arrow",
        )
        report_status.config(text="Generating…")

        def work() -> None:
            body = generate_journal_report_from_sources(
                text_box.get("1.0", "end-1c"),
                stt_box.get("1.0", "end-1c"),
            )
            root.after(0, lambda b=body: on_generate_report_done(b))

        threading.Thread(target=work, daemon=True).start()

    def on_generate_report_done(body: str) -> None:
        t = th()
        bg, fg, abg, afg = t.side_action_config()
        gen_button.config(
            state="normal",
            bg=bg,
            fg=fg,
            activebackground=abg,
            activeforeground=afg,
            cursor="hand2",
        )
        report_status.config(text="")
        if _is_likely_api_error_message(body):
            messagebox.showerror("AI report", body[:4000])
            return
        report_box.delete("1.0", "end")
        report_box.insert("1.0", body.strip())
        save_draft()
        refresh_save_entry_state()

    gen_button.config(command=run_generate_report)

    def generate_report_tooltip_text() -> str:
        return tr("tip.generate_report")

    bind_hover_tooltip(gen_button, generate_report_tooltip_text)

    saved = {"value": False}
    autosave_id = {"value": None}

    def build_draft_dict() -> Dict[str, object]:
        return {
            "text": text_box.get("1.0", "end-1c"),
            "speech_transcript": stt_box.get("1.0", "end-1c"),
            "ai_report": report_box.get("1.0", "end-1c"),
            "date": date_entry.get().strip(),
            "time": time_entry.get().strip(),
            "edit_target_sheet": edit_target_sheet,
            "edit_target_row": edit_target_row,
            "updated_at": datetime.now().isoformat(),
        }

    def save_draft() -> None:
        save_journal_window_draft(build_draft_dict())

    def autosave() -> None:
        save_draft()
        autosave_id["value"] = root.after(1500, autosave)

    def do_save() -> None:
        if not (
            text_box.get("1.0", "end-1c").strip()
            or stt_box.get("1.0", "end-1c").strip()
            or report_box.get("1.0", "end-1c").strip()
        ):
            return
        raw_date = date_entry.get().strip()
        parsed_date = parse_flexible_date(raw_date, now.year)
        if parsed_date is None:
            messagebox.showerror("Journal Window", "Invalid date. Example: 04/20/2026 or Apr 20")
            return
        date_value = parsed_date.strftime("%m/%d/%Y")
        normalized_time = normalize_window_time_input(time_entry.get().strip())
        if normalized_time is None:
            messagebox.showerror("Journal Window", "Invalid time. Example: 2:03PM")
            return
        text_value = text_box.get("1.0", "end-1c").strip()
        if not text_value and is_edit_mode["v"]:
            should_delete = messagebox.askyesno(
                "Clear Entry",
                "Text is empty. Saving now will delete the previous entry. Are you sure?",
            )
            if not should_delete:
                return
            deleted_ok = delete_journal_entry_at(edit_target_sheet, edit_target_row)
            if not deleted_ok:
                messagebox.showerror(
                    "Journal Window",
                    "Could not delete previous entry. It may have changed. Try again.",
                )
                return
            clear_journal_window_draft()
            saved["value"] = True
            if autosave_id["value"] is not None:
                root.after_cancel(autosave_id["value"])
            root.destroy()
            return
        if not text_value:
            text_value = "(no details entered)"
        speech_value = stt_box.get("1.0", "end-1c").strip()
        report_value = report_box.get("1.0", "end-1c").strip()
        row_payload = [date_value, normalized_time, text_value, speech_value, report_value]
        if edit_target_sheet and edit_target_row > 0:
            saved_ok = update_journal_entry_at(
                edit_target_sheet,
                edit_target_row,
                row_payload,
            )
            if not saved_ok:
                messagebox.showerror(
                    "Journal Window",
                    "Could not update previous entry. It may have changed. Try again.",
                )
                return
        else:
            append_row(MODULES["J"], row_payload)
        clear_journal_window_draft()
        saved["value"] = True
        if autosave_id["value"] is not None:
            root.after_cancel(autosave_id["value"])
        root.destroy()

    def on_close(event=None) -> None:
        has_content = any(
            [
                text_box.get("1.0", "end-1c").strip(),
                stt_box.get("1.0", "end-1c").strip(),
                report_box.get("1.0", "end-1c").strip(),
            ]
        )
        if not has_content:
            clear_journal_window_draft()
            if autosave_id["value"] is not None:
                root.after_cancel(autosave_id["value"])
            root.destroy()
            return
        save_choice = messagebox.askyesnocancel(
            "Close Journal Window",
            "Do you want to save this journal entry before closing?",
        )
        if save_choice is None:
            return
        if save_choice:
            do_save()
            return
        should_discard = messagebox.askyesno(
            "Discard Changes",
            "Are you sure you want to close without saving to journal? Draft backup is kept.",
        )
        if not should_discard:
            return
        save_draft()
        if autosave_id["value"] is not None:
            root.after_cancel(autosave_id["value"])
        root.destroy()

    console_history: List[str] = []
    console_hist_index = {"value": 0}
    prefs_for_console = load_preferences()
    console_app_name = {"value": prefs_for_console.get("app_name", "Daily Logger") or "Daily Logger"}
    # GUI console "JS" state: when user types JS, we enter a non-freezing sub-mode.
    js_gui_state: Dict[str, object] = {"active": False}

    def console_append(text: str) -> None:
        if not text:
            return
        console_output.config(state="normal")
        console_output.insert("end", text.rstrip("\n") + "\n")
        console_output.see("end")
        console_output.config(state="disabled")

    def run_console_command() -> None:
        if console_entry_state["placeholder"]:
            return
        raw = console_entry.get().strip()
        console_entry.delete(0, "end")
        if not raw:
            return
        console_append(f"> {raw}")
        if not console_history or console_history[-1] != raw:
            console_history.append(raw)
        console_hist_index["value"] = len(console_history)
        cmd = raw.upper()
        if bool(js_gui_state.get("active")):
            # First token after JS is treated as "journal_settings_menu" choice.
            # This mirrors the old CLI menu, but stays non-blocking for the Tk GUI.
            try:
                choice = raw.strip()
                choice_key = choice.upper()
                help_text = (
                    "Journal settings:\n"
                    "  WINDOW               - open window editor\n"
                    "  CONSOLE              - type journal text in console\n"
                    "  EDITPREV             - edit latest entry in window\n"
                    "  DP                   - delete latest entry\n"
                    "  RESTORE              - reopen latest unsaved draft\n"
                    "  HELP                 - show this list\n"
                    "  Enter                - return to main menu\n"
                    "  DEFAULT WINDOWS     - set preferred journal input to window\n"
                    "  DEFAULT CONSOLE      - set preferred journal input to console"
                )
                if is_enter_equivalent(choice_key) or not choice:
                    js_gui_state["active"] = False
                    console_append("JS menu closed.")
                    return
                if choice_key == "HELP":
                    console_append(help_text)
                    return
                if choice_key in ("W", "WINDOW", "WINDOWS"):
                    js_gui_state["active"] = False
                    show_page("journal")
                    open_journal_window_editor()
                    console_append("Opened Journal window editor.")
                    return
                if choice_key in ("C", "CONSOLE", "CONSOLE", "COINSOLE"):
                    # Multi-step: typed note + date/time.
                    js_gui_state["active"] = False
                    show_page("journal")
                    typed_note = _ask_typed_note_gui(root)
                    if typed_note is None:
                        console_append("Journal CONSOLE cancelled.")
                        return
                    dt = ask_entry_date_time_gui(root)
                    if dt is None:
                        console_append("Journal date/time cancelled.")
                        return
                    date_value, time_value = dt
                    append_row(MODULES["J"], [date_value, time_value, typed_note, "", ""])
                    console_append(f'Journal saved to: {DATA_DIR / MODULES["J"].workbook_name}')
                    return
                if choice_key in ("EDITPREV", "EDIT PREV", "EDIT PREVIOUS", "OPENPREV", "OPEN PREV", "OPENPREVIOUS", "OPEN PREVIOUS"):
                    js_gui_state["active"] = False
                    show_page("journal")
                    latest = get_latest_journal_entry_for_edit()
                    if not latest:
                        console_append("No previous journal entry found to edit.")
                        return
                    try:
                        load_latest_entry_into_current_journal(latest)
                        console_append("Loaded latest journal entry into current textbox (edit mode on).")
                    except Exception as exc:
                        console_append(f"EDITPREV failed: {exc}")
                    return
                if choice_key == "DP":
                    js_gui_state["active"] = False
                    show_page("journal")
                    latest = get_latest_journal_entry_for_delete()
                    if not latest:
                        console_append("No previous journal entry found to delete.")
                        return
                    date_label = str(latest.get("date", "")).strip() or "(unknown date)"
                    time_label = str(latest.get("time", "")).strip() or "(unknown time)"
                    if messagebox.askyesno(
                        "Delete previous journal entry",
                        f"Delete previous journal entry at {date_label} {time_label}?",
                    ):
                        delete_latest_journal_entry()
                        console_append("Deleted previous journal entry.")
                    else:
                        console_append("Delete cancelled.")
                    return
                if choice_key == "RESTORE":
                    js_gui_state["active"] = False
                    show_page("journal")
                    draft = load_journal_window_draft()
                    if not draft:
                        console_append("No journal draft to restore.")
                        return
                    open_journal_window_editor(draft)
                    console_append("Restored draft opened.")
                    return
                if choice_key in ("DEFAULT WINDOWS", "DEFAULT CONSOLE"):
                    js_gui_state["active"] = False
                    show_page("journal")
                    prefs = load_preferences()
                    default_mode = "windows" if choice_key == "DEFAULT WINDOWS" else "console"
                    prefs["journal_input_default"] = default_mode
                    if save_preferences(prefs):
                        console_append(f"Default set to {default_mode}.")
                    else:
                        console_append("Could not save default journal input preference.")
                    return
                if choice_key == "J":
                    js_gui_state["active"] = False
                    show_page("journal")
                    console_append("Switched to Journal page.")
                    return
                console_append("Unknown JS choice. Type HELP for options.")
            except Exception as exc:
                console_append(f"JS menu error: {exc}")
            return

        # Non-JS direct commands in GUI console:
        # - RESTORE: load saved journal draft into the existing editor widgets
        # - EDITPREV/OPENPREV: load latest journal entry into the existing editor widgets
        if cmd in {"RESTORE"}:
            js_gui_state["active"] = False
            show_page("journal")
            ok = load_draft_into_current_journal()
            if not ok:
                console_append("No journal draft to restore.")
            else:
                console_append("Restored draft into current journal textbox.")
            return
        if cmd in {"EDITPREV", "EDIT PREV", "EDIT PREVIOUS", "OPENPREV", "OPEN PREV", "OPEN PREVIOUS"}:
            js_gui_state["active"] = False
            show_page("journal")
            latest = get_latest_journal_entry_for_edit()
            if not latest:
                console_append("No previous journal entry found to edit.")
                return
            load_latest_entry_into_current_journal(latest)
            console_append("Loaded latest journal entry into current textbox (edit mode on).")
            return
        if cmd in {"NEW", "NEW JOURNAL"}:
            js_gui_state["active"] = False
            show_page("journal")
            ok = start_new_journal()
            if ok:
                console_append("Started new journal. Editor cleared.")
            else:
                console_append("New journal cancelled.")
            return
        # Prevent GUI freeze: "JS" triggers an interactive CLI prompt (`input()`),
        # which can block the Tk main thread.
        # GUI version uses Tk dialogs instead of ``input()``.
        if cmd in {"J SETTINGS", "J SETTING", "JOURNAL SETTINGS", "JS"}:
            show_page("journal")
            js_gui_state["active"] = True
            console_append("JS menu opened. Type HELP for available choices, then submit one choice.")
            console_append(
                "Journal settings: "
                "WINDOW | CONSOLE | EDITPREV | DP | RESTORE | DEFAULT WINDOWS | DEFAULT CONSOLE | HELP"
            )
            return
        if cmd in {"J", "JOURNAL", "WINDOW"}:
            show_page("journal")
            console_append("Switched to Journal page.")
            return
        if cmd in {"R", "RT"}:
            show_page("ai_recap")
            console_append("Switched to AI Recap page.")
            return
        if cmd in {"C", "CT"}:
            show_page("chatbot")
            console_append("Switched to Chatbot page.")
            return
        if cmd == "CONSOLE":
            if sys.platform != "win32":
                console_append("Native console show is supported on Windows only.")
                return
            try:
                console_hwnd = ctypes.windll.kernel32.GetConsoleWindow()
                if console_hwnd:
                    ctypes.windll.user32.ShowWindow(console_hwnd, 5)  # SW_SHOW
                    console_append("Native console window shown.")
                else:
                    launch_cmd = (
                        f'Set-Location -LiteralPath "{str(BASE_DIR)}"; '
                        "Write-Host \"Daily Logger on-demand console\"; "
                        "Write-Host \"You can run commands here.\"; "
                        "Write-Host \"Close this window when finished.\""
                    )
                    subprocess.Popen(
                        ["powershell", "-NoExit", "-Command", launch_cmd],
                        creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0),
                    )
                    console_append("Opened a new on-demand console window.")
            except Exception:
                console_append("Could not show native console window.")
            return
        capture = io.StringIO()
        keep_running = True
        try:
            with contextlib.redirect_stdout(capture):
                keep_running, next_name = handle_choice(raw, console_app_name["value"])
            console_app_name["value"] = next_name
        except Exception as exc:
            console_append(f"Error: {exc}")
            return
        output = capture.getvalue().strip()
        if output:
            console_append(output)
        if not keep_running:
            on_close()

    def _console_history_up(_evt: Optional[Any] = None) -> str:
        if not console_history:
            return "break"
        _clear_console_placeholder()
        console_hist_index["value"] = max(0, console_hist_index["value"] - 1)
        console_entry.delete(0, "end")
        console_entry.insert(0, console_history[console_hist_index["value"]])
        return "break"

    def _console_history_down(_evt: Optional[Any] = None) -> str:
        if not console_history:
            return "break"
        _clear_console_placeholder()
        console_hist_index["value"] = min(len(console_history), console_hist_index["value"] + 1)
        console_entry.delete(0, "end")
        if console_hist_index["value"] < len(console_history):
            console_entry.insert(0, console_history[console_hist_index["value"]])
        else:
            _set_console_placeholder()
        return "break"

    def _console_entry_focus_in(_evt: Optional[Any] = None) -> None:
        _clear_console_placeholder()

    def _console_entry_focus_out(_evt: Optional[Any] = None) -> None:
        if not console_entry.get().strip():
            console_entry.delete(0, "end")
            _set_console_placeholder()

    def _console_entry_keypress(evt: Optional[Any] = None) -> Optional[str]:
        if evt is None:
            return None
        if not console_entry_state["placeholder"]:
            return None
        if evt.keysym in {"Left", "Right", "Home", "End"}:
            return "break"
        if evt.keysym in {"BackSpace", "Delete"}:
            _clear_console_placeholder()
            return "break"
        if evt.char and evt.char >= " ":
            _clear_console_placeholder()
        return None

    def _console_tab_complete(_evt: Optional[Any] = None) -> str:
        current = "" if console_entry_state["placeholder"] else console_entry.get()
        if not current.strip():
            capture = io.StringIO()
            try:
                with contextlib.redirect_stdout(capture):
                    print_main_help()
            except Exception:
                capture = io.StringIO()
            output = capture.getvalue().strip()
            if output:
                console_append(output)
            return "break"
        completed, extended = _line_tab_extend(current, MAIN_MENU_COMPLETIONS)
        if extended:
            _clear_console_placeholder()
            console_entry.delete(0, "end")
            console_entry.insert(0, completed)
            console_entry.icursor("end")
        return "break"

    button_row = tk.Frame(journal_page, bg=t_init.surface)
    button_row.pack(
        fill="x",
        padx=t_init.pad_outer,
        pady=(0, t_init.pad_button_y),
    )
    save_entry_btn = tk.Button(
        button_row,
        text="Save Entry",
        command=do_save,
        bg=t_init.btn_disabled,
        fg=t_init.disabled_fg,
        activebackground=t_init.hover_save,
        activeforeground="white",
        disabledforeground=t_init.disabled_fg,
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        padx=18,
        pady=8,
        state="disabled",
        cursor="arrow",
    )
    save_entry_btn.pack(side="right")
    save_entry_btn_holder["btn"] = save_entry_btn

    def save_rest_style() -> Tuple[str, str, str, str, str]:
        t = th()
        if str(save_entry_btn.cget("state")) != "normal":
            return t.save_bind_disabled()
        return ("normal", t.accent, "white", t.hover_save, "white")

    bind_button_hover_if_enabled(
        save_entry_btn,
        save_rest_style,
        lambda: th().hover_save,
        lambda: "white",
    )

    def _on_journal_text_changed(_evt: Optional[Any] = None) -> None:
        refresh_save_entry_state()

    def _journal_delete_prev_word(_evt: Optional[Any] = None) -> str:
        w = root.focus_get()
        if not isinstance(w, tk.Text):
            return "break"
        try:
            if w.tag_ranges("sel"):
                w.delete("sel.first", "sel.last")
                refresh_save_entry_state()
                return "break"
            left = w.get("1.0", "insert")
            if not left:
                return "break"
            end_non_ws = len(left)
            while end_non_ws > 0 and left[end_non_ws - 1].isspace():
                end_non_ws -= 1
            start = end_non_ws
            while start > 0 and not left[start - 1].isspace():
                start -= 1
            delete_chars = len(left) - start
            if delete_chars > 0:
                w.delete(f"insert-{delete_chars}c", "insert")
        except tk.TclError:
            pass
        refresh_save_entry_state()
        return "break"

    def _journal_delete_next_word(_evt: Optional[Any] = None) -> str:
        w = root.focus_get()
        if not isinstance(w, tk.Text):
            return "break"
        try:
            if w.tag_ranges("sel"):
                w.delete("sel.first", "sel.last")
                refresh_save_entry_state()
                return "break"
            right = w.get("insert", "end-1c")
            if not right:
                return "break"
            i = 0
            n = len(right)
            while i < n and right[i].isspace():
                i += 1
            while i < n and not right[i].isspace():
                i += 1
            if i > 0:
                w.delete("insert", f"insert+{i}c")
        except tk.TclError:
            pass
        refresh_save_entry_state()
        return "break"

    def _journal_undo(_evt: Optional[Any] = None) -> str:
        w = root.focus_get()
        if not isinstance(w, tk.Text):
            return "break"
        try:
            w.edit_undo()
        except tk.TclError:
            pass
        refresh_save_entry_state()
        return "break"

    def _journal_redo(_evt: Optional[Any] = None) -> str:
        w = root.focus_get()
        if not isinstance(w, tk.Text):
            return "break"
        try:
            w.edit_redo()
        except tk.TclError:
            pass
        refresh_save_entry_state()
        return "break"

    for _tb in (text_box, stt_box, report_box):
        _tb.bind("<KeyRelease>", _on_journal_text_changed, add="+")
        _tb.bind("<FocusIn>", lambda _e, _w=_tb: _journal_find_state.__setitem__("widget", _w), add="+")
        _tb.bind("<ButtonRelease-1>", _on_journal_text_changed, add="+")
        _tb.bind("<Control-BackSpace>", _journal_delete_prev_word, add="+")
        _tb.bind("<Control-w>", _journal_delete_prev_word, add="+")
        _tb.bind("<Control-Delete>", _journal_delete_next_word, add="+")
        _tb.bind("<Control-z>", _journal_undo, add="+")
        _tb.bind("<Control-y>", _journal_redo, add="+")
        _tb.bind("<Control-Z>", _journal_redo, add="+")
        _tb.bind("<Control-a>", _journal_select_all, add="+")
        _tb.bind("<Control-f>", _find_open, add="+")
    root.bind("<Control-f>", _find_open, add="+")

    def apply_journal_window_i18n() -> None:
        try:
            splash_title.config(text=tr("splash.title", app=window_app_name))
        except tk.TclError:
            pass
        settings_title.config(text=tr("settings.title"))
        for _lbl, _key in settings_label_keys:
            _lbl.config(text=tr(_key))
        try:
            if lang_ui_combo is not None and ttk is not None:
                lang_ui_combo.config(
                    values=(tr("settings.lang.english"), tr("settings.lang.chinese"))
                )
        except tk.TclError:
            pass
        _want_ui_lang = (
            tr("settings.lang.chinese")
            if ui_lang_holder[0] == "zh"
            else tr("settings.lang.english")
        )
        if ui_lang_var.get().strip() != _want_ui_lang.strip():
            ui_lang_var.set(_want_ui_lang)
        rename_btn.config(text=tr("settings.rename_btn"))
        startup_toggle_btn.config(
            text=tr("settings.on") if startup_state["enabled"] else tr("settings.off")
        )
        settings_theme_btn.config(
            text=tr("theme.dark") if th().is_dark else tr("theme.light")
        )
        backup_mode_btn.config(text=_backup_mode_btn_label(backup_mode["value"]))
        backup_manual_btn.config(text=tr("settings.manual"))
        token_save_btn.config(text=tr("settings.save"))
        token_copy_btn.config(text=tr("settings.copy"))
        start_menu_app_btn.config(text=tr("settings.start_menu_app"))
        start_menu_journal_btn.config(text=tr("settings.start_menu_journal"))
        nav_buttons["journal"].config(text=tr("nav.journal"))
        nav_buttons["ai_recap"].config(text=tr("nav.ai_recap"))
        nav_buttons["chatbot"].config(text=tr("nav.chatbot"))
        nav_buttons["console"].config(text=tr("nav.console"))
        nav_settings_btn.config(text=tr("nav.settings"))
        date_lbl.config(text=tr("journal.date"))
        time_lbl.config(text=tr("journal.time"))
        update_time_btn.config(text=tr("journal.update_time"))
        restore_draft_btn.config(text=tr("journal.restore_draft"))
        find_lbl.config(text=tr("find.label"))
        find_scope_all_rb.config(text=tr("find.all"))
        find_scope_one_rb.config(text=tr("find.current_box"))
        find_case_chk.config(text=tr("find.case"))
        find_word_chk.config(text=tr("find.word"))
        find_prev_btn.config(text=tr("find.prev"))
        find_next_btn.config(text=tr("find.next"))
        find_close_btn.config(text=tr("find.close"))
        journal_title_lbl.config(text=tr("journal.section.journal"))
        stt_title_lbl.config(text=tr("journal.section.stt"))
        report_title_lbl.config(text=tr("journal.section.report"))
        open_recording_btn.config(text=tr("journal.open"))
        stt_lang_lbl.config(text=tr("journal.lang_label"))
        transcribe_btn.config(text=tr("journal.transcribe"))
        gen_button.config(text=tr("journal.generate_report"))
        save_entry_btn.config(text=tr("journal.save_entry"))
        start_rec_button.config(text=tr("journal.rec.start"))
        stop_rec_button.config(text=tr("journal.rec.stop"))
        if recording_ui_busy["v"]:
            pause_rec_button.config(
                text=tr("journal.rec.resume")
                if record_pause.is_set()
                else tr("journal.rec.pause")
            )
        else:
            pause_rec_button.config(text=tr("journal.rec.pause"))
        console_title.config(text=tr("console.title"))
        _show_console_hint_placeholder()
        theme_toggle_btn.config(text=tr("theme.dark") if th().is_dark else tr("theme.light"))
        _ai_i18n = getattr(build_ai_recap_and_chatbot_pages, "_i18n", None)
        if callable(_ai_i18n):
            _ai_i18n()

    def _apply_console_ui_language(new_lang: str) -> None:
        ui_lang_holder[0] = new_lang
        apply_journal_window_i18n()

    set_journal_ui_language_changed_hook(_apply_console_ui_language)

    def _on_journal_root_destroy(event: Any) -> None:
        if getattr(event, "widget", None) is root:
            set_journal_ui_language_changed_hook(None)

    root.bind("<Destroy>", _on_journal_root_destroy, add="+")

    def apply_journal_window_colors() -> None:
        t = th()
        root.configure(bg=t.surface)
        shell.configure(bg=t.surface)
        nav_rail.configure(bg=t.panel)
        nav_title.configure(bg=t.panel, fg=t.muted)
        nav_summon_btn.configure(
            bg=t.toolbar_btn_config()[0],
            fg=t.toolbar_btn_config()[1],
            activebackground=t.toolbar_btn_config()[2],
            activeforeground=t.toolbar_btn_config()[3],
        )
        nav_settings_btn.configure(
            bg=t.btn_secondary,
            fg=t.text,
            activebackground=t.secondary_hover,
            activeforeground=t.text,
        )
        content_host.configure(bg=t.surface)
        journal_page.configure(bg=t.surface)
        ai_recap_page.configure(bg=t.surface)
        chatbot_page.configure(bg=t.surface)
        console_page.configure(bg=t.surface)
        settings_page.configure(bg=t.surface)
        for _w in placeholder_frames:
            _w.configure(bg=t.surface)
        for _w in placeholder_title_labels:
            _w.configure(bg=t.surface, fg=t.text)
        for _w in placeholder_body_labels:
            _w.configure(bg=t.surface, fg=t.muted)
        _ai_theme_fn = getattr(build_ai_recap_and_chatbot_pages, "_apply_theme", None)
        if callable(_ai_theme_fn):
            _ai_theme_fn()
        settings_wrap.configure(bg=t.surface)
        settings_title.configure(bg=t.surface, fg=t.text)
        settings_status_lbl.configure(bg=t.surface, fg=t.muted)
        for _w in settings_rows:
            _w.configure(bg=t.surface)
        for _w in settings_labels:
            _w.configure(bg=t.surface, fg=t.muted)
        rename_entry.config(
            bg=t.field,
            fg=t.text,
            insertbackground=t.text,
            highlightbackground=t.border,
            highlightcolor=t.accent,
        )
        token_entry.config(
            bg=t.field,
            fg=t.text,
            insertbackground=t.text,
            highlightbackground=t.border,
            highlightcolor=t.accent,
        )
        for _btn in (
            rename_btn,
            startup_toggle_btn,
            settings_theme_btn,
            backup_mode_btn,
            backup_manual_btn,
            token_save_btn,
            token_copy_btn,
            start_menu_app_btn,
            start_menu_journal_btn,
        ):
            _btn.configure(
                bg=t.btn_secondary,
                fg=t.text,
                activebackground=t.secondary_hover,
                activeforeground=t.text,
            )
        for _btn in page_toggle_buttons:
            _btn.configure(
                bg=t.toolbar_btn_config()[0],
                fg=t.toolbar_btn_config()[1],
                activebackground=t.toolbar_btn_config()[2],
                activeforeground=t.toolbar_btn_config()[3],
            )
        for key, btn in nav_buttons.items():
            if active_page["key"] == key:
                btn.config(
                    bg=t.accent,
                    fg="white",
                    activebackground=t.hover_primary,
                    activeforeground="white",
                )
            else:
                btn.config(
                    bg=t.btn_secondary,
                    fg=t.text,
                    activebackground=t.secondary_hover,
                    activeforeground=t.text,
                )
        top.configure(bg=t.panel)
        top.pack_configure(padx=t.pad_outer, pady=t.pad_top_y)
        find_row.configure(bg=t.panel)
        find_row.pack_configure(padx=t.pad_outer, pady=(0, 6))
        find_lbl.configure(bg=t.panel, fg=t.muted)
        find_entry.config(
            bg=t.field,
            fg=t.text,
            insertbackground=t.text,
            highlightbackground=t.border,
            highlightcolor=t.accent,
        )
        find_status.configure(bg=t.panel, fg=t.muted)
        find_scope_all_rb.configure(
            bg=t.panel,
            fg=t.muted,
            activebackground=t.panel,
            activeforeground=t.text,
            selectcolor=t.panel,
        )
        find_scope_one_rb.configure(
            bg=t.panel,
            fg=t.muted,
            activebackground=t.panel,
            activeforeground=t.text,
            selectcolor=t.panel,
        )
        find_case_chk.configure(
            bg=t.panel,
            fg=t.muted,
            activebackground=t.panel,
            activeforeground=t.text,
            selectcolor=t.panel,
        )
        find_word_chk.configure(
            bg=t.panel,
            fg=t.muted,
            activebackground=t.panel,
            activeforeground=t.text,
            selectcolor=t.panel,
        )
        console_wrap.configure(bg=t.surface)
        console_title.configure(bg=t.surface, fg=t.text)
        console_output.config(
            bg=t.field,
            fg=t.text,
            insertbackground=t.text,
            highlightbackground=t.border,
            highlightcolor=t.accent,
        )
        console_scroll.config(bg=t.panel, troughcolor=t.field, activebackground=t.accent)
        console_input_row.configure(bg=t.surface)
        console_prompt.configure(bg=t.surface, fg=t.muted)
        console_entry.config(
            bg=t.field,
            fg=(t.muted if console_entry_state["placeholder"] else t.text),
            insertbackground=t.text,
            insertwidth=(0 if console_entry_state["placeholder"] else console_insertwidth_normal),
            highlightbackground=t.border,
            highlightcolor=t.accent,
            font=("Consolas", 10, "italic") if console_entry_state["placeholder"] else ("Consolas", 10),
        )
        date_lbl.configure(bg=t.panel, fg=t.muted, font=t.date_label_font)
        time_lbl.configure(bg=t.panel, fg=t.muted, font=t.date_label_font)
        try:
            date_entry.config(background=t.field, foreground=t.text)
        except tk.TclError:
            try:
                date_entry.config(
                    bg=t.field,
                    fg=t.text,
                    insertbackground=t.text,
                    highlightbackground=t.border,
                    highlightcolor=t.accent,
                )
            except tk.TclError:
                pass
        time_entry.config(
            bg=t.field,
            fg=t.text,
            insertbackground=t.text,
            highlightbackground=t.border,
            highlightcolor=t.accent,
        )
        tbg, tfg, tabg, tafg = t.toolbar_btn_config()
        update_time_btn.config(
            bg=tbg, fg=tfg, activebackground=tabg, activeforeground=tafg
        )
        restore_draft_btn.config(bg=tbg, fg=tfg, activebackground=tabg, activeforeground=tafg)
        journal_top_actions.configure(bg=t.panel)
        for _b in (find_prev_btn, find_next_btn, find_close_btn):
            _b.config(bg=tbg, fg=tfg, activebackground=tabg, activeforeground=tafg)
        theme_toggle_btn.config(
            text=t.toggle_label,
            bg=t.btn_secondary,
            fg=t.text,
            activebackground=t.secondary_hover,
            activeforeground=t.text,
        )
        settings_theme_btn.config(
            text=t.toggle_label,
            bg=t.btn_secondary,
            fg=t.text,
            activebackground=t.secondary_hover,
            activeforeground=t.text,
        )
        center.configure(bg=t.surface)
        center.pack_configure(padx=t.pad_outer, pady=(0, t.pad_center_y))
        left_col.configure(bg=t.surface)
        journal_title_lbl.configure(
            bg=t.surface, fg=t.muted, font=t.section_label_font
        )
        editor_frame.configure(bg=t.panel)
        text_box.config(
            bg=t.field,
            fg=t.text,
            insertbackground=t.text,
            highlightbackground=t.border,
            highlightcolor=t.accent,
        )
        scroll_bar.config(bg=t.panel, troughcolor=t.field, activebackground=t.accent)
        right_col.configure(bg=t.surface)
        stt_outer.configure(bg=t.surface)
        stt_header.configure(bg=t.surface)
        stt_title_lbl.configure(
            bg=t.surface, fg=t.muted, font=t.section_label_font
        )
        stt_saved_path_entry.config(
            readonlybackground=t.surface,
            fg=t.muted,
            highlightbackground=t.surface,
            selectbackground=t.accent,
        )
        open_recording_btn.config(
            bg=tbg, fg=tfg, activebackground=tabg, activeforeground=tafg
        )
        stt_top.configure(bg=t.panel)
        stt_status.configure(bg=t.panel, fg=t.muted)
        stt_frame.configure(bg=t.panel)
        transcribe_hover.configure(bg=t.panel)
        wave_canvas.config(
            bg=t.field, highlightbackground=t.border, highlightcolor=t.accent
        )
        stt_box.config(
            bg=t.field,
            fg=t.text,
            insertbackground=t.text,
            highlightbackground=t.border,
            highlightcolor=t.accent,
        )
        stt_scroll.config(bg=t.panel, troughcolor=t.field, activebackground=t.accent)
        report_outer.configure(bg=t.surface)
        report_header.configure(bg=t.surface)
        report_title_lbl.configure(
            bg=t.surface, fg=t.muted, font=t.section_label_font
        )
        report_status.configure(bg=t.surface, fg=t.muted)
        report_frame.configure(bg=t.panel)
        gen_report_hover.configure(bg=t.panel)
        report_box.config(
            bg=t.field,
            fg=t.text,
            insertbackground=t.text,
            highlightbackground=t.border,
            highlightcolor=t.accent,
        )
        report_scroll.config(bg=t.panel, troughcolor=t.field, activebackground=t.accent)
        stt_lang_lbl.configure(bg=t.panel, fg=t.muted)
        if ttk is not None and _jw_style is not None:
            _jw_style.configure("Journal.TCombobox", **t.ttk_combobox_kwargs())
            if t.is_dark:
                _jw_style.map(
                    "Journal.TCombobox",
                    fieldbackground=[
                        ("readonly", t.field),
                        ("disabled", t.btn_disabled),
                    ],
                    selectbackground=[("readonly", t.accent)],
                    selectforeground=[("readonly", "white")],
                )
            else:
                _jw_style.map(
                    "Journal.TCombobox",
                    fieldbackground=[
                        ("readonly", t.field),
                        ("disabled", t.btn_disabled),
                    ],
                )
        else:
            try:
                lang_combo.config(bg=t.panel, fg=t.text)
            except tk.TclError:
                pass
        if str(gen_button.cget("state")) == "normal":
            _gs, gb, gf, gab, gaf = t.gen_bind_rest()
            gen_button.config(
                bg=gb,
                fg=gf,
                activebackground=gab,
                activeforeground=gaf,
                cursor="hand2",
            )
        else:
            _ds, gb, gf, _dab, _daf = t.gen_bind_disabled()
            gen_button.config(
                bg=gb,
                fg=gf,
                disabledforeground=gf,
                cursor="arrow",
            )
        update_transcribe_ui()
        for _b in (start_rec_button, pause_rec_button, stop_rec_button):
            _journal_rec_btn_set(_b, str(_b.cget("state")) == "normal")
        button_row.configure(bg=t.surface)
        button_row.pack_configure(padx=t.pad_outer, pady=(0, t.pad_button_y))
        refresh_save_entry_state()
        redraw_waveform_canvas()
        apply_journal_window_i18n()

    ui_lang_var.trace_add("write", lambda *_a: root.after_idle(_on_ui_language_selected))

    _startup_step("splash.detail.pages")

    def toggle_journal_window_theme() -> None:
        prefs = load_preferences()
        cur = normalize_journal_window_theme_key(th().id)
        nxt = "dark" if cur == "light" else "light"
        prefs[JOURNAL_PREF_THEME_KEY] = nxt
        save_preferences(prefs)
        theme_holder[0] = JOURNAL_THEME_DARK if nxt == "dark" else JOURNAL_THEME_LIGHT
        apply_journal_window_colors()

    theme_toggle_btn = tk.Button(
        top,
        text=t_init.toggle_label,
        command=toggle_journal_window_theme,
        bg=_ut_bg,
        fg=_ut_fg,
        activebackground=_ut_abg,
        activeforeground=_ut_afg,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=10,
        pady=6,
        cursor="hand2",
    )
    theme_toggle_btn.grid(row=0, column=6, sticky="e", padx=(8, 12), pady=12)
    bind_button_hover_if_enabled(
        theme_toggle_btn,
        lambda: (
            "normal",
            th().btn_secondary,
            th().text,
            th().secondary_hover,
            th().text,
        ),
        lambda: th().secondary_hover,
        lambda: th().text,
    )

    nav_specs: List[Tuple[str, str]] = [
        ("journal", "Journal"),
        ("ai_recap", "AI Recap"),
        ("chatbot", "Chatbot"),
        ("console", "Console"),
    ]
    for _idx, (_key, _label) in enumerate(nav_specs, start=1):
        _btn = tk.Button(
            nav_rail,
            text=_label,
            command=lambda k=_key: show_page(k),
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            padx=10,
            pady=8,
            anchor="w",
            cursor="hand2",
        )
        _btn.grid(row=_idx, column=0, sticky="ew", padx=10, pady=(0, 8))
        nav_buttons[_key] = _btn
        bind_button_hover_if_enabled(
            _btn,
            lambda b=_btn, k=_key: (
                "normal",
                th().accent if active_page["key"] == k else th().btn_secondary,
                "white" if active_page["key"] == k else th().text,
                th().hover_primary if active_page["key"] == k else th().secondary_hover,
                "white" if active_page["key"] == k else th().text,
            ),
            lambda k=_key: th().hover_primary if active_page["key"] == k else th().secondary_hover,
            lambda k=_key: "white" if active_page["key"] == k else th().text,
        )

    nav_settings_btn = tk.Button(
        nav_rail,
        text="Settings",
        command=lambda: show_page("settings"),
        bg=t_init.btn_secondary,
        fg=t_init.text,
        activebackground=t_init.secondary_hover,
        activeforeground=t_init.text,
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=10,
        pady=8,
        cursor="hand2",
        bd=0,
        highlightthickness=0,
    )
    nav_settings_btn.grid(row=101, column=0, sticky="e", padx=10, pady=(0, 10))
    nav_buttons["settings"] = nav_settings_btn
    bind_button_hover_if_enabled(
        nav_settings_btn,
        lambda: (
            "normal",
            th().btn_secondary,
            th().text,
            th().secondary_hover,
            th().text,
        ),
        lambda: th().secondary_hover,
        lambda: th().text,
    )

    nav_summon_btn.config(command=lambda: set_nav_visible(True))
    def _on_content_host_configure(_e: Optional[Any] = None) -> None:
        if nav_collapsed["value"] and not nav_animating["value"]:
            _place_nav_summon()
        frame = active_page_frame.get("frame")
        if frame is not None:
            _layout_console_row(frame)

    content_host.bind("<Configure>", _on_content_host_configure, add="+")
    bind_button_hover_if_enabled(
        nav_summon_btn,
        lambda: (
            "normal",
            th().toolbar_btn_config()[0],
            th().toolbar_btn_config()[1],
            th().toolbar_btn_config()[2],
            th().toolbar_btn_config()[3],
        ),
        lambda: th().toolbar_hover()[0],
        lambda: th().toolbar_hover()[1],
    )

    console_entry.bind("<Return>", lambda _e: (run_console_command(), "break")[1], add="+")
    console_entry.bind("<Up>", _console_history_up, add="+")
    console_entry.bind("<Down>", _console_history_down, add="+")
    console_entry.bind("<Tab>", _console_tab_complete, add="+")
    console_entry.bind("<KeyPress>", _console_entry_keypress, add="+")
    console_entry.bind("<FocusIn>", _console_entry_focus_in, add="+")
    console_entry.bind("<FocusOut>", _console_entry_focus_out, add="+")
    console_prompt.bind("<Button-1>", lambda _e: run_console_command(), add="+")

    def _unfocus_console_on_button_click(evt: Optional[Any] = None) -> None:
        if evt is None:
            return
        w = getattr(evt, "widget", None)
        if isinstance(w, tk.Button) and w is not console_entry:
            root.focus_set()

    root.bind_all("<Button-1>", _unfocus_console_on_button_click, add="+")
    _startup_step("splash.detail.finalize")
    show_page("journal")

    def _on_escape(event=None) -> None:
        if str(find_row.winfo_manager()) == "pack":
            _find_close()
            return
        on_close(event)

    root.bind("<Escape>", _on_escape)
    root.protocol("WM_DELETE_WINDOW", on_close)
    _startup_step("splash.detail.journal_ready")
    startup_overlay.destroy()
    root.lift()
    apply_journal_window_colors()

    def _background_post_init() -> None:
        _startup_step("splash.detail.other_pages")
        set_nav_visible(True)
        _startup_step("splash.detail.autosave")
        autosave()
        refresh_save_entry_state()
        _startup_step("splash.detail.ready")

    root.after(1, _background_post_init)
    root.mainloop()
    return saved["value"]


def maybe_prompt_startup_on_first_run() -> None:
    prefs = load_preferences()
    if prefs.get("startup_prompt_done", "").lower() == "true":
        return
    print("Open logger automatically when computer starts? (y/N): ", end="")
    try:
        answer = input().strip().lower()
    except (EOFError, RuntimeError):
        # Windowed launches (pythonw / console=False EXE) may not provide stdin.
        prefs["startup_enabled"] = "true" if is_startup_enabled() else "false"
        prefs["startup_prompt_done"] = "true"
        save_preferences(prefs)
        return
    if answer in ("y", "yes"):
        if create_startup_shortcut():
            print("Startup enabled.")
            prefs["startup_enabled"] = "true"
        else:
            print("Could not enable startup shortcut.")
            prefs["startup_enabled"] = "false"
    else:
        prefs["startup_enabled"] = "false"
        print("Startup remains disabled.")
    prefs["startup_prompt_done"] = "true"
    if not save_preferences(prefs):
        print("Warning: could not save startup preference.")


def setup_first_time_preferences() -> str:
    prefs = load_preferences()
    if prefs.get("initial_setup_done", "").lower() == "true":
        app_name = prefs.get("app_name", "").strip()
        if app_name:
            return app_name
        return get_or_create_app_name()

    print("First time setup: use default settings? (y/N): ", end="")
    try:
        answer = input().strip().lower()
    except (EOFError, RuntimeError):
        # Non-interactive/windowed launch: apply safe defaults without prompting.
        app_name = prefs.get("app_name", "").strip() or "Daily Logger"
        prefs["app_name"] = app_name
        prefs["startup_enabled"] = "true" if is_startup_enabled() else "false"
        prefs["startup_prompt_done"] = "true"
        prefs["initial_setup_done"] = "true"
        if not save_preferences(prefs):
            print("Warning: could not save initial preferences.")
        return app_name
    use_default = answer in ("y", "yes")

    if use_default:
        prefs["app_name"] = "Daily Logger"
        if create_startup_shortcut():
            prefs["startup_enabled"] = "true"
            print("Default setup applied: startup enabled.")
        else:
            prefs["startup_enabled"] = "false"
            print("Default setup applied: could not enable startup.")
        prefs["startup_prompt_done"] = "true"
    else:
        prefs["app_name"] = prompt_for_app_name()
        print("Open logger automatically when computer starts? (y/N): ", end="")
        try:
            startup_answer = input().strip().lower()
        except (EOFError, RuntimeError):
            startup_answer = ""
        if startup_answer in ("y", "yes"):
            if create_startup_shortcut():
                prefs["startup_enabled"] = "true"
                print("Startup enabled.")
            else:
                prefs["startup_enabled"] = "false"
                print("Could not enable startup shortcut.")
        else:
            prefs["startup_enabled"] = "false"
            print("Startup remains disabled.")
        prefs["startup_prompt_done"] = "true"

    prefs["initial_setup_done"] = "true"
    if not save_preferences(prefs):
        print("Warning: could not save initial preferences.")
    return prefs.get("app_name", "Daily Logger")


def get_chat_completions_url() -> str:
    return os.getenv("OPENAI_CHAT_COMPLETIONS_URL", OPENAI_CHAT_COMPLETIONS_URL).strip()


def _image_mime_for_path(path: Path) -> str:
    suf = path.suffix.lower()
    if suf in (".jpg", ".jpeg"):
        return "image/jpeg"
    if suf == ".gif":
        return "image/gif"
    if suf == ".webp":
        return "image/webp"
    return "image/png"


def build_user_message_with_attachments(
    question: str,
    image_paths: List[Path],
    file_paths: List[Path],
) -> Dict[str, object]:
    """Build a user message with optional images (vision) and text file excerpts."""
    text_chunks: List[str] = []
    q = (question or "").strip()
    if q:
        text_chunks.append(q)
    for fp in file_paths:
        ctx, _resolved, err = load_recap_context_from_file(str(fp))
        if err:
            text_chunks.append(f"[Attachment {fp.name}: {err}]")
        elif ctx:
            label = str(fp.resolve())
            clip = ctx if len(ctx) <= 48000 else ctx[:48000] + "\n\n[Truncated attachment]"
            text_chunks.append(f"[Attached file: {label}]\n{clip}")
    combined_text = "\n\n".join(text_chunks).strip() or "(no text)"
    parts: List[Dict[str, object]] = [{"type": "text", "text": combined_text}]
    for img_path in image_paths:
        try:
            raw = img_path.read_bytes()
        except OSError as exc:
            parts.append({"type": "text", "text": f"[Image {img_path.name} unreadable: {exc}]"})
            continue
        try:
            mime = _image_mime_for_path(img_path)
            b64 = base64.b64encode(raw).decode("ascii")
            parts.append(
                {
                    "type": "image_url",
                    "image_url": {"url": f"data:{mime};base64,{b64}"},
                }
            )
        except Exception as exc:
            parts.append(
                {"type": "text", "text": f"[Image {img_path.name} could not attach: {exc}]"}
            )

    if len(parts) == 1 and parts[0].get("type") == "text":
        return {"role": "user", "content": combined_text}
    return {"role": "user", "content": parts}


def chat_completion(
    messages: List[Dict[str, object]],
    model: str = OPENAI_MODEL,
    reasoning_effort: Optional[str] = None,
) -> str:
    api_key = get_openai_api_key()
    if not api_key:
        return "OPENAI_API_KEY is not set. Set it, then try again."

    chat_url = get_chat_completions_url()
    payload_data: Dict[str, object] = {
        "model": model,
        "messages": messages,
    }
    if reasoning_effort:
        payload_data["reasoning_effort"] = reasoning_effort
    payload = json.dumps(payload_data).encode("utf-8")
    req = request.Request(
        chat_url,
        data=payload,
        method="POST",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
    )

    body = None
    last_exception: Optional[Exception] = None
    for attempt in range(3):
        try:
            with request.urlopen(req, timeout=60) as response:
                body = response.read().decode("utf-8")
                break
        except error.HTTPError as exc:
            try:
                details = exc.read().decode("utf-8")
            except Exception:
                details = str(exc)
            return f"ChatGPT API error ({exc.code}): {details}"
        except Exception as exc:
            last_exception = exc
            if attempt < 2:
                time.sleep(2 * (attempt + 1))
            continue

    if body is None:
        return (
            "Failed to contact ChatGPT API after retries. "
            f"Last error: {last_exception}. "
            f"URL: {chat_url}. "
            "Check internet, firewall, VPN, or proxy settings."
        )

    try:
        parsed = json.loads(body)
        return parsed["choices"][0]["message"]["content"].strip()
    except Exception:
        return "ChatGPT returned an unexpected response format."


def chat_completion_with_spinner(
    messages: List[Dict[str, object]],
    model: str = OPENAI_MODEL,
    reasoning_effort: Optional[str] = None,
) -> str:
    holder: Dict[str, str] = {}

    def worker() -> None:
        holder["response"] = chat_completion(messages, model=model, reasoning_effort=reasoning_effort)

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()

    spinner = ["-", "\\", "|", "/"]
    spinner_colors = [
        "\033[31m",  # red
        "\033[33m",  # yellow
        "\033[32m",  # green
        "\033[36m",  # cyan
        "\033[34m",  # blue
        "\033[35m",  # magenta
    ]
    color_index = 0
    color_enabled = False
    index = 0
    cancelled = False
    start_time = time.time()
    print("AI is thinking... (press Enter to cancel)", end="", flush=True)
    while thread.is_alive():
        if msvcrt is not None and msvcrt.kbhit():
            char = msvcrt.getwch()
            if char in ("\r", "\n"):
                cancelled = True
                break
            if char == " ":
                color_enabled = True
                color_index = (color_index + 1) % len(spinner_colors)
        elapsed = time.time() - start_time
        spinner_char = spinner[index % len(spinner)]
        if color_enabled:
            spinner_char = f"{spinner_colors[color_index]}{spinner_char}\033[0m"
        sys.stdout.write(
            f"\rAI is thinking... {spinner_char} ({elapsed:.1f}s, press Enter to cancel)"
        )
        sys.stdout.flush()
        index += 1
        time.sleep(0.12)

    sys.stdout.write("\r" + (" " * 80) + "\r")
    sys.stdout.flush()
    if cancelled:
        return "Response cancelled by user."
    thread.join()
    return holder.get("response", "No response received.")


def print_chat_help() -> None:
    print("Chat commands:")
    print("  help - show this help")
    print("  ts   - take screenshot and attach to next AI message")
    print("  rs   - remove pending screenshot attachment")
    print("  Tab  - complete help / ts / rs; empty line + Tab shows this help")
    print("  Enter on empty line - exit chat")


def take_chat_screenshot_hidden_console() -> Optional[Path]:
    SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = SCREENSHOT_DIR / f"chat_{timestamp}.png"
    console_hwnd = None
    try:
        console_hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    except Exception:
        console_hwnd = None

    if console_hwnd:
        try:
            ctypes.windll.user32.ShowWindow(console_hwnd, 0)  # SW_HIDE
            time.sleep(0.6)
        except Exception:
            console_hwnd = None

    try:
        import mss.tools
        from mss import MSS

        with MSS() as sct:
            monitor = sct.monitors[0]
            shot = sct.grab(monitor)
            mss.tools.to_png(shot.rgb, shot.size, output=str(output_path))
        return output_path
    except Exception as exc:
        print(f"Could not capture screenshot: {exc}")
        return None
    finally:
        if console_hwnd:
            try:
                ctypes.windll.user32.ShowWindow(console_hwnd, 5)  # SW_SHOW
            except Exception:
                pass


def build_user_message(question: str, screenshot_path: Optional[Path]) -> Dict[str, object]:
    if screenshot_path is None:
        return {"role": "user", "content": question}
    try:
        image_b64 = base64.b64encode(screenshot_path.read_bytes()).decode("ascii")
    except OSError as exc:
        print(f"Could not read screenshot for attachment: {exc}")
        return {"role": "user", "content": question}
    return {
        "role": "user",
        "content": [
            {"type": "text", "text": question},
            {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_b64}"}},
        ],
    }


def run_chat_mode(
    with_journal_context: bool,
    use_thinking_model: bool = False,
    recap_date_range: Optional[Tuple[datetime, datetime]] = None,
    recap_context_override: Optional[str] = None,
    recap_context_label: Optional[str] = None,
) -> None:
    base_mode_label = "Recap" if with_journal_context else "Chatbot"
    if use_thinking_model and with_journal_context:
        base_mode_label = "Recap (Thinking)"
    if use_thinking_model and not with_journal_context:
        base_mode_label = "Chatbot(Thinking)"
    maybe_warn_for_current_wifi()
    if not ensure_openai_api_key_for_ai():
        return
    model_name = OPENAI_THINKING_MODEL if use_thinking_model else OPENAI_MODEL

    system_message = "You are a helpful assistant."
    if with_journal_context:
        if recap_context_override is not None:
            journal_context = recap_context_override
            if recap_context_label:
                print(f"Recap source: {recap_context_label}")
        else:
            journal_context = build_journal_context_for_range(recap_date_range)
        if recap_context_override is None and recap_date_range is not None:
            included_dates = list_journal_dates_in_range(recap_date_range)
            if included_dates:
                print("Recap includes journal dates:")
                print("  " + ", ".join(included_dates))
            else:
                print("Recap includes journal dates: (none found in selected range)")
        system_message = (
            "You answer questions only using the user's journal context. "
            "If the answer is not in the journal, say you do not know based on the journal."
        )
        context_message = f"Journal context:\n{journal_context}"
        messages: List[Dict[str, object]] = [
            {"role": "system", "content": system_message},
            {"role": "system", "content": context_message},
        ]
    else:
        messages = [{"role": "system", "content": system_message}]

    pending_screenshot: Optional[Path] = None

    def format_mode_label() -> str:
        if pending_screenshot is None:
            return base_mode_label
        if base_mode_label == "Chatbot(Thinking)":
            return "Chatbot(Thinking, Screenshot Attached)"
        if base_mode_label == "Chatbot":
            return "Chatbot (Screenshot Attached)"
        return base_mode_label

    print(f"\n=== {format_mode_label()} ===")
    if not with_journal_context:
        print('GPT: Hello, how can I help you? If you are stuck type "help"')
    while True:
        question_prompt = "Recap: " if with_journal_context else "You: "
        question = input_line_with_tab_completions(
            question_prompt, CHAT_LINE_COMPLETIONS, on_empty_tab=print_chat_help
        )
        if is_enter_equivalent(question):
            print(f"Leaving {base_mode_label}.")
            return
        if question.lower() == "help":
            print_chat_help()
            continue
        if not with_journal_context and question.lower() == "ts":
            print("Taking screenshot...")
            pending_screenshot = take_chat_screenshot_hidden_console()
            if pending_screenshot:
                print(f"Screenshot ready: {pending_screenshot}")
                print(f"=== {format_mode_label()} ===")
            continue
        if not with_journal_context and question.lower() == "rs":
            if pending_screenshot is None:
                print("No pending screenshot to remove.")
            else:
                pending_screenshot = None
                print("Pending screenshot removed.")
                print(f"=== {format_mode_label()} ===")
            continue

        had_screenshot = pending_screenshot is not None
        user_message = build_user_message(question, pending_screenshot)
        messages.append(user_message)
        pending_screenshot = None
        effort = "high" if use_thinking_model else None
        answer = chat_completion_with_spinner(messages, model=model_name, reasoning_effort=effort)
        print(f"GPT: {answer}")
        if had_screenshot:
            print(f"=== {format_mode_label()} ===")
        messages.append({"role": "assistant", "content": answer})


def is_row_empty(values: List[object]) -> bool:
    return all(value is None or (isinstance(value, str) and not value.strip()) for value in values)


def find_first_empty_data_row(ws, column_count: int) -> int:
    for row_index in range(2, ws.max_row + 1):
        values = [ws.cell(row=row_index, column=col).value for col in range(1, column_count + 1)]
        if is_row_empty(values):
            return row_index
    return ws.max_row + 1


def ask_entry_date_time() -> Optional[Tuple[str, str]]:
    now = datetime.now()
    default_date = now.strftime("%m/%d/%Y")

    while True:
        date_input = input(
            f"Entry date (mm/dd/yyyy, Enter for today {default_date}): "
        ).strip()
        if not date_input:
            date_value = default_date
            break
        if date_input.upper() == "X":
            confirm = input("Return to main menu? (y/N): ").strip().lower()
            if confirm in ("y", "yes"):
                return None
            continue
        parsed_date = parse_flexible_date(date_input, now.year)
        if parsed_date:
            date_value = parsed_date.strftime("%m/%d/%Y")
            break
        print("Invalid date. Try 04/20/2026, 4/26, Apr 26, or April 26.")

    while True:
        time_input = input(
            "Entry time (example: 11:00AM, type rn for now, Enter for N/A): "
        ).strip()
        if not time_input:
            time_value = "N/A"
            break
        if time_input.upper() == "X":
            confirm = input("Return to main menu? (y/N): ").strip().lower()
            if confirm in ("y", "yes"):
                return None
            continue
        if time_input.lower() == "rn":
            time_value = datetime.now().strftime("%I:%M%p").lstrip("0")
            break

        normalized = time_input.upper().replace(" ", "")
        try:
            parsed = datetime.strptime(normalized, "%I:%M%p")
            time_value = parsed.strftime("%I:%M%p").lstrip("0")
            break
        except ValueError:
            print("Invalid time format. Use hh:mmAM/PM (example: 2:03PM), or rn for current time.")

    return date_value, time_value


def ask_entry_date_time_gui(parent: Optional[Any] = None) -> Optional[Tuple[str, str]]:
    """
    GUI version of ``ask_entry_date_time()``.
    The CLI version uses ``input()`` which freezes the Tk main thread in the GUI console.
    """
    if tk is None:
        return ask_entry_date_time()

    _parent = parent if parent is not None else tk._default_root  # type: ignore[attr-defined]
    if _parent is None:
        # Last-resort fallback: no Tk root to attach to.
        return ask_entry_date_time()

    now = datetime.now()
    default_date = now.strftime("%m/%d/%Y")

    dlg = tk.Toplevel(_parent)
    dlg.title("Entry date & time")
    dlg.transient(_parent)

    # Make modal without blocking background threads; Tk will run its own nested event loop.
    result: Dict[str, Optional[Tuple[str, str]]] = {"v": None}
    dlg.grab_set()

    tk.Label(dlg, text=f"Entry date (mm/dd/yyyy, Enter for today {default_date}):").pack(
        padx=12, pady=(12, 4)
    )
    date_var = tk.StringVar(value=default_date)
    date_entry = tk.Entry(dlg, textvariable=date_var, width=22)
    date_entry.pack(padx=12, pady=(0, 8))

    tk.Label(dlg, text="Entry time (example: 11:00AM, type rn for now, Enter for N/A):").pack(
        padx=12, pady=(0, 4)
    )
    time_var = tk.StringVar(value="")
    time_entry = tk.Entry(dlg, textvariable=time_var, width=22)
    time_entry.pack(padx=12, pady=(0, 8))

    def _parse_date(date_input: str) -> Optional[str]:
        di = date_input.strip()
        if not di:
            return default_date
        if di.upper() == "X":
            return None
        parsed_date = parse_flexible_date(di, now.year)
        if parsed_date:
            return parsed_date.strftime("%m/%d/%Y")
        return None

    def _parse_time(time_input: str) -> Optional[str]:
        ti = time_input.strip()
        if not ti:
            return "N/A"
        if ti.upper() == "X":
            return None
        if ti.lower() == "rn":
            return datetime.now().strftime("%I:%M%p").lstrip("0")
        normalized = ti.upper().replace(" ", "")
        try:
            parsed = datetime.strptime(normalized, "%I:%M%p")
            return parsed.strftime("%I:%M%p").lstrip("0")
        except ValueError:
            return None

    def on_ok() -> None:
        date_input = date_var.get()
        time_input = time_var.get()
        parsed_date = _parse_date(date_input)
        if parsed_date is None:
            # "X" or invalid date -> treat X as cancel, invalid as error.
            if date_input.strip().upper() == "X":
                result["v"] = None
                dlg.destroy()
                return
            messagebox.showerror("Invalid date", "Enter a valid date like 04/20/2026 or April 26.")
            return

        parsed_time = _parse_time(time_input)
        if parsed_time is None:
            if time_input.strip().upper() == "X":
                result["v"] = None
                dlg.destroy()
                return
            messagebox.showerror(
                "Invalid time",
                "Enter time like 11:00AM or type rn for now, or leave blank for N/A.",
            )
            return

        result["v"] = (parsed_date, parsed_time)
        dlg.destroy()

    def on_cancel() -> None:
        result["v"] = None
        dlg.destroy()

    btn_row = tk.Frame(dlg)
    btn_row.pack(padx=12, pady=12)
    tk.Button(btn_row, text="OK", command=on_ok, width=10).pack(side="left", padx=(0, 8))
    tk.Button(btn_row, text="Cancel", command=on_cancel, width=10).pack(side="left")

    dlg.protocol("WM_DELETE_WINDOW", on_cancel)
    # Focus defaults for quick keyboard entry.
    try:
        date_entry.focus_set()
    except tk.TclError:
        pass
    dlg.wait_window()
    return result["v"]


def _ask_typed_note_gui(parent: Optional[Any] = None) -> Optional[str]:
    if tk is None:
        return None
    _parent = parent if parent is not None else tk._default_root  # type: ignore[attr-defined]
    if _parent is None:
        return None
    dlg = tk.Toplevel(_parent)
    dlg.title("What happened today?")
    dlg.transient(_parent)
    dlg.grab_set()

    result: Dict[str, Optional[str]] = {"v": None}
    tk.Label(dlg, text="Type what happened today:").pack(padx=12, pady=(12, 4))
    box = tk.Text(dlg, height=8, width=52)
    box.pack(padx=12, pady=(0, 8))

    def on_ok() -> None:
        result["v"] = box.get("1.0", "end-1c").strip()
        dlg.destroy()

    def on_cancel() -> None:
        result["v"] = None
        dlg.destroy()

    btn_row = tk.Frame(dlg)
    btn_row.pack(padx=12, pady=12)
    tk.Button(btn_row, text="OK", command=on_ok, width=10).pack(side="left", padx=(0, 8))
    tk.Button(btn_row, text="Cancel", command=on_cancel, width=10).pack(side="left")
    dlg.protocol("WM_DELETE_WINDOW", on_cancel)
    try:
        box.focus_set()
    except tk.TclError:
        pass
    dlg.wait_window()
    return result["v"]


def journal_settings_menu_gui(parent: Optional[Any] = None) -> Optional[List[str]]:
    """
    GUI replacement for ``journal_settings_menu()``.
    Mirrors the same choices, but avoids blocking ``input()`` used by the CLI menu.
    """
    if tk is None:
        return journal_settings_menu()

    _parent = parent if parent is not None else tk._default_root  # type: ignore[attr-defined]
    if _parent is None:
        return journal_settings_menu()

    result: Dict[str, Optional[List[str]]] = {"v": None}
    dlg = tk.Toplevel(_parent)
    dlg.title("Journal settings")
    dlg.transient(_parent)
    dlg.grab_set()

    def _show_help() -> None:
        help_text = (
            "Journal choices:\n"
            "  WINDOW               - open window editor\n"
            "  CONSOLE              - type journal text in console\n"
            "  EDITPREV             - edit latest entry in window\n"
            "  DP                   - delete latest entry\n"
            "  RESTORE              - reopen latest unsaved draft\n"
            "  HELP                 - show this list\n"
            "  Enter                - return to main menu\n"
            "  DEFAULT WINDOWS     - set preferred journal input to window\n"
            "  DEFAULT CONSOLE      - set preferred journal input to console\n"
        )
        messagebox.showinfo("Journal settings help", help_text)

    tk.Label(dlg, text="Journal choice (type HELP for options):").pack(padx=12, pady=(12, 4))
    cmd_var = tk.StringVar(value="")
    entry = tk.Entry(dlg, textvariable=cmd_var, width=44)
    entry.pack(padx=12, pady=(0, 10))

    # For multi-step flows (CONSOLE), keep a pointer so we can close the menu once they finish.
    def on_submit() -> None:
        note = (cmd_var.get() or "").strip()
        key = note.lower()

        if is_enter_equivalent(note.upper()):
            dlg.destroy()
            return
        if key == "help":
            _show_help()
            return

        if key in ("c", "console", "coinsole"):
            typed = _ask_typed_note_gui(dlg)
            if typed is None:
                dlg.destroy()
                return
            dt = ask_entry_date_time_gui(dlg)
            if dt is None:
                dlg.destroy()
                return
            date_value, time_value = dt
            result["v"] = [date_value, time_value, typed, "", ""]
            dlg.destroy()
            return

        if key in (
            "editprev",
            "edit prev",
            "edit previous",
            "openprev",
            "open prev",
            "openprevious",
            "open previous",
        ):
            latest = get_latest_journal_entry_for_edit()
            if not latest:
                messagebox.showinfo("Edit previous", "No previous journal entry found to edit.")
                return
            open_journal_window_editor(
                {
                    "text": str(latest.get("text", "")),
                    "speech_transcript": str(latest.get("speech_transcript", "")),
                    "ai_report": str(latest.get("ai_report", "")),
                    "date": str(latest.get("date", "")),
                    "time": str(latest.get("time", "")),
                    "images": [],
                    "edit_target_sheet": str(latest.get("sheet_name", "")),
                    "edit_target_row": int(latest.get("row_index", 0) or 0),
                }
            )
            dlg.destroy()
            return

        if key in ("w", "window", "windows"):
            open_journal_window_editor()
            dlg.destroy()
            return

        if key in ("default windows", "default console"):
            prefs = load_preferences()
            default_mode = "windows" if key.endswith("windows") else "console"
            prefs["journal_input_default"] = default_mode
            if save_preferences(prefs):
                messagebox.showinfo(
                    "Default updated",
                    f"Default journal input set to {default_mode}.",
                )
            dlg.destroy()
            return

        if key == "restore":
            draft = load_journal_window_draft()
            if not draft:
                messagebox.showinfo("Restore", "No journal draft to restore.")
                dlg.destroy()
                return
            open_journal_window_editor(draft)
            dlg.destroy()
            return

        if key.upper() == "DP":
            latest = get_latest_journal_entry_for_delete()
            if not latest:
                messagebox.showinfo("Delete previous", "No previous journal entry found to delete.")
                return
            date_label = str(latest.get("date", "")).strip() or "(unknown date)"
            time_label = str(latest.get("time", "")).strip() or "(unknown time)"
            should_delete = messagebox.askyesno(
                "Delete previous journal entry",
                f"Delete previous journal entry at {date_label} {time_label}?",
            )
            if should_delete:
                delete_latest_journal_entry()
                dlg.destroy()
            return

        messagebox.showerror("Unknown choice", "Unknown journal choice. Type HELP to see valid options.")

    def on_cancel() -> None:
        result["v"] = None
        dlg.destroy()

    entry.bind("<Return>", lambda _e: on_submit())
    btn_row = tk.Frame(dlg)
    btn_row.pack(padx=12, pady=(0, 12))
    tk.Button(btn_row, text="Submit", command=on_submit, width=10).pack(side="left", padx=(0, 8))
    tk.Button(btn_row, text="Cancel", command=on_cancel, width=10).pack(side="left")

    dlg.protocol("WM_DELETE_WINDOW", on_cancel)
    try:
        entry.focus_set()
    except tk.TclError:
        pass
    dlg.wait_window()
    return result["v"]


def parse_flexible_date(raw: str, default_year: int):
    cleaned = " ".join(raw.strip().split())
    if not cleaned:
        return None

    slash_parts = cleaned.split("/")
    if len(slash_parts) in (2, 3):
        try:
            month = int(slash_parts[0])
            day = int(slash_parts[1])
            if len(slash_parts) == 3:
                year = int(slash_parts[2])
                if year < 100:
                    year += 2000
            else:
                year = default_year
            return datetime(year, month, day)
        except ValueError:
            pass

    text_parts = cleaned.replace(",", "").split()
    if len(text_parts) in (2, 3):
        month_token = text_parts[0].lower()
        month_map = {
            "jan": 1,
            "january": 1,
            "feb": 2,
            "february": 2,
            "mar": 3,
            "march": 3,
            "apr": 4,
            "april": 4,
            "may": 5,
            "jun": 6,
            "june": 6,
            "jul": 7,
            "july": 7,
            "aug": 8,
            "august": 8,
            "sep": 9,
            "sept": 9,
            "september": 9,
            "oct": 10,
            "october": 10,
            "nov": 11,
            "november": 11,
            "dec": 12,
            "december": 12,
        }
        month = month_map.get(month_token)
        if month is not None:
            try:
                day = int(text_parts[1])
                if len(text_parts) == 3:
                    year = int(text_parts[2])
                    if year < 100:
                        year += 2000
                else:
                    year = default_year
                return datetime(year, month, day)
            except ValueError:
                pass

    return None


def sync_journal_workbook() -> None:
    module = MODULES["J"]
    workbook_path = ensure_workbook(module)
    wb = load_workbook_with_retry(workbook_path)
    rebuild_master_journal_from_daily_pages(wb, module)
    reorder_journal_sheets(wb)
    save_workbook_with_retry(wb, workbook_path)


def is_journal_workbook_write_locked() -> bool:
    journal_path = DATA_DIR / MODULES["J"].workbook_name
    if not journal_path.exists():
        return False
    try:
        with open(journal_path, "r+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        return False


def journal_settings_menu() -> Optional[List[str]]:
    def print_journal_choice_help() -> None:
        print("Journal choices:")
        print("  WINDOW               - open window editor")
        print("  COINSOLE             - type journal text in console")
        print("  EDITPREV             - edit latest entry in window")
        print("  DP                   - delete latest entry")
        print("  RESTORE              - reopen latest unsaved draft")
        print("  HELP                 - show this list")
        print("  Enter                - return to main menu")
        print("  DEFAULT WINDOWS      - set preferred journal input to window")
        print("  DEFAULT CONSOLE      - set preferred journal input to console")

    while True:
        print_journal_choice_help()
        note = input_line_with_tab_completions(
            "Journal choice: ",
            (
                "help",
                "c",
                "console",
                "coinsole",
                "dp",
                "w",
                "window",
                "windows",
                "editprev",
                "edit previous",
                "openprev",
                "open previous",
                "restore",
                "default windows",
                "default console",
            ),
        )
        if is_enter_equivalent(note):
            return None
        if note.lower() == "help":
            print_journal_choice_help()
            continue
        if note.lower() in ("c", "console", "coinsole"):
            typed_note = input("What happened today? ").strip()
            if is_enter_equivalent(typed_note):
                return None
            date_time = ask_entry_date_time()
            if date_time is None:
                return None
            date_value, time_value = date_time
            return [date_value, time_value, typed_note, "", ""]
        if note.lower() in (
            "editprev",
            "edit prev",
            "edit previous",
            "openprev",
            "open prev",
            "openprevious",
            "open previous",
        ):
            latest = get_latest_journal_entry_for_edit()
            if not latest:
                print("No previous journal entry found to edit.")
                return None
            open_journal_window_editor(
                {
                    "text": str(latest.get("text", "")),
                    "speech_transcript": str(latest.get("speech_transcript", "")),
                    "ai_report": str(latest.get("ai_report", "")),
                    "date": str(latest.get("date", "")),
                    "time": str(latest.get("time", "")),
                    "images": [],
                    "edit_target_sheet": str(latest.get("sheet_name", "")),
                    "edit_target_row": int(latest.get("row_index", 0) or 0),
                }
            )
            return None
        if note.lower() in ("w", "window", "windows"):
            open_journal_window_editor()
            return None
        if note.lower() in ("default windows", "default console"):
            prefs = load_preferences()
            default_mode = "windows" if note.lower().endswith("windows") else "console"
            prefs["journal_input_default"] = default_mode
            if save_preferences(prefs):
                if default_mode == "windows":
                    print("Default set to windows. Typing J opens the window editor.")
                else:
                    print("Default set to console. Typing J shows journal choices.")
            else:
                print("Could not save default journal input preference.")
            return None
        if note.lower() == "restore":
            draft = load_journal_window_draft()
            if not draft:
                print("No journal draft to restore.")
                return None
            restored = open_journal_window_editor(draft)
            if restored:
                print("Restored draft saved.")
            else:
                print("Draft restore opened. Unsaved draft remains available.")
            return None
        if note.upper() == "DP":
            latest = get_latest_journal_entry_for_delete()
            if not latest:
                print("No previous journal entry found to delete.")
                return None
            date_label = str(latest.get("date", "")).strip() or "(unknown date)"
            time_label = str(latest.get("time", "")).strip() or "(unknown time)"
            text_label = str(latest.get("text", "")).strip()
            while True:
                confirm = input(
                    f'Delete previous journal entry at {date_label} {time_label}? (y/N or type "expand"): '
                ).strip().lower()
                if confirm == "expand":
                    if text_label:
                        print("Entry text:")
                        print(text_label)
                    else:
                        print("(Entry text is empty.)")
                    continue
                if confirm in ("y", "yes"):
                    deleted = delete_latest_journal_entry()
                    if deleted:
                        print("Previous journal entry deleted.")
                    else:
                        print("No previous journal entry found to delete.")
                    break
                print("Delete cancelled.")
                break
            return None
        print(
            "Unknown journal choice. Type HELP to see commands, or use C/CONSOLE to write in console."
        )


def journal_prompts() -> Optional[List[str]]:
    prefs = load_preferences()
    default_mode = prefs.get("journal_input_default", "").strip().lower() or "windows"
    if default_mode == "windows":
        open_journal_window_editor()
        return None
    if default_mode == "console":
        typed_note = input("What happened today? ").strip()
        if is_enter_equivalent(typed_note):
            return None
        date_time = ask_entry_date_time()
        if date_time is None:
            return None
        date_value, time_value = date_time
        return [date_value, time_value, typed_note, "", ""]
    return journal_settings_menu()


MODULES: Dict[str, ModuleConfig] = {
    "J": ModuleConfig(
        name="Journal",
        workbook_name="Journal.xlsx",
        sheet_name="Journal",
        headers=["Date", "Time", "Journal", "Speech to text", "AI report"],
        prompt_builder=journal_prompts,
    ),
}


def print_main_help() -> None:
    print("Main commands:")
    print("  J      - Journal")
    print("  R      - Recap")
    print("  RT     - Recap (thinking)")
    print("  R [date range] / RT [date range] - recap only within date range")
    print("      Examples: 4.27 4.30 | 4/27 4/30 | 4/27 - 4/30 | 4/27/2026 - 4/30/2026")
    print("  R [file] / RT [file] - recap using file text as context")
    print("      Examples: R notes.txt | RT daily_logs/meeting.md")
    print("  C      - Chatbot")
    print("  CT     - Chatbot (thinking)")
    print("  H/HELP - show this help")
    print("  J SETTINGS / J SETTING / JOURNAL SETTINGS / JS - open journal command menu")
    print("  RENAME - change app name")
    print("  STARTUP TRUE  - enable startup shortcut")
    print("  STARTUP FALSE - disable startup shortcut")
    print("  DEFAULT WINDOWS - typing J opens journal window directly")
    print("  DEFAULT CONSOLE - typing J shows journal command choices")
    print("  OPEN DIRECTORY   - open app data folder")
    print("  OPEN JOURNAL     - open Journal.xlsx")
    print("  OPEN SCREENSHOTS - open chat_screenshots folder")
    print("  DIRECTOR OPEN - open app data folder in File Explorer")
    print("  BACKUP START   - create backup zip in daily_logs/backup")
    print("  BACKUP TRUE    - auto backup once on each new day (default)")
    print("  BACKUP FALSE   - disable auto backup")
    print("  BACKUP LIMITED - keep at most 3 zip files; remove latest when adding")
    print("  UNINSTALL - request uninstall (requires CONFIRM UNINSTALL)")
    print("  CONFIRM UNINSTALL - permanently remove app data/app files")
    print("  WIFI WARN [name] - warn when connected to that Wi-Fi")
    print("  RESTORE - reopen latest unsaved journal window draft")
    print("  TOKEN ADD [token] - save API token")
    print("  TOKEN RESET - delete saved API token")
    print("  TOKEN COPY - copy current API token")
    print("  LAN cn | LAN en | LANGUAGE Chinese | LANGUAGE English - UI language")
    print("  SB bat     - Start Menu shortcut so Windows Search finds the .bat launcher")
    print("  SB journal - Start Menu shortcut so Windows Search finds Journal.xlsx")
    print("  Enter  - Continue/Exit")
    print("  X      - Exit")
    print("  TS     - take screenshot now (not attached outside chat mode)")
    print("  Tab    - complete a command; empty line + Tab shows this help")


def print_menu(app_name: str) -> None:
    print(f"\n=== {app_name} ===")
    print("Select an option below:")
    has_api_key = get_openai_api_key() is not None
    recap_label = "R = AI Recap" if has_api_key else "R = AI Recap (No API Key)"
    chat_label = "C = Chatbot" if has_api_key else "C = Chatbot (No API Key)"
    print("J = Journal")
    print(recap_label)
    print(chat_label)
    print("H = Commands")
    print("Enter = Skip/Exit")


def handle_choice(choice: str, app_name: str) -> Tuple[bool, str]:
    global PENDING_UNINSTALL_CONFIRM
    raw = choice.strip()
    key = raw.upper()
    if is_enter_equivalent(key):
        print("Skipped. See you next time.")
        return False, app_name
    if key in ("H", "HELP"):
        print_main_help()
        return True, app_name
    if key == "UNINSTALL":
        PENDING_UNINSTALL_CONFIRM = True
        print('Uninstall requested. Type "CONFIRM UNINSTALL" to continue.')
        return True, app_name
    if key == "CONFIRM UNINSTALL":
        if not PENDING_UNINSTALL_CONFIRM:
            print('Type "UNINSTALL" first, then "CONFIRM UNINSTALL".')
            return True, app_name
        run_clean_uninstall()
        return False, app_name
    if key in ("J SETTINGS", "J SETTING", "JOURNAL SETTINGS", "JS"):
        values = journal_settings_menu()
        if values is None:
            if load_journal_window_draft():
                print("Draft saved without journal entry. Use RESTORE to reopen it.")
            return True, app_name
        append_row(MODULES["J"], values)
        print(f'Journal saved to: {DATA_DIR / MODULES["J"].workbook_name}')
        return True, app_name
    if key == "X":
        print("Exit requested.")
        return False, app_name
    if key == "RENAME":
        app_name = rename_app_name()
        return True, app_name
    if key.startswith("RENAME "):
        app_name = rename_app_name_to(raw[7:].strip())
        return True, app_name
    if key.startswith("REANAME "):
        app_name = rename_app_name_to(raw[8:].strip())
        return True, app_name
    if key.startswith("WIFI WARN "):
        wifi_name = raw[10:].strip()
        if not wifi_name:
            print("Usage: wifi warn [wifi name]")
            return True, app_name
        if add_wifi_warn_name(wifi_name):
            print(f'Wi-Fi warning added for "{wifi_name}".')
        else:
            print("Could not save Wi-Fi warning list.")
        return True, app_name
    if key == "RESTORE":
        draft = load_journal_window_draft()
        if not draft:
            print("No journal draft to restore.")
            return True, app_name
        restored = open_journal_window_editor(draft)
        if restored:
            print("Restored draft saved.")
        else:
            print("Draft restore opened. Unsaved draft remains available.")
        return True, app_name
    _menu_parts = raw.split(None, 1)
    if _menu_parts and _menu_parts[0].upper() in ("LAN", "LANGUAGE"):
        _arg = _menu_parts[1].strip() if len(_menu_parts) > 1 else ""
        if not _arg:
            print("Usage: LAN cn | LAN en | LANGUAGE Chinese | LANGUAGE English")
            return True, app_name
        new_lang = normalize_ui_language(_arg)
        prefs = load_preferences()
        cur = normalize_ui_language(str(prefs.get(UI_LANGUAGE_PREF_KEY, "en")))
        if new_lang == cur:
            print(f"UI language is already {'Chinese' if new_lang == 'zh' else 'English'}.")
            return True, app_name
        prefs[UI_LANGUAGE_PREF_KEY] = new_lang
        if not save_preferences(prefs):
            print("Could not save language preference.")
            return True, app_name
        hook = _journal_ui_language_changed_hook
        if hook is not None:
            hook(new_lang)
        print(f"UI language set to {'Chinese' if new_lang == 'zh' else 'English'}.")
        return True, app_name
    if key.startswith("TOKEN ADD "):
        token_value = raw[10:].strip()
        if not token_value:
            print("Usage: token add [token]")
            return True, app_name
        if save_openai_api_key(token_value):
            print("API token saved.")
        else:
            print("Could not save API token.")
        return True, app_name
    if key == "TOKEN RESET":
        confirm = input("Are you sure you want to delete saved API token? (y/N): ").strip().lower()
        if confirm in ("y", "yes"):
            if delete_openai_api_key():
                print("Saved API token deleted.")
            else:
                print("Could not delete saved API token.")
        else:
            print("Token reset cancelled.")
        return True, app_name
    if key == "TOKEN COPY":
        token_value = get_openai_api_key()
        if not token_value:
            print("No current API token found.")
            return True, app_name
        if copy_text_to_clipboard(token_value):
            print("Current API token copied to clipboard.")
        else:
            print("Could not copy token to clipboard.")
        return True, app_name
    if key == "STARTUP TRUE":
        if is_startup_enabled():
            print("Startup is already enabled.")
            return True, app_name
        if create_startup_shortcut():
            print("Startup enabled.")
            prefs = load_preferences()
            prefs["startup_enabled"] = "true"
            prefs["startup_prompt_done"] = "true"
            save_preferences(prefs)
        else:
            print("Could not enable startup shortcut.")
        return True, app_name
    if key == "STARTUP FALSE":
        if remove_startup_shortcut():
            print("Startup disabled.")
            prefs = load_preferences()
            prefs["startup_enabled"] = "false"
            prefs["startup_prompt_done"] = "true"
            save_preferences(prefs)
        else:
            print("Could not disable startup shortcut.")
        return True, app_name
    if key == "DEFAULT WINDOWS":
        prefs = load_preferences()
        prefs["journal_input_default"] = "windows"
        if save_preferences(prefs):
            print("Default set to windows. Typing J opens the window editor.")
        else:
            print("Could not save default journal input preference.")
        return True, app_name
    if key == "DEFAULT CONSOLE":
        prefs = load_preferences()
        prefs["journal_input_default"] = "console"
        if save_preferences(prefs):
            print("Default set to console. Typing J shows journal choices.")
        else:
            print("Could not save default journal input preference.")
        return True, app_name
    if key == "DIRECTOR OPEN":
        if open_current_directory_in_explorer():
            print(f"Opened folder: {USER_DATA_ROOT}")
        else:
            print("Could not open current folder in File Explorer.")
        return True, app_name
    if key.startswith("OPEN "):
        open_target = raw[5:].strip().upper()
        if open_target == "DIRECTORY":
            if open_current_directory_in_explorer():
                print(f"Opened folder: {USER_DATA_ROOT}")
            else:
                print("Could not open current folder in File Explorer.")
            return True, app_name
        if open_target == "JOURNAL":
            journal_path = ensure_workbook(MODULES["J"])
            if open_path_with_default_app(journal_path):
                print(f"Opened journal file: {journal_path}")
            else:
                print("Could not open Journal.xlsx.")
            return True, app_name
        if open_target == "SCREENSHOTS":
            SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
            if open_path_with_default_app(SCREENSHOT_DIR):
                print(f"Opened screenshots folder: {SCREENSHOT_DIR}")
            else:
                print("Could not open screenshots folder.")
            return True, app_name
        print("Usage: OPEN DIRECTORY | OPEN JOURNAL | OPEN SCREENSHOTS")
        return True, app_name
    if key == "TS":
        print("Taking screenshot...")
        screenshot_path = take_chat_screenshot_hidden_console()
        if screenshot_path:
            print(f"Screenshot saved: {screenshot_path}")
        return True, app_name
    if key == "BACKUP START":
        prefs = load_preferences()
        evict_oldest_backup_if_limited_full(prefs)
        backup_path = run_backup_now()
        if backup_path is None:
            print("No files/folders in daily_logs to back up.")
            return True, app_name
        trim_backups_if_limited(prefs)
        prefs["last_backup_date"] = datetime.now().strftime("%Y-%m-%d")
        save_preferences(prefs)
        print(f"Backup created: {backup_path}")
        return True, app_name
    if key == "BACKUP TRUE":
        prefs = load_preferences()
        prefs["backup_enabled"] = "true"
        if save_preferences(prefs):
            print("Auto backup enabled.")
        else:
            print("Could not save backup preference.")
        return True, app_name
    if key == "BACKUP FALSE":
        prefs = load_preferences()
        prefs["backup_enabled"] = "false"
        if save_preferences(prefs):
            print("Auto backup disabled.")
        else:
            print("Could not save backup preference.")
        return True, app_name
    if key == "BACKUP LIMITED":
        prefs = load_preferences()
        prefs["backup_limited"] = "true"
        trim_backups_if_limited(prefs)
        if save_preferences(prefs):
            print("Backup limited mode enabled (max 3 zip files).")
        else:
            print("Could not save backup limit preference.")
        return True, app_name
    if key.startswith("SB "):
        sub = raw[3:].strip().upper()
        if sub == "BAT":
            if not sb_create_bat_search_shortcut():
                print("Could not create BAT search shortcut.")
        elif sub == "JOURNAL":
            if not sb_create_journal_search_shortcut():
                print("Could not create Journal search shortcut.")
        else:
            print('Usage: SB bat   or   SB journal')
        return True, app_name
    if key.startswith("RT "):
        recap_range, file_context, file_path, recap_err = resolve_recap_target(
            raw[3:].strip(), datetime.now().year
        )
        if recap_err is not None:
            print(recap_err)
            return True, app_name
        run_chat_mode(
            with_journal_context=True,
            use_thinking_model=True,
            recap_date_range=recap_range,
            recap_context_override=file_context,
            recap_context_label=file_path,
        )
        return True, app_name
    if key.startswith("R "):
        recap_range, file_context, file_path, recap_err = resolve_recap_target(
            raw[2:].strip(), datetime.now().year
        )
        if recap_err is not None:
            print(recap_err)
            return True, app_name
        run_chat_mode(
            with_journal_context=True,
            recap_date_range=recap_range,
            recap_context_override=file_context,
            recap_context_label=file_path,
        )
        return True, app_name
    if key == "R":
        run_chat_mode(with_journal_context=True)
        return True, app_name
    if key == "RT":
        run_chat_mode(with_journal_context=True, use_thinking_model=True)
        return True, app_name
    if key == "C":
        run_chat_mode(with_journal_context=False)
        return True, app_name
    if key == "CT":
        run_chat_mode(with_journal_context=False, use_thinking_model=True)
        return True, app_name

    if key == "J" and is_journal_workbook_write_locked():
        print("Journal is currently open in another program.")
        input("Close Journal.xlsx and press Enter to return to main menu...")
        return True, app_name

    module = MODULES.get(key)
    if not module:
        print(
            "Unknown choice. Please enter J, J SETTINGS, J SETTING, JOURNAL SETTINGS, JS, R, RT, C, CT, H, HELP, RENAME, STARTUP TRUE/FALSE, DEFAULT WINDOWS/CONSOLE, OPEN DIRECTORY/JOURNAL/SCREENSHOTS, DIRECTOR OPEN, BACKUP START/TRUE/FALSE/LIMITED, TS, UNINSTALL, CONFIRM UNINSTALL, SB bat/journal, WIFI WARN [name], RESTORE, LAN/LANGUAGE, TOKEN ADD/RESET/COPY, or press Enter to skip."
        )
        return True, app_name

    values = module.prompt_builder()
    if values is None:
        if key == "J" and load_journal_window_draft():
            print("Draft saved without journal entry. Use RESTORE to reopen it.")
        return True, app_name
    append_row(module, values)
    print(f"{module.name} saved to: {DATA_DIR / module.workbook_name}")
    return True, app_name


# Full-line menu strings for Tab completion (canonical spelling).
MAIN_MENU_COMPLETIONS: Tuple[str, ...] = tuple(
    sorted(
        {
            "J",
            "J SETTINGS",
            "J SETTING",
            "JOURNAL SETTINGS",
            "JS",
            "R",
            "R ",
            "RT",
            "RT ",
            "C",
            "CT",
            "H",
            "HELP",
            "X",
            "RENAME",
            "RESTORE",
            "STARTUP TRUE",
            "STARTUP FALSE",
            "DEFAULT WINDOWS",
            "DEFAULT CONSOLE",
            "OPEN DIRECTORY",
            "OPEN JOURNAL",
            "OPEN SCREENSHOTS",
            "DIRECTOR OPEN",
            "BACKUP START",
            "BACKUP TRUE",
            "BACKUP FALSE",
            "BACKUP LIMITED",
            "UNINSTALL",
            "CONFIRM UNINSTALL",
            "TS",
            "WIFI WARN ",
            "TOKEN ADD ",
            "TOKEN RESET",
            "TOKEN COPY",
            "SB bat",
            "SB journal",
            "REANAME ",
            "LAN ",
            "LAN CN",
            "LAN EN",
            "LANGUAGE ",
            "LANGUAGE CHINESE",
            "LANGUAGE ENGLISH",
        },
        key=lambda s: s.upper(),
    )
)


def _lcp_length_case_insensitive(strings: List[str]) -> int:
    if not strings:
        return 0
    upper = [s.upper() for s in strings]
    limit = min(len(s) for s in upper)
    i = 0
    while i < limit and all(s[i] == upper[0][i] for s in upper):
        i += 1
    return i


CHAT_LINE_COMPLETIONS: Tuple[str, ...] = ("help", "rs", "ts")
WINDOWS_CONSOLE_LINE_HISTORY: List[str] = []


def _apply_typing_casing(user_line: str, completed_canonical: str) -> str:
    """Match completion casing to how the user typed (see print_main_help Tab note)."""
    if not user_line:
        return completed_canonical
    if user_line.islower():
        return completed_canonical.lower()
    # Sentence-style: "Startup t", "Startup ", "Startup true" — first word Title, rest lowercase.
    if " " in user_line:
        first_sp = user_line.find(" ")
        first_word = user_line[:first_sp]
        after_last_sp = user_line[user_line.rfind(" ") + 1 :]
        if (
            first_word
            and first_word[0].isupper()
            and (len(first_word) == 1 or first_word[1:].islower())
            and (after_last_sp == "" or after_last_sp.islower())
        ):
            return completed_canonical.lower()
    if (
        len(user_line) >= 2
        and user_line[0].isupper()
        and user_line[1:].islower()
    ):
        return completed_canonical.lower()
    return completed_canonical.upper()


def _readline_completion_suffix(before: str, cased_full: str, raw_m: str) -> str:
    """Return the string readline should insert; align before/cased_full by case-insensitive prefix."""
    n = min(len(before), len(cased_full))
    i = 0
    while i < n and before[i].upper() == cased_full[i].upper():
        i += 1
    if i == len(before):
        return cased_full[i:]
    if cased_full.upper().startswith(before.upper()):
        return cased_full[len(before) :]
    return raw_m[len(before) :]


def _line_tab_extend(line: str, completions: Tuple[str, ...]) -> Tuple[str, bool]:
    """Return (new_line, extended) after one Tab press for a fixed completion list."""
    matches = [c for c in completions if c.upper().startswith(line.upper())]
    if not matches:
        return line, False
    if len(matches) == 1:
        m = matches[0]
        if line.upper() == m.upper():
            return line, False
        cased = _apply_typing_casing(line, m)
        return cased, True
    k = _lcp_length_case_insensitive(matches)
    unified_canon = matches[0][:k]
    cased = _apply_typing_casing(line, unified_canon)
    if cased != line:
        return cased, True
    return line, False


def _build_readline_line_completer(
    completions: Tuple[str, ...],
    on_empty_tab: Optional[Callable[[], None]] = None,
):
    def completer(text: str, state: int) -> Optional[str]:
        if _readline is None:
            return None
        if state == 0:
            line0 = _readline.get_line_buffer()
            if on_empty_tab and not line0.strip():
                on_empty_tab()
                completer._matches = []  # type: ignore[attr-defined]
                completer._empty_tab_only = True  # type: ignore[attr-defined]
            else:
                completer._empty_tab_only = False  # type: ignore[attr-defined]
                beg = _readline.get_begidx()
                before = line0[:beg]
                stem_u = (before + text).upper()
                completer._matches = sorted(  # type: ignore[attr-defined]
                    [m for m in completions if m.upper().startswith(stem_u)],
                    key=lambda s: (len(s), s.upper()),
                )
                completer._line0 = line0  # type: ignore[attr-defined]
                completer._beg0 = beg  # type: ignore[attr-defined]
        if getattr(completer, "_empty_tab_only", False):
            return None
        matches: List[str] = getattr(completer, "_matches", [])
        line0 = getattr(completer, "_line0", _readline.get_line_buffer())
        beg0 = getattr(completer, "_beg0", _readline.get_begidx())
        try:
            m = matches[state]
            before = line0[:beg0]
            cased_full = _apply_typing_casing(line0, m)
            suffix = _readline_completion_suffix(before, cased_full, m)
            return suffix
        except (IndexError, AttributeError):
            return None

    return completer


def input_line_with_tab_completions(
    prompt: str,
    completions: Tuple[str, ...],
    on_empty_tab: Optional[Callable[[], None]] = None,
) -> str:
    """Read one line with Tab completing against a fixed list (readline, Windows msvcrt, or plain input)."""
    if _readline is not None:
        old_completer = _readline.get_completer()
        old_delims = _readline.get_completer_delims()
        try:
            _readline.set_completer(
                _build_readline_line_completer(completions, on_empty_tab=on_empty_tab)
            )
            _readline.set_completer_delims(" \t\n`!@#$%^&*()-=+[{]}\\|;:'\",<>/?")
            _readline.parse_and_bind("tab: complete")
            return input(prompt).strip()
        finally:
            _readline.set_completer(old_completer)
            _readline.set_completer_delims(old_delims)

    if msvcrt is not None and sys.platform == "win32":
        sys.stdout.write(prompt)
        sys.stdout.flush()
        buf: List[str] = []
        cursor = 0
        history = WINDOWS_CONSOLE_LINE_HISTORY
        hist_index = len(history)

        def _move_left(count: int = 1) -> None:
            nonlocal cursor
            n = max(0, min(count, cursor))
            if n:
                sys.stdout.write("\b" * n)
                sys.stdout.flush()
                cursor -= n

        def _move_right(count: int = 1) -> None:
            nonlocal cursor
            n = max(0, min(count, len(buf) - cursor))
            if n:
                sys.stdout.write("".join(buf[cursor : cursor + n]))
                sys.stdout.flush()
                cursor += n

        def _replace_tail_after_cursor(old_tail_len: int) -> None:
            tail = "".join(buf[cursor:])
            sys.stdout.write(tail)
            if old_tail_len > len(tail):
                sys.stdout.write(" " * (old_tail_len - len(tail)))
            back = max(len(tail), old_tail_len)
            if back:
                sys.stdout.write("\b" * back)
            sys.stdout.flush()

        def _insert_text(text: str) -> None:
            nonlocal cursor
            if not text:
                return
            old_tail_len = len(buf) - cursor
            buf[cursor:cursor] = list(text)
            cursor += len(text)
            sys.stdout.write(text)
            _replace_tail_after_cursor(old_tail_len)

        def _delete_left(count: int = 1) -> None:
            nonlocal cursor
            n = max(0, min(count, cursor))
            if n == 0:
                return
            _move_left(n)
            del buf[cursor : cursor + n]
            _replace_tail_after_cursor((len(buf) - cursor) + n)

        def _delete_right(count: int = 1) -> None:
            n = max(0, min(count, len(buf) - cursor))
            if n == 0:
                return
            del buf[cursor : cursor + n]
            _replace_tail_after_cursor((len(buf) - cursor) + n)

        def _erase_previous_word() -> None:
            # Match common text-box behavior: delete spaces first, then one word.
            n = 0
            i = cursor - 1
            while i >= 0 and buf[i].isspace():
                n += 1
                i -= 1
            while i >= 0 and not buf[i].isspace():
                n += 1
                i -= 1
            _delete_left(n)

        def _replace_line(new_line: str) -> None:
            nonlocal cursor
            old_len = len(buf)
            buf.clear()
            buf.extend(list(new_line))
            cursor = len(buf)
            # Carriage-return redraw is more stable in PowerShell/Windows Terminal
            # than backspace-based in-place erasing for tab-completion replacement.
            sys.stdout.write("\r" + prompt + new_line)
            if old_len > len(new_line):
                sys.stdout.write(" " * (old_len - len(new_line)))
            # Put visual cursor at logical cursor location (always end after replace).
            sys.stdout.write("\r" + prompt + "".join(buf))
            sys.stdout.flush()

        while True:
            ch = msvcrt.getwch()
            if ch in ("\x00", "\xe0"):
                code = msvcrt.getwch()
                if code == "H":  # Up
                    if history and hist_index > 0:
                        hist_index -= 1
                        _replace_line(history[hist_index])
                    else:
                        sys.stdout.write("\a")
                        sys.stdout.flush()
                elif code == "P":  # Down
                    if hist_index < len(history):
                        hist_index += 1
                        line = history[hist_index] if hist_index < len(history) else ""
                        _replace_line(line)
                    else:
                        sys.stdout.write("\a")
                        sys.stdout.flush()
                elif code == "K":  # Left
                    _move_left(1)
                elif code == "M":  # Right
                    _move_right(1)
                elif code == "G":  # Home
                    _move_left(cursor)
                elif code == "O":  # End
                    _move_right(len(buf) - cursor)
                elif code == "S":  # Delete
                    _delete_right(1)
                # Ignore remaining arrows/function keys.
                continue
            if ch in "\r\n":
                line = "".join(buf).strip()
                if line:
                    if not history or history[-1] != line:
                        history.append(line)
                hist_index = len(history)
                sys.stdout.write("\n")
                sys.stdout.flush()
                return line
            if ch == "\x03":
                sys.stdout.write("\n")
                raise KeyboardInterrupt
            if ch == "\x08":
                _delete_left(1)
                continue
            if ch in ("\x7f", "\x17"):
                # Ctrl+Backspace often arrives as DEL (\x7f); Ctrl+W as ETB (\x17).
                _erase_previous_word()
                continue
            if ch == "\t":
                line = "".join(buf)
                if not line.strip() and on_empty_tab:
                    on_empty_tab()
                    sys.stdout.write("\n")
                    sys.stdout.write(prompt)
                    sys.stdout.flush()
                    continue
                new_line, extended = _line_tab_extend(line, completions)
                if extended:
                    _replace_line(new_line)
                else:
                    matches = [
                        c for c in completions if c.upper().startswith(line.upper())
                    ]
                    if len(matches) > 1:
                        sys.stdout.write("\n  " + "\n  ".join(matches) + "\n")
                        sys.stdout.write(prompt + "".join(buf))
                        _move_left(len(buf) - cursor)
                        sys.stdout.flush()
                    else:
                        sys.stdout.write("\a")
                        sys.stdout.flush()
                continue
            if ord(ch) >= 32:
                _insert_text(ch)

    return input(prompt).strip()


def input_menu_choice(prompt: str) -> str:
    """Read main menu input with Tab completing known commands."""
    return input_line_with_tab_completions(
        prompt, MAIN_MENU_COMPLETIONS, on_empty_tab=print_main_help
    )


def run() -> None:
    _deps_ok = ensure_runtime_dependencies()
    if not _deps_ok:
        return
    migrate_legacy_storage_if_needed()
    setup_first_time_preferences()
    ensure_backup_folder()
    maybe_run_daily_auto_backup()
    if sys.platform == "win32":
        try:
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE
        except Exception:
            pass
    # Do NOT auto-restore on app launch; only restore when user types `restore`.
    # This prevents the journal editor from popping up with an old unsaved draft.
    open_journal_window_editor(None)


if __name__ == "__main__":
    run()
