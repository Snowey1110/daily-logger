from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import base64
import ctypes
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
    from tkcalendar import DateEntry
except Exception:
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
DATA_DIR = BASE_DIR / "daily_logs"
RECORDING_DIR = DATA_DIR / "Recording"
BACKUP_DIR = DATA_DIR / "backup"
SETTINGS_DIR = BASE_DIR / "settings"
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
WHISPER_TRANSCRIBE_PROMPT_CHAR_LIMIT = 600
# Hover tooltips: narrow wrap → shorter line length, more lines (taller block).
TOOLTIP_WRAP_PX = 220
TOOLTIP_WRAP_PX_MAX = 280
JOURNAL_PREF_THEME_KEY = "journal_window_theme"


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
        entries = [
            item
            for item in entries
            if start_date.date() <= item[0].date() <= end_date.date()
        ]
        if not entries:
            return "No journal entries available in the selected date range."
    lines = []
    for _, when_value, text in entries:
        lines.append(f"- [{when_value}] {text}")
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
    matched_dates = {
        entry_date.strftime("%m/%d/%Y")
        for entry_date, _, _ in load_all_journal_entries()
        if start_date.date() <= entry_date.date() <= end_date.date()
    }
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
    return open_path_with_default_app(BASE_DIR)


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
) -> str:
    """Single-request Whisper upload. Caller handles retries/fallback strategy."""
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

    add_field("model", "whisper-1")
    if language:
        add_field("language", language)
    if prompt and prompt.strip():
        add_field("prompt", prompt.strip()[:WHISPER_TRANSCRIBE_PROMPT_CHAR_LIMIT])
    add_field("temperature", str(temperature))

    filename = upload_path.name
    try:
        audio_bytes = upload_path.read_bytes()
    except OSError as exc:
        if temp_upload is not None:
            try:
                temp_upload.unlink(missing_ok=True)
            except OSError:
                pass
        return f"Could not read audio file: {exc}"

    body_chunks.append(b"--" + boundary + crlf)
    body_chunks.append(
        f'Content-Disposition: form-data; name="file"; filename="{filename}"'.encode("utf-8")
        + crlf
    )
    body_chunks.append(b"Content-Type: audio/wav" + crlf + crlf)
    body_chunks.append(audio_bytes + crlf)
    body_chunks.append(b"--" + boundary + b"--" + crlf)
    body = b"".join(body_chunks)

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
            raw = response.read().decode("utf-8")
    except error.HTTPError as exc:
        try:
            details = exc.read().decode("utf-8")
        except Exception:
            details = str(exc)
        result = f"Whisper API error ({exc.code}): {details}"
    except Exception as exc:
        result = f"Whisper request failed: {exc}"
    else:
        try:
            parsed = json.loads(raw)
            text = parsed.get("text")
            if isinstance(text, str):
                result = text.strip()
            else:
                result = "Whisper returned an unexpected response format."
        except json.JSONDecodeError:
            result = "Whisper returned invalid JSON."

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
    )
    return any(m in needle for m in markers)


def _transcribe_audio_openai_chunked(
    file_path: Path,
    language: Optional[str],
    *,
    prompt: Optional[str] = None,
    temperature: float = 0.0,
) -> str:
    """Fallback for oversized uploads/context: split WAV and merge partial transcripts."""
    mono, rate, read_err = _read_wav_mono_int16(file_path)
    if read_err is not None or mono is None:
        return _transcribe_audio_openai_single(
            file_path, language, prompt=prompt, temperature=temperature
        )
    chunk_samples = max(int(rate * WHISPER_TRANSCRIBE_CHUNK_SEC), 1)
    if int(mono.shape[0]) <= chunk_samples:
        return _transcribe_audio_openai_single(
            file_path, language, prompt=prompt, temperature=temperature
        )

    transcripts: List[str] = []
    sample_count = int(mono.shape[0])
    for start in range(0, sample_count, chunk_samples):
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
                tmp, language, prompt=prompt, temperature=temperature
            ).strip()
        finally:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
        if _is_likely_api_error_message(chunk_result):
            return chunk_result
        if chunk_result:
            transcripts.append(chunk_result)
    merged = " ".join(t for t in transcripts if t.strip()).strip()
    if merged:
        return merged
    return "Whisper returned empty text."


def transcribe_audio_openai(
    file_path: Path,
    language: Optional[str],
    *,
    prompt: Optional[str] = None,
    temperature: float = 0.0,
) -> str:
    """Send local audio to Whisper with fallback for long context/uploads."""
    upload_path, prep_err, temp_upload = prepare_wav_path_for_whisper(file_path)
    if prep_err is not None:
        return prep_err
    try:
        first_try = _transcribe_audio_openai_single(
            upload_path, language, prompt=prompt, temperature=temperature
        )
        if _whisper_context_too_long_error(first_try):
            return _transcribe_audio_openai_chunked(
                upload_path, language, prompt=prompt, temperature=temperature
            )
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
    root.title("Journal Window")
    root.geometry("1360x720")
    root.minsize(1020, 620)
    theme_holder: List[JournalWindowThemeSpec] = [load_journal_window_theme_spec()]

    def th() -> JournalWindowThemeSpec:
        return theme_holder[0]

    t_init = th()
    root.configure(bg=t_init.surface)
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
    is_edit_mode = bool(edit_target_sheet and edit_target_row > 0)

    top = tk.Frame(root, bg=t_init.panel, bd=0, highlightthickness=0)
    top.pack(fill="x", padx=t_init.pad_outer, pady=t_init.pad_top_y)
    top.grid_columnconfigure(5, weight=1)
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
    _uth_bg, _uth_fg = t_init.toolbar_hover()
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
    find_row = tk.Frame(root, bg=t_init.panel, bd=0, highlightthickness=0)
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
        lambda: "Search across Journal Text, Speech to text, and AI report.",
    )
    bind_hover_tooltip(
        find_scope_one_rb,
        lambda: "Search only in the currently active text box.",
    )
    bind_hover_tooltip(
        find_case_chk,
        lambda: "Match uppercase/lowercase exactly when enabled.",
    )
    bind_hover_tooltip(
        find_word_chk,
        lambda: "Match whole words only when enabled.",
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

    center = tk.Frame(root, bg=t_init.surface)
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
        return "Opens the recording directory."

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

    record_stop = threading.Event()
    record_pause = threading.Event()
    record_thread_holder: Dict[str, object] = {"thread": None}
    record_path_holder: Dict[str, object] = {"path": None}
    recording_ui_busy = {"v": False}
    last_journal_wav: Dict[str, Optional[Path]] = {"path": None}
    transcribing_busy = {"v": False}
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
        p = last_journal_wav.get("path")
        if p is not None and isinstance(p, Path) and p.exists():
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
            return "Transcribing…"
        p = last_journal_wav.get("path")
        if p is None or not isinstance(p, Path) or not p.exists():
            return "Record an audio first"
        return (
            "Transcribe previous recording to text.\n"
            "Uses a small amount of API cost."
        )

    def run_transcribe() -> None:
        p = last_journal_wav.get("path")
        if p is None or not isinstance(p, Path) or not p.exists():
            return
        if transcribing_busy["v"]:
            return
        if not get_openai_api_key():
            messagebox.showerror(
                "Speech to text",
                "No OpenAI API key. Use TOKEN ADD in the main menu or set OPENAI_API_KEY.",
            )
            return
        transcribing_busy["v"] = True
        update_transcribe_ui()
        stt_status.config(text="Transcribing…")
        lang_snap = _language_code_for_whisper()

        def work() -> None:
            result = transcribe_audio_openai(p, lang_snap, temperature=0.0)

            def done() -> None:
                transcribing_busy["v"] = False
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
        p = last_journal_wav.get("path")
        if p is not None and isinstance(p, Path) and p.exists():
            return t.side_action_bind_rest()
        b0, b1, b2, b3, b4 = t.transcribe_idle_disabled_config()
        return ("disabled", b0, b1, b2, b3)

    bind_button_hover_if_enabled(
        transcribe_btn,
        transcribe_rest_style,
        lambda: th().hover_primary,
        lambda: "white",
    )
    wave_canvas.bind("<Configure>", lambda _e: redraw_waveform_canvas())
    wave_canvas.after(80, redraw_waveform_canvas)

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
            _set_stt_saved_path_display(f"Saved: {str(dest)}")
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
        return (
            "Uses the AI report feature (ChatGPT) to build a summary from your journal text "
            "and speech transcript. Requires an API key and uses paid API usage."
        )

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
        if not text_value and is_edit_mode:
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

    button_row = tk.Frame(root, bg=t_init.surface)
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

    def apply_journal_window_colors() -> None:
        t = th()
        root.configure(bg=t.surface)
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
        for _b in (find_prev_btn, find_next_btn, find_close_btn):
            _b.config(bg=tbg, fg=tfg, activebackground=tabg, activeforeground=tafg)
        theme_toggle_btn.config(
            text=t.toggle_label,
            bg=tbg,
            fg=tfg,
            activebackground=tabg,
            activeforeground=tafg,
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
    theme_toggle_btn.grid(row=0, column=5, sticky="e", padx=(0, 12), pady=12)
    bind_button_hover_if_enabled(
        theme_toggle_btn,
        lambda: th().toolbar_bind_rest(),
        lambda: th().toolbar_hover()[0],
        lambda: th().toolbar_hover()[1],
    )

    def _on_escape(event=None) -> None:
        if str(find_row.winfo_manager()) == "pack":
            _find_close()
            return
        on_close(event)

    root.bind("<Escape>", _on_escape)
    root.protocol("WM_DELETE_WINDOW", on_close)
    autosave()
    refresh_save_entry_state()
    root.mainloop()
    return saved["value"]


def maybe_prompt_startup_on_first_run() -> None:
    prefs = load_preferences()
    if prefs.get("startup_prompt_done", "").lower() == "true":
        return
    print("Open logger automatically when computer starts? (y/N): ", end="")
    answer = input().strip().lower()
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
    answer = input().strip().lower()
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
        startup_answer = input().strip().lower()
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
    print("  OPEN DIRECTORY   - open current app folder")
    print("  OPEN JOURNAL     - open Journal.xlsx")
    print("  OPEN SCREENSHOTS - open chat_screenshots folder")
    print("  DIRECTOR OPEN - open current app folder in File Explorer")
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
            print(f"Opened folder: {BASE_DIR}")
        else:
            print("Could not open current folder in File Explorer.")
        return True, app_name
    if key.startswith("OPEN "):
        open_target = raw[5:].strip().upper()
        if open_target == "DIRECTORY":
            if open_current_directory_in_explorer():
                print(f"Opened folder: {BASE_DIR}")
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
            "Unknown choice. Please enter J, J SETTINGS, J SETTING, JOURNAL SETTINGS, JS, R, RT, C, CT, H, HELP, RENAME, STARTUP TRUE/FALSE, DEFAULT WINDOWS/CONSOLE, OPEN DIRECTORY/JOURNAL/SCREENSHOTS, DIRECTOR OPEN, BACKUP START/TRUE/FALSE/LIMITED, TS, UNINSTALL, CONFIRM UNINSTALL, SB bat/journal, WIFI WARN [name], RESTORE, TOKEN ADD/RESET/COPY, or press Enter to skip."
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
    if not ensure_runtime_dependencies():
        return
    app_name = setup_first_time_preferences()
    ensure_backup_folder()
    maybe_run_daily_auto_backup()
    print(f"{app_name} started.")
    while True:
        print_menu(app_name)
        choice = input_menu_choice("Your choice: ")
        keep_running, app_name = handle_choice(choice, app_name)
        if not keep_running:
            break

    print("Goodbye.")


if __name__ == "__main__":
    run()
