from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import base64
import ctypes
import importlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from typing import Callable, Dict, List, Optional, Tuple
from urllib import error, request
import zipfile

try:
    from openpyxl import Workbook, load_workbook
except Exception:
    Workbook = None  # type: ignore[assignment]
    load_workbook = None  # type: ignore[assignment]
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except Exception:
    tk = None
    filedialog = None
    messagebox = None
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
BACKUP_DIR = DATA_DIR / "backup"
SETTINGS_DIR = BASE_DIR / "settings"
MASTER_JOURNAL_SHEET = "Master Journal"
OPENAI_CHAT_COMPLETIONS_URL = "https://api.openai.com/v1/chat/completions"
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


def ensure_runtime_dependencies() -> bool:
    required_packages = [
        ("openpyxl", "openpyxl"),
        ("mss", "mss"),
    ]
    missing = [
        package_name
        for module_name, package_name in required_packages
        if importlib.util.find_spec(module_name) is None
    ]
    if missing:
        print("Missing required Python package(s):")
        for package_name in missing:
            print(f"  - {package_name}")
        print("Install all missing packages now? (y/N): ", end="")
        answer = input().strip().lower()
        if answer in ("y", "yes"):
            print("Installing missing packages...")
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", *missing],
                capture_output=False,
                text=True,
                check=False,
            )
            if result.returncode != 0:
                print("Package installation failed. Please install manually and try again.")
                return False
        else:
            print("Skipped package installation.")

    if not bind_openpyxl_symbols():
        print("openpyxl is required to run this app. Please install it and retry.")
        return False
    return True


def is_enter_equivalent(value: str) -> bool:
    return not value or value.upper() == "X"


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

    wb = load_workbook(workbook_path)
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
    wb = load_workbook(workbook_path)

    if module.name == "Journal":
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
    wb = load_workbook(workbook_path)

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
    wb = load_workbook(workbook_path)

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
    return {
        "sheet_name": latest_sheet.title,
        "row_index": latest_row,
        "date": date_value,
        "time": time_value,
        "text": journal_value,
        "images": [],
    }


def get_latest_journal_entry_for_delete() -> Optional[Dict[str, object]]:
    return get_latest_journal_entry_for_edit()


def update_journal_entry_at(sheet_name: str, row_index: int, row_values: List[str]) -> bool:
    module = MODULES["J"]
    workbook_path = ensure_workbook(module)
    wb = load_workbook(workbook_path)
    if sheet_name not in wb.sheetnames:
        return False
    ws = wb[sheet_name]
    if row_index < 2:
        return False

    for col_index, value in enumerate(row_values, start=1):
        ws.cell(row=row_index, column=col_index, value=value)

    rebuild_master_journal_from_daily_pages(wb, module)
    reorder_journal_sheets(wb)
    save_workbook_with_retry(wb, workbook_path)
    return True


def delete_journal_entry_at(sheet_name: str, row_index: int) -> bool:
    module = MODULES["J"]
    workbook_path = ensure_workbook(module)
    wb = load_workbook(workbook_path)
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
    wb = load_workbook(workbook_path)
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
        if len(parts) != 2:
            return None
        tokens = parts
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


def _schedule_windows_self_delete(exe_path: Path, base_dir: Path) -> None:
    script_path = Path(tempfile.gettempdir()) / f"daily_logger_uninstall_{int(time.time())}.cmd"
    script = (
        "@echo off\n"
        "timeout /t 2 /nobreak >nul\n"
        f'del /f /q "{exe_path}" >nul 2>&1\n'
        f'rd /s /q "{base_dir}" >nul 2>&1\n'
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
        _schedule_windows_self_delete(exe_path, exe_path.parent)
        print("Uninstall started. App files will be removed after this window closes.")
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


def open_journal_window_editor(draft_data: Optional[Dict[str, object]] = None) -> bool:
    if tk is None or messagebox is None:
        print("Window mode is not available on this Python setup.")
        return False

    now = datetime.now()
    default_date = now.strftime("%m/%d/%Y")
    default_time = now.strftime("%I:%M%p").lstrip("0")
    draft_text = ""
    draft_date = default_date
    draft_time = default_time
    edit_target_sheet = ""
    edit_target_row = 0
    if draft_data:
        draft_text = str(draft_data.get("text", "") or "")
        draft_date = str(draft_data.get("date", default_date) or default_date)
        draft_time = str(draft_data.get("time", default_time) or default_time)
        edit_target_sheet = str(draft_data.get("edit_target_sheet", "") or "")
        try:
            edit_target_row = int(draft_data.get("edit_target_row", 0) or 0)
        except (TypeError, ValueError):
            edit_target_row = 0

    root = tk.Tk()
    root.title("Journal Window")
    root.geometry("860x620")
    root.minsize(760, 560)
    surface_color = "#0F0F0F"
    panel_color = "#1A1A1A"
    field_color = "#141414"
    text_color = "#E6E6E6"
    muted_text_color = "#B5B5B5"
    accent_color = "#2E5A88"
    root.configure(bg=surface_color)
    # Bring the journal window to front so it does not hide behind the console.
    root.lift()
    root.attributes("-topmost", True)
    root.after(250, lambda: root.attributes("-topmost", False))
    root.focus_force()
    is_edit_mode = bool(edit_target_sheet and edit_target_row > 0)

    top = tk.Frame(root, bg=panel_color, bd=0, highlightthickness=0)
    top.pack(fill="x", padx=14, pady=(14, 10))
    top.grid_columnconfigure(5, weight=1)
    tk.Label(
        top,
        text="Date (mm/dd/yyyy):",
        bg=panel_color,
        fg=muted_text_color,
        font=("Segoe UI", 10, "bold"),
    ).grid(row=0, column=0, sticky="w", padx=(12, 0), pady=12)
    date_entry: object
    if DateEntry is not None:
        date_entry = DateEntry(
            top,
            width=14,
            date_pattern="mm/dd/yyyy",
            state="normal",  # Keep typing enabled while allowing popup calendar selection.
            background=accent_color,
            foreground="white",
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
            bg=field_color,
            fg=text_color,
            insertbackground=text_color,
            relief="flat",
            highlightthickness=1,
            highlightbackground="#2B2B2B",
            highlightcolor=accent_color,
            font=("Segoe UI", 10),
        )
        date_entry.grid(row=0, column=1, padx=(8, 20), pady=12, sticky="w")
        date_entry.insert(0, draft_date)
    tk.Label(
        top,
        text="Time (hh:mmAM/PM or rn):",
        bg=panel_color,
        fg=muted_text_color,
        font=("Segoe UI", 10, "bold"),
    ).grid(row=0, column=2, sticky="w", pady=12)
    time_entry = tk.Entry(
        top,
        width=16,
        bg=field_color,
        fg=text_color,
        insertbackground=text_color,
        relief="flat",
        highlightthickness=1,
        highlightbackground="#2B2B2B",
        highlightcolor=accent_color,
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
    tk.Button(
        top,
        text="Update Time",
        command=update_date_time_to_now,
        bg="#253F5A",
        fg=text_color,
        activebackground=accent_color,
        activeforeground="white",
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        padx=12,
        pady=6,
        cursor="hand2",
    ).grid(
        row=0, column=4, padx=(12, 12), sticky="w"
    )

    tk.Label(
        root,
        text="Journal Text",
        bg=surface_color,
        fg=muted_text_color,
        font=("Segoe UI", 10, "bold"),
    ).pack(anchor="w", padx=16, pady=(0, 6))
    editor_frame = tk.Frame(root, bg=panel_color, bd=0, highlightthickness=0)
    editor_frame.pack(fill="both", expand=True, padx=14, pady=(0, 10))
    text_box = tk.Text(
        editor_frame,
        wrap="word",
        height=18,
        bg=field_color,
        fg=text_color,
        insertbackground=text_color,
        relief="flat",
        padx=12,
        pady=12,
        font=("Consolas", 11),
        highlightthickness=1,
        highlightbackground="#2B2B2B",
        highlightcolor=accent_color,
    )
    scroll_bar = tk.Scrollbar(editor_frame, command=text_box.yview)
    text_box.configure(yscrollcommand=scroll_bar.set)
    text_box.pack(side="left", fill="both", expand=True, padx=(12, 0), pady=12)
    scroll_bar.pack(side="right", fill="y", padx=(0, 12), pady=12)
    text_box.insert("1.0", draft_text)
    text_box.focus_set()
    root.after(50, text_box.focus_set)

    saved = {"value": False}
    autosave_id = {"value": None}

    def build_draft_dict() -> Dict[str, object]:
        return {
            "text": text_box.get("1.0", "end-1c"),
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
        if edit_target_sheet and edit_target_row > 0:
            saved_ok = update_journal_entry_at(
                edit_target_sheet,
                edit_target_row,
                [date_value, normalized_time, text_value],
            )
            if not saved_ok:
                messagebox.showerror(
                    "Journal Window",
                    "Could not update previous entry. It may have changed. Try again.",
                )
                return
        else:
            append_row(MODULES["J"], [date_value, normalized_time, text_value])
        clear_journal_window_draft()
        saved["value"] = True
        if autosave_id["value"] is not None:
            root.after_cancel(autosave_id["value"])
        root.destroy()

    def on_close(event=None) -> None:
        current_text = text_box.get("1.0", "end-1c").strip()
        if not current_text:
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

    button_row = tk.Frame(root, bg=surface_color)
    button_row.pack(fill="x", padx=14, pady=(0, 14))
    tk.Button(
        button_row,
        text="Save Entry",
        command=do_save,
        bg=accent_color,
        fg="white",
        activebackground="#3D6C9F",
        activeforeground="white",
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        padx=18,
        pady=8,
        cursor="hand2",
    ).pack(side="right")

    root.bind("<Escape>", on_close)
    root.protocol("WM_DELETE_WINDOW", on_close)
    autosave()
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
        journal_context = build_journal_context_for_range(recap_date_range)
        if recap_date_range is not None:
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
    wb = load_workbook(workbook_path)
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
            return [date_value, time_value, typed_note]
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
        return [date_value, time_value, typed_note]
    return journal_settings_menu()


MODULES: Dict[str, ModuleConfig] = {
    "J": ModuleConfig(
        name="Journal",
        workbook_name="Journal.xlsx",
        sheet_name="Journal",
        headers=["Date", "Time", "Journal"],
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
        recap_range = parse_recap_date_range(raw[3:].strip(), datetime.now().year)
        if recap_range is None:
            print("Invalid date range. Examples: 4.27 4.30 | 4/27 4/30 | 4/27 - 4/30 | 4/27/2026 - 4/30/2026")
            return True, app_name
        run_chat_mode(with_journal_context=True, use_thinking_model=True, recap_date_range=recap_range)
        return True, app_name
    if key.startswith("R "):
        recap_range = parse_recap_date_range(raw[2:].strip(), datetime.now().year)
        if recap_range is None:
            print("Invalid date range. Examples: 4.27 4.30 | 4/27 4/30 | 4/27 - 4/30 | 4/27/2026 - 4/30/2026")
            return True, app_name
        run_chat_mode(with_journal_context=True, recap_date_range=recap_range)
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
            return _readline_completion_suffix(before, cased_full, m)
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
        while True:
            ch = msvcrt.getwch()
            if ch in ("\x00", "\xe0"):
                # Consume second byte for arrow/function keys and ignore them.
                msvcrt.getwch()
                continue
            if ch in "\r\n":
                sys.stdout.write("\n")
                sys.stdout.flush()
                return "".join(buf).strip()
            if ch == "\x03":
                sys.stdout.write("\n")
                raise KeyboardInterrupt
            if ch in ("\x08", "\x7f"):
                if buf:
                    buf.pop()
                    sys.stdout.write("\b \b")
                    sys.stdout.flush()
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
                    for _ in range(len(line)):
                        sys.stdout.write("\b \b")
                    sys.stdout.write(new_line)
                    buf = list(new_line)
                    sys.stdout.flush()
                else:
                    matches = [
                        c for c in completions if c.upper().startswith(line.upper())
                    ]
                    if len(matches) > 1:
                        sys.stdout.write("\n  " + "\n  ".join(matches) + "\n")
                        sys.stdout.write(prompt + "".join(buf))
                        sys.stdout.flush()
                    else:
                        sys.stdout.write("\a")
                        sys.stdout.flush()
                continue
            if ord(ch) >= 32:
                buf.append(ch)
                sys.stdout.write(ch)
                sys.stdout.flush()

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
