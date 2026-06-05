from __future__ import annotations

import json
import shutil
from pathlib import Path
import zipfile

import PyInstaller.__main__


ADDON_NAME = "DailyLoggerLocalTranscriptionAddon.zip"
HELPER_NAME = "DailyLoggerLocalTranscriber"
HELPER_SCRIPT = Path("local_transcriber_helper.py")
BUILD_ROOT = Path("build") / "local_transcriber_helper"
HELPER_DIST_ROOT = Path("dist") / "DailyLoggerLocalTranscriberBuild"
HELPER_DIST = HELPER_DIST_ROOT / HELPER_NAME
STAGING_ROOT = Path("dist") / "DailyLoggerLocalTranscriptionAddon"
ADDON_ROOT = STAGING_ROOT / "local_transcription_runtime"
ADDON_HELPER_ROOT = ADDON_ROOT / HELPER_NAME


def remove_tree(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path)


def build_helper() -> Path:
    if not HELPER_SCRIPT.is_file():
        raise SystemExit(f"Missing helper script: {HELPER_SCRIPT}")
    remove_tree(BUILD_ROOT)
    remove_tree(HELPER_DIST_ROOT)
    PyInstaller.__main__.run([
        str(HELPER_SCRIPT),
        "--name",
        HELPER_NAME,
        "--onedir",
        "--noconfirm",
        "--clean",
        "--workpath",
        str(BUILD_ROOT),
        "--distpath",
        str(HELPER_DIST_ROOT),
        "--hidden-import",
        "faster_whisper",
        "--hidden-import",
        "faster_whisper.transcribe",
        "--hidden-import",
        "faster_whisper.utils",
        "--hidden-import",
        "ctranslate2",
        "--hidden-import",
        "tokenizers",
        "--hidden-import",
        "huggingface_hub",
        "--collect-all",
        "ctranslate2",
        "--collect-all",
        "tokenizers",
        "--collect-all",
        "huggingface_hub",
        "--collect-all",
        "faster_whisper",
    ])
    exe = HELPER_DIST / f"{HELPER_NAME}.exe"
    if not exe.is_file():
        raise SystemExit(f"Helper build did not produce {exe}")
    return HELPER_DIST


def build_addon() -> Path:
    helper_dir = build_helper()
    remove_tree(STAGING_ROOT)
    ADDON_ROOT.mkdir(parents=True, exist_ok=True)
    shutil.copytree(helper_dir, ADDON_HELPER_ROOT)
    (ADDON_ROOT / "addon.json").write_text(
        json.dumps(
            {
                "name": "Daily Logger Local Transcription Add-on",
                "version": "helper-v3",
                "runtime": "DailyLoggerLocalTranscriber",
            },
            indent=2,
        ),
        encoding="utf-8",
    )

    zip_path = Path("dist") / ADDON_NAME
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in ADDON_ROOT.rglob("*"):
            if path.is_file():
                archive.write(path, path.relative_to(STAGING_ROOT).as_posix())
    return zip_path


if __name__ == "__main__":
    built = build_addon()
    print(built)
