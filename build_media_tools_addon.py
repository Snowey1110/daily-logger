from __future__ import annotations

import json
import shutil
from pathlib import Path
import zipfile


ADDON_NAME = "DailyLoggerMediaToolsAddon.zip"
STAGING_ROOT = Path("dist") / "DailyLoggerMediaToolsAddon"
ADDON_ROOT = STAGING_ROOT / "media_tools"


def build_addon() -> Path:
    try:
        import imageio_ffmpeg
    except Exception as exc:
        raise SystemExit(f"imageio-ffmpeg is required to build the Media Tools add-on: {exc}")

    ffmpeg_path = Path(imageio_ffmpeg.get_ffmpeg_exe())
    if not ffmpeg_path.is_file():
        raise SystemExit(f"Could not locate ffmpeg executable: {ffmpeg_path}")

    if STAGING_ROOT.exists():
        shutil.rmtree(STAGING_ROOT)
    ADDON_ROOT.mkdir(parents=True, exist_ok=True)
    shutil.copy2(ffmpeg_path, ADDON_ROOT / "ffmpeg.exe")
    (ADDON_ROOT / "addon.json").write_text(
        json.dumps(
            {
                "name": "Daily Logger Media Tools Add-on",
                "version": "1",
                "tools": ["ffmpeg"],
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
