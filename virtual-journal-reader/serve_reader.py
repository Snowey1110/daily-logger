#!/usr/bin/env python3
"""Serve Virtual Journal Reader (Vite dist/) and JSON API for Journal.xlsx."""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import unquote, urlparse

_REPO_ROOT = Path(__file__).resolve().parent.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

import daily_logger as dl  # noqa: E402

READER_BUILD = 13


def _dist_dir() -> Path:
    override = os.environ.get("VIRTUAL_READER_DIST", "").strip()
    if override:
        return Path(override)
    return Path(__file__).resolve().parent / "dist"


def _sketches_path() -> Path:
    return dl.SETTINGS_DIR / "journal_reader_sketches.json"


def _load_sketches_v2() -> List[Dict[str, str]]:
    """Load sketches in v2 format (list of positioned sketches).  Auto-migrates v1."""
    path = _sketches_path()
    if not path.is_file():
        return []
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    if isinstance(raw, dict) and raw.get("version") == 2:
        return list(raw.get("sketches", []))
    # v1 migration: flat { entryId: dataUrl } dict
    if isinstance(raw, dict):
        migrated: List[Dict[str, str]] = []
        for entry_id, data_url in raw.items():
            if entry_id == "version":
                continue
            if isinstance(data_url, str) and data_url.startswith("data:"):
                migrated.append({
                    "id": f"sk_{entry_id}",
                    "afterEntryId": str(entry_id),
                    "dataUrl": data_url,
                    "createdAt": "1970-01-01T00:00:00Z",
                })
        _save_sketches_v2(migrated)
        return migrated
    return []


def _save_sketches_v2(sketches: List[Dict[str, str]]) -> None:
    dl.SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
    payload = {"version": 2, "sketches": sketches}
    _sketches_path().write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _parse_id(entry_id: str) -> Optional[Tuple[str, int]]:
    if "|" not in entry_id:
        return None
    sheet, _, row_s = entry_id.rpartition("|")
    if not sheet:
        return None
    try:
        row = int(row_s)
    except ValueError:
        return None
    return sheet, row


class ReaderHandler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    dist: Path = _dist_dir()

    def log_message(self, fmt: str, *args: Any) -> None:
        return

    def _send(
        self,
        code: int,
        body: bytes,
        content_type: str,
        *,
        cache_control: Optional[str] = None,
    ) -> None:
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        if cache_control:
            self.send_header("Cache-Control", cache_control)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Connection", "close")
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, code: int, obj: object) -> None:
        data = json.dumps(obj).encode("utf-8")
        self._send(code, data, "application/json; charset=utf-8", cache_control="no-store, max-age=0")

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path or "/"

        if path == "/api/health":
            dist = _dist_dir()
            idx = dist / "index.html"
            try:
                dist_mtime = int(idx.stat().st_mtime)
            except OSError:
                dist_mtime = 0
            self._send_json(200, {"ok": True, "readerBuild": READER_BUILD, "distMtime": dist_mtime})
            return
        if path == "/api/entries":
            entries, err = dl.load_journal_reader_entries()
            sketches = _load_sketches_v2()
            payload: Dict[str, Any] = {"entries": entries, "sketches": sketches}
            if err:
                payload["error"] = err
            prefs = dl.load_preferences()
            payload["appName"] = str(prefs.get("app_name", "") or "").strip() or "Daily Logger"
            self._send_json(200, payload)
            return

        rel = unquote(path.lstrip("/"))
        if rel == "" or rel.endswith("/"):
            rel = "index.html"
        file_path = (self.dist / rel).resolve()
        try:
            file_path.relative_to(self.dist.resolve())
        except ValueError:
            self._send_json(404, {"error": "Not found"})
            return
        if not file_path.is_file():
            file_path = (self.dist / "index.html").resolve()
            if not file_path.is_file():
                self._send_json(404, {"error": "index.html missing; run npm run build"})
                return
        ext = file_path.suffix.lower()
        ctype = {
            ".html": "text/html; charset=utf-8",
            ".js": "text/javascript; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".json": "application/json; charset=utf-8",
            ".svg": "image/svg+xml",
            ".png": "image/png",
            ".ico": "image/x-icon",
            ".woff2": "font/woff2",
            ".woff": "font/woff",
        }.get(ext, "application/octet-stream")
        cc = "no-store, max-age=0" if ext in (".html", ".js", ".css", ".json", ".wasm") else None
        try:
            data = file_path.read_bytes()
        except OSError:
            self._send_json(500, {"error": "Read failed"})
            return
        self._send(200, data, ctype, cache_control=cc)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path or "/"
        length_s = self.headers.get("Content-Length", "0")
        try:
            length = int(length_s)
        except ValueError:
            length = 0
        body = self.rfile.read(length) if length > 0 else b"{}"
        try:
            payload = json.loads(body.decode("utf-8"))
        except json.JSONDecodeError:
            self._send_json(400, {"ok": False, "error": "Invalid JSON"})
            return

        if path == "/api/entry/create":
            date_val = str(payload.get("date", "")).strip()
            time_val = str(payload.get("time", "")).strip()
            if not date_val or not time_val:
                self._send_json(400, {"ok": False, "error": "Missing date or time"})
                return
            ok, msg, entry_id = dl.create_journal_reader_entry(date_val, time_val)
            if ok:
                self._send_json(200, {"ok": True, "id": entry_id})
            else:
                self._send_json(409, {"ok": False, "error": msg})
            return

        if path == "/api/entry/delete":
            entry_id = str(payload.get("id", "")).strip()
            parsed_id = _parse_id(entry_id)
            if not parsed_id:
                self._send_json(400, {"ok": False, "error": "Missing or invalid id"})
                return
            sheet_name, row_index = parsed_id
            ok, msg = dl.delete_journal_reader_entry(sheet_name, row_index)
            self._send_json(200 if ok else 409, {"ok": ok, "error": msg})
            return

        if path == "/api/entry":
            entry_id = str(payload.get("id", "")).strip()
            parsed_id = _parse_id(entry_id)
            if not parsed_id:
                self._send_json(400, {"ok": False, "error": "Missing or invalid id"})
                return
            sheet_name, row_index = parsed_id
            journal = payload.get("journal")
            speech = payload.get("speechToText")
            ai_rep = payload.get("aiReport")
            kwargs = {}
            if journal is not None:
                kwargs["journal"] = str(journal)
            if speech is not None:
                kwargs["speech_to_text"] = str(speech)
            if ai_rep is not None:
                kwargs["ai_report"] = str(ai_rep)
            if not kwargs:
                self._send_json(400, {"ok": False, "error": "No fields to update"})
                return
            ok, msg = dl.patch_journal_reader_entry(sheet_name, row_index, **kwargs)
            self._send_json(200 if ok else 409, {"ok": ok, "error": msg})
            return

        if path == "/api/sketch":
            sketches = _load_sketches_v2()

            if payload.get("delete"):
                sketch_id = str(payload.get("id", "")).strip()
                if not sketch_id:
                    self._send_json(400, {"ok": False, "error": "Missing id for delete"})
                    return
                sketches = [s for s in sketches if s.get("id") != sketch_id]
                try:
                    _save_sketches_v2(sketches)
                except OSError as exc:
                    self._send_json(500, {"ok": False, "error": str(exc)})
                    return
                self._send_json(200, {"ok": True})
                return

            sketch_id = str(payload.get("id", "")).strip()
            data_url = payload.get("dataUrl")
            if not isinstance(data_url, str) or not data_url.startswith("data:"):
                self._send_json(400, {"ok": False, "error": "Missing or invalid dataUrl"})
                return

            if sketch_id:
                found = False
                for s in sketches:
                    if s.get("id") == sketch_id:
                        s["dataUrl"] = data_url
                        found = True
                        break
                if not found:
                    self._send_json(404, {"ok": False, "error": "Sketch not found"})
                    return
            else:
                after_entry_id = str(payload.get("afterEntryId", "")).strip()
                if not after_entry_id:
                    self._send_json(400, {"ok": False, "error": "Missing afterEntryId for new sketch"})
                    return
                new_id = f"sk_{int(time.time() * 1000)}"
                sketches.append({
                    "id": new_id,
                    "afterEntryId": after_entry_id,
                    "dataUrl": data_url,
                    "createdAt": datetime.now(timezone.utc).isoformat(),
                })
                sketch_id = new_id

            try:
                _save_sketches_v2(sketches)
            except OSError as exc:
                self._send_json(500, {"ok": False, "error": str(exc)})
                return
            self._send_json(200, {"ok": True, "id": sketch_id})
            return

        self._send_json(404, {"error": "Not found"})


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()
    ReaderHandler.dist = _dist_dir()
    if not (ReaderHandler.dist / "index.html").is_file():
        print(f"Missing dist: {ReaderHandler.dist / 'index.html'} — run npm run build", file=sys.stderr)
        sys.exit(1)
    server = ThreadingHTTPServer(("127.0.0.1", args.port), ReaderHandler)
    print(f"Virtual Journal Reader at http://127.0.0.1:{args.port}/", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
