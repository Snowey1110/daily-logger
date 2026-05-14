#!/usr/bin/env python3
"""Serve Virtual Journal Reader (Vite dist/) and JSON API for Journal.xlsx."""

from __future__ import annotations

import argparse
import json
import os
import socket
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

_lan_access = False
_lan_ip = "127.0.0.1"


def _dist_dir() -> Path:
    override = os.environ.get("VIRTUAL_READER_DIST", "").strip()
    if override:
        return Path(override)
    return Path(__file__).resolve().parent / "dist"


def _sketches_path() -> Path:
    return dl.SETTINGS_DIR / "journal_reader_sketches.json"


def _reader_settings_path() -> Path:
    return dl.SETTINGS_DIR / "journal_reader_settings.json"


def _load_reader_settings() -> Dict[str, Any]:
    path = _reader_settings_path()
    if not path.is_file():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def _save_reader_settings(settings: Dict[str, Any]) -> None:
    dl.SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
    _reader_settings_path().write_text(
        json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _load_data() -> Tuple[List[Dict[str, str]], Dict[str, Any]]:
    """Load sketches + overlays (v3).  Auto-migrates v1/v2."""
    path = _sketches_path()
    if not path.is_file():
        return [], {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return [], {}
    if isinstance(raw, dict) and raw.get("version") == 3:
        return list(raw.get("sketches", [])), dict(raw.get("overlays", {}))
    if isinstance(raw, dict) and raw.get("version") == 2:
        sketches = list(raw.get("sketches", []))
        _save_data(sketches, {})
        return sketches, {}
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
        _save_data(migrated, {})
        return migrated, {}
    return [], {}


def _save_data(sketches: List[Dict[str, str]], overlays: Dict[str, Any]) -> None:
    dl.SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
    payload = {"version": 3, "sketches": sketches, "overlays": overlays}
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

    def _is_local(self) -> bool:
        client_ip = self.client_address[0]
        return client_ip in ("127.0.0.1", "::1", "localhost")

    def _gate_lan(self) -> bool:
        """Return True if the request should be blocked (non-local + LAN disabled)."""
        if _lan_access or self._is_local():
            return False
        self._send_json(403, {"ok": False, "error": "LAN access is disabled"})
        return True

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
        if self._gate_lan():
            return
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
        if path == "/api/lan-status":
            self._send_json(200, {"enabled": _lan_access, "ip": _lan_ip, "port": self.server.server_address[1]})
            return
        if path == "/api/reader-settings":
            self._send_json(200, _load_reader_settings())
            return
        if path == "/api/entries":
            entries, err = dl.load_journal_reader_entries()
            sketches, overlays = _load_data()
            payload: Dict[str, Any] = {"entries": entries, "sketches": sketches, "overlays": overlays}
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
        if self._gate_lan():
            return
        parsed = urlparse(self.path)
        path = parsed.path or "/"

        if path == "/api/lan-toggle":
            if not self._is_local():
                self._send_json(403, {"ok": False, "error": "Only the host machine can toggle LAN access"})
                return
            global _lan_access
            _lan_access = not _lan_access
            state = "enabled" if _lan_access else "disabled"
            print(f"  LAN access {state}", flush=True)
            self._send_json(200, {"ok": True, "enabled": _lan_access, "ip": _lan_ip, "port": self.server.server_address[1]})
            return

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

        if path == "/api/reader-settings":
            allowed_keys = {"coverTheme", "bgTheme", "sortOrder", "singlePageMode"}
            current = _load_reader_settings()
            for k in allowed_keys:
                if k in payload:
                    current[k] = payload[k]
            _save_reader_settings(current)
            self._send_json(200, {"ok": True, **current})
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
            if ok:
                sketches, overlays = _load_data()
                if entry_id in overlays:
                    del overlays[entry_id]
                    try:
                        _save_data(sketches, overlays)
                    except OSError:
                        pass
            self._send_json(200 if ok else 409, {"ok": ok, "error": msg})
            return

        if path == "/api/entry":
            entry_id = str(payload.get("id", "")).strip()
            parsed_id = _parse_id(entry_id)
            if not parsed_id:
                self._send_json(400, {"ok": False, "error": "Missing or invalid id"})
                return
            sheet_name, row_index = parsed_id
            date_val = payload.get("date")
            time_val = payload.get("time")
            journal = payload.get("journal")
            speech = payload.get("speechToText")
            ai_rep = payload.get("aiReport")
            kwargs = {}
            if date_val is not None:
                kwargs["date"] = str(date_val)
            if time_val is not None:
                kwargs["time"] = str(time_val)
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

        if path == "/api/page-overlay":
            entry_id = str(payload.get("entryId", "")).strip()
            if not entry_id:
                self._send_json(400, {"ok": False, "error": "Missing entryId"})
                return
            sketch_data = payload.get("sketchDataUrl") or ""
            images = payload.get("images") or []
            layer_order = payload.get("layerOrder") or ["text", "sketch", "images"]
            sketches, overlays = _load_data()
            has_content = bool(sketch_data) or bool(images)
            if has_content:
                overlays[entry_id] = {
                    "sketchDataUrl": sketch_data,
                    "images": images,
                    "layerOrder": layer_order,
                }
            else:
                overlays.pop(entry_id, None)
            try:
                _save_data(sketches, overlays)
            except OSError as exc:
                self._send_json(500, {"ok": False, "error": str(exc)})
                return
            self._send_json(200, {"ok": True})
            return

        if path == "/api/sketch":
            sketches, overlays = _load_data()

            if payload.get("delete"):
                sketch_id = str(payload.get("id", "")).strip()
                if not sketch_id:
                    self._send_json(400, {"ok": False, "error": "Missing id for delete"})
                    return
                sketches = [s for s in sketches if s.get("id") != sketch_id]
                try:
                    _save_data(sketches, overlays)
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
                _save_data(sketches, overlays)
            except OSError as exc:
                self._send_json(500, {"ok": False, "error": str(exc)})
                return
            self._send_json(200, {"ok": True, "id": sketch_id})
            return

        self._send_json(404, {"error": "Not found"})


def _get_lan_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--host", type=str, default="0.0.0.0",
                        help="Bind address (default 0.0.0.0 for LAN access, use 127.0.0.1 for localhost only)")
    args = parser.parse_args()
    ReaderHandler.dist = _dist_dir()
    if not (ReaderHandler.dist / "index.html").is_file():
        print(f"Missing dist: {ReaderHandler.dist / 'index.html'} — run npm run build", file=sys.stderr)
        sys.exit(1)
    global _lan_ip
    server = ThreadingHTTPServer((args.host, args.port), ReaderHandler)
    _lan_ip = _get_lan_ip()
    print(f"Virtual Journal Reader:", flush=True)
    print(f"  Local:   http://127.0.0.1:{args.port}/", flush=True)
    if args.host == "0.0.0.0":
        print(f"  Network: http://{_lan_ip}:{args.port}/  (enable LAN access in settings)", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
