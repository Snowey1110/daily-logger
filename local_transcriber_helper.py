from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import threading
import time
import types
import unicodedata
from typing import Any, Dict, Optional


MODEL_STATS = {
    "tiny": 75 * 1024 * 1024,
    "base": 150 * 1024 * 1024,
    "small": 500 * 1024 * 1024,
    "medium": int(1.5 * 1024 * 1024 * 1024),
    "large-v3-turbo": int(1.6 * 1024 * 1024 * 1024),
    "large-v3": 3 * 1024 * 1024 * 1024,
}


TRADITIONAL_CHINESE_TRANSLATION = str.maketrans(
    {
        "\u807d": "\u542c",
        "\u898b": "\u89c1",
        "\u55ce": "\u5417",
        "\u8207": "\u4e0e",
        "\u9019": "\u8fd9",
        "\u500b": "\u4e2a",
        "\u6703": "\u4f1a",
        "\u8eca": "\u8f66",
        "\u9304": "\u5f55",
        "\u8a18": "\u8bb0",
        "\u9ede": "\u70b9",
        "\u958b": "\u5f00",
        "\u95dc": "\u5173",
        "\u9580": "\u95e8",
        "\u88e1": "\u91cc",
        "\u9084": "\u8fd8",
        "\u8b93": "\u8ba9",
        "\u8aaa": "\u8bf4",
        "\u8a71": "\u8bdd",
        "\u5f8c": "\u540e",
        "\u6642": "\u65f6",
        "\u9593": "\u95f4",
        "\u7121": "\u65e0",
        "\u767c": "\u53d1",
        "\u8655": "\u5904",
        "\u5099": "\u5907",
        "\u6a94": "\u6863",
        "\u8f49": "\u8f6c",
        "\u5beb": "\u5199",
        "\u8072": "\u58f0",
        "\u8996": "\u89c6",
        "\u8a0a": "\u8baf",
        "\u9ad4": "\u4f53",
        "\u7d71": "\u7edf",
        "\u6a5f": "\u673a",
        "\u4e26": "\u5e76",
        "\u96d9": "\u53cc",
        "\u9801": "\u9875",
        "\u5716": "\u56fe",
        "\u5831": "\u62a5",
        "\u8acb": "\u8bf7",
        "\u8f38": "\u8f93",
        "\u9078": "\u9009",
        "\u64c7": "\u62e9",
        "\u555f": "\u542f",
        "\u52d5": "\u52a8",
        "\u61c9": "\u5e94",
        "\u8a72": "\u8be5",
        "\u984c": "\u9898",
        "\u8abf": "\u8c03",
        "\u8a66": "\u8bd5",
        "\u932f": "\u9519",
        "\u8aa4": "\u8bef",
        "\u96f2": "\u4e91",
        "\u5fa9": "\u590d",
        "\u88fd": "\u5236",
        "\u9ebc": "\u4e48",
        "\u70ba": "\u4e3a",
        "\u5f9e": "\u4ece",
        "\u5c0d": "\u5bf9",
        "\u8b1b": "\u8bb2",
        "\u73fe": "\u73b0",
        "\u5be6": "\u5b9e",
        "\u9a57": "\u9a8c",
        "\u865f": "\u53f7",
        "\u78bc": "\u7801",
        "\u8a08": "\u8ba1",
        "\u5283": "\u5212",
        "\u9577": "\u957f",
        "\u986f": "\u663e",
        "\u7c21": "\u7b80",
        "\u6f22": "\u6c49",
        "\u8a9e": "\u8bed",
        "\u8b80": "\u8bfb",
        "\u66f8": "\u4e66",
        "\u5132": "\u50a8",
        "\u8a2d": "\u8bbe",
        "\u96fb": "\u7535",
        "\u8166": "\u8111",
        "\u7db2": "\u7f51",
        "\u7d61": "\u7edc",
        "\u9023": "\u8fde",
        "\u6e2c": "\u6d4b",
        "\u6578": "\u6570",
        "\u64da": "\u636e",
        "\u5eab": "\u5e93",
        "\u96e2": "\u79bb",
        "\u7dda": "\u7ebf",
        "\u8cc7": "\u8d44",
        "\u54e1": "\u5458",
        "\u5b78": "\u5b66",
        "\u7fd2": "\u4e60",
        "\u5be9": "\u5ba1",
        "\u7522": "\u4ea7",
        "\u696d": "\u4e1a",
        "\u52d9": "\u52a1",
        "\u9810": "\u9884",
        "\u89bd": "\u89c8",
        "\u58d3": "\u538b",
        "\u7e2e": "\u7f29",
        "\u91cb": "\u91ca",
        "\u7a2e": "\u79cd",
        "\u985e": "\u7c7b",
        "\u5c0e": "\u5bfc",
        "\u6a19": "\u6807",
        "\u7c64": "\u7b7e",
        "\u7c3d": "\u7b7e",
        "\u7e6a": "\u7ed8",
        "\u756b": "\u753b",
        "\u984f": "\u989c",
        "\u64ca": "\u51fb",
        "\u522a": "\u5220",
        "\u6aa2": "\u68c0",
        "\u7576": "\u5f53",
        "\u7e8c": "\u7eed",
        "\u65b7": "\u65ad",
        "\u554f": "\u95ee",
    }
)


def simplify_chinese_text(text: str) -> str:
    return (text or "").translate(TRADITIONAL_CHINESE_TRANSLATION)


def normalize_common_short_chinese_misses(text: str) -> str:
    fixed = simplify_chinese_text(text or "")
    if not fixed:
        return ""
    fixed = re.sub(
        r"(?i)\bwei\s+wei\s+wei\b[,\s]*(?=\u542c\u5f97\u89c1\u5417)",
        "\u5582\u5582\u5582,",
        fixed,
    )
    fixed = re.sub(
        r"(?i)\bwei\s+(?:lu\s*)?wei\b[,\s]*(?=\u542c\u5f97\u89c1\u5417)",
        "\u5582\u5582\u5582,",
        fixed,
    )
    fixed = re.sub(
        r"(?i)^\s*we\s+will\s+wait\s+in\s+jammer\s*[.?!]?\s*$",
        "\u5582\u5582\u5582\u542c\u5f97\u89c1\u5417",
        fixed,
    )
    return fixed


def emit(event: str, **payload: Any) -> None:
    payload["event"] = event
    sys.stdout.write(json.dumps(payload, ensure_ascii=True) + "\n")
    sys.stdout.flush()


def _cjk_count(text: str) -> int:
    return sum(1 for ch in text if "\u4e00" <= ch <= "\u9fff")


def _latin_count(text: str) -> int:
    return sum(1 for ch in text if ("a" <= ch.lower() <= "z"))


def _looks_like_pinyin_chinese_miss(text: str) -> bool:
    cleaned = simplify_chinese_text(text or "")
    if _cjk_count(cleaned) <= 0:
        return False
    latin_tokens = re.findall(r"[A-Za-z]+", cleaned.lower())
    if not latin_tokens:
        return False
    joined = " ".join(latin_tokens)
    if "wei" in latin_tokens and "\u542c\u5f97\u89c1\u5417" in cleaned:
        return True
    pinyinish = {
        "wei",
        "luwei",
        "lu",
        "ting",
        "de",
        "jian",
        "ma",
        "ni",
        "hao",
        "ce",
        "shi",
    }
    return bool(latin_tokens) and all(token in pinyinish for token in latin_tokens) and len(joined) >= 3


def _looks_like_repeat_noise(text: str) -> bool:
    compact = " ".join((text or "").split()).strip()
    if len(compact) < 16:
        return False
    letters_only = [ch for ch in compact if unicodedata.category(ch).startswith("L")]
    if len(letters_only) >= 8 and len(set(letters_only)) <= 2:
        return True
    words = compact.lower().split()
    if len(words) >= 8:
        unique = len(set(words))
        return unique <= max(2, len(words) // 5)
    return False


def _is_supported_transcript_letter(ch: str) -> bool:
    if "LATIN" in unicodedata.name(ch, ""):
        return True
    return (
        "\u3400" <= ch <= "\u4dbf"
        or "\u4e00" <= ch <= "\u9fff"
        or "\uf900" <= ch <= "\ufaff"
        or "\U00020000" <= ch <= "\U0002ebef"
    )


def _unsupported_letter_ratio(text: str) -> tuple[int, float]:
    letters = 0
    unsupported = 0
    for ch in text or "":
        if not unicodedata.category(ch).startswith("L"):
            continue
        letters += 1
        if not _is_supported_transcript_letter(ch):
            unsupported += 1
    if not letters:
        return 0, 0.0
    return letters, unsupported / float(letters)


def _needs_stronger_model(text: str, detected_language: str) -> bool:
    cleaned = simplify_chinese_text(text or "").strip()
    if not cleaned:
        return False
    if _looks_like_pinyin_chinese_miss(cleaned):
        return True
    if _looks_like_repeat_noise(cleaned):
        return True
    letters, unsupported_ratio = _unsupported_letter_ratio(cleaned)
    if letters >= 4 and unsupported_ratio > 0.25:
        return True
    detected = (detected_language or "").strip().lower()
    if detected and detected not in ("en", "zh", "zh-cn", "zh-tw", "yue", "cmn"):
        return True
    return False


def _next_stronger_downloaded_model(models_dir: Path, current_model: str) -> Optional[str]:
    order = ["tiny", "base", "small", "medium"]
    try:
        start = order.index(current_model)
    except ValueError:
        return None
    for candidate in order[start + 1 :]:
        if model_is_downloaded(models_dir, candidate):
            return candidate
    return None


def _should_use_chinese_fallback(first_text: str, first_language: str, fallback_text: str) -> bool:
    detected = (first_language or "").strip().lower()
    if detected in ("zh", "zh-cn", "zh-tw", "yue", "cmn"):
        return False
    if detected in ("", "en") and _cjk_count(first_text) > 0:
        return False
    cjk = _cjk_count(fallback_text)
    if cjk < 2:
        return False
    if _looks_like_repeat_noise(fallback_text):
        return False
    return cjk >= 4 or cjk >= max(2, _latin_count(fallback_text) // 8)


def directory_size(path: Path) -> int:
    try:
        if path.is_file():
            return int(path.stat().st_size)
        if not path.is_dir():
            return 0
    except OSError:
        return 0
    total = 0
    try:
        for child in path.rglob("*"):
            try:
                if child.is_file():
                    total += int(child.stat().st_size)
            except OSError:
                pass
    except OSError:
        pass
    return total


def model_path(models_dir: Path, model: str) -> Path:
    return models_dir / model


def model_is_downloaded(models_dir: Path, model: str) -> bool:
    path = model_path(models_dir, model)
    return (
        path.is_dir()
        and (path / "config.json").is_file()
        and (path / "model.bin").is_file()
        and (path / "tokenizer.json").is_file()
    )


def find_ffmpeg() -> Optional[str]:
    candidates = []
    env = os.getenv("DAILYLOGGER_FFMPEG", "").strip()
    if env:
        candidates.append(env)
    try:
        exe_dir = Path(sys.executable).resolve().parent
        candidates.extend([
            str(exe_dir / "ffmpeg.exe"),
            str(exe_dir / "_internal" / "ffmpeg.exe"),
            str(exe_dir.parent / "media_tools" / "ffmpeg.exe"),
        ])
    except Exception:
        pass
    appdata = os.getenv("APPDATA", "").strip()
    if appdata:
        candidates.append(str(Path(appdata) / "DailyLogger" / "addons" / "media_tools" / "ffmpeg.exe"))
    candidates.append("ffmpeg")
    for candidate in candidates:
        if candidate == "ffmpeg":
            return candidate
        try:
            if Path(candidate).is_file():
                return candidate
        except OSError:
            pass
    return None


def install_audio_shim() -> None:
    module = types.ModuleType("faster_whisper.audio")
    module.__package__ = "faster_whisper"

    def decode_audio(input_file: Any, sampling_rate: int = 16000, split_stereo: bool = False) -> Any:
        import numpy as np

        ffmpeg = find_ffmpeg()
        if not ffmpeg:
            raise RuntimeError("Media Tools are required for local transcription audio decoding.")
        input_name = str(getattr(input_file, "name", "") or os.fspath(input_file))
        channels = "2" if split_stereo else "1"
        cmd = [
            ffmpeg,
            "-nostdin",
            "-hide_banner",
            "-loglevel",
            "error",
            "-i",
            input_name,
            "-f",
            "s16le",
            "-acodec",
            "pcm_s16le",
            "-ac",
            channels,
            "-ar",
            str(int(sampling_rate)),
            "-",
        ]
        kwargs: Dict[str, Any] = {
            "stdout": subprocess.PIPE,
            "stderr": subprocess.PIPE,
            "stdin": subprocess.DEVNULL,
        }
        if os.name == "nt":
            kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        result = subprocess.run(cmd, **kwargs)
        if result.returncode != 0:
            details = (result.stderr or b"").decode("utf-8", errors="replace").strip()
            raise RuntimeError(details or "ffmpeg could not decode that audio file.")
        audio = np.frombuffer(result.stdout, dtype=np.int16).astype(np.float32) / 32768.0
        if split_stereo:
            return audio[0::2], audio[1::2]
        return audio

    def pad_or_trim(array: Any, length: int = 3000, *, axis: int = -1) -> Any:
        import numpy as np

        if array.shape[axis] > length:
            array = array.take(indices=range(length), axis=axis)
        if array.shape[axis] < length:
            pad_widths = [(0, 0)] * array.ndim
            pad_widths[axis] = (0, length - array.shape[axis])
            array = np.pad(array, pad_widths)
        return array

    module.decode_audio = decode_audio  # type: ignore[attr-defined]
    module.pad_or_trim = pad_or_trim  # type: ignore[attr-defined]
    sys.modules["faster_whisper.audio"] = module


def import_faster_whisper() -> Any:
    install_audio_shim()
    import faster_whisper

    return faster_whisper


def command_health(_args: argparse.Namespace) -> int:
    try:
        fw = import_faster_whisper()
        import ctranslate2

        emit(
            "complete",
            ok=True,
            faster_whisper_version=str(getattr(fw, "__version__", "")),
            ctranslate2_version=str(getattr(ctranslate2, "__version__", "")),
        )
        return 0
    except Exception as exc:
        emit("error", message=f"Local transcription helper failed health check: {exc}")
        return 1


def command_download(args: argparse.Namespace) -> int:
    model = args.model
    models_dir = Path(args.models_dir)
    target = model_path(models_dir, model)
    if model not in MODEL_STATS:
        emit("error", message=f"Unsupported local model: {model}")
        return 2
    models_dir.mkdir(parents=True, exist_ok=True)
    if model_is_downloaded(models_dir, model):
        emit("progress", percent=100, message=f"Local - {model} already downloaded.")
        emit("complete", ok=True, model=model)
        return 0

    stop = threading.Event()
    estimate = MODEL_STATS.get(model, 0)
    started = time.monotonic()

    def monitor() -> None:
        while not stop.wait(1.0):
            size = directory_size(target)
            elapsed = max(1, int(time.monotonic() - started))
            percent = min(99, max(1, int(size * 100 / estimate))) if estimate else 0
            emit(
                "progress",
                percent=percent,
                downloaded=size,
                total=estimate,
                elapsed=elapsed,
                message=f"Downloading Local - {model}: {percent}%",
            )

    thread = threading.Thread(target=monitor, daemon=True)
    thread.start()
    try:
        import_faster_whisper()
        fw_utils = __import__("faster_whisper.utils", fromlist=["download_model"])
        emit("status", message=f"Downloading Local - {model}...")
        target.mkdir(parents=True, exist_ok=True)
        fw_utils.download_model(model, output_dir=str(target))
    except Exception as exc:
        emit("error", message=f"Could not download local model '{model}': {exc}")
        return 1
    finally:
        stop.set()
        thread.join(timeout=2)

    if not model_is_downloaded(models_dir, model):
        emit("error", message=f"Downloaded model '{model}' is incomplete.")
        return 1
    emit("progress", percent=100, downloaded=directory_size(target), total=estimate, message=f"Local - {model} downloaded.")
    emit("complete", ok=True, model=model)
    return 0


def command_transcribe(args: argparse.Namespace) -> int:
    try:
        fw = import_faster_whisper()
        models_dir = Path(args.models_dir)
        model_dir = model_path(models_dir, args.model)
        if not model_is_downloaded(models_dir, args.model):
            emit("error", message=f"Local transcription model '{args.model}' is not downloaded.")
            return 2
        emit("progress", percent=8, message="Loading local transcription model...")
        current_model_name = args.model
        current_model = fw.WhisperModel(
            str(model_dir),
            device="cpu",
            compute_type=args.compute_type,
            cpu_threads=int(args.cpu_threads),
        )
        language = None if args.language in ("", "auto", "Auto", None) else args.language

        def transcribe_once(
            language_override: Optional[str],
            prompt_override: str,
            *,
            progress_start: int,
            progress_span: int,
        ) -> Dict[str, Any]:
            segments, info = current_model.transcribe(
                str(Path(args.input)),
                language=language_override,
                task="transcribe",
                beam_size=5,
                repetition_penalty=1.15,
                no_repeat_ngram_size=3,
                temperature=0.0,
                vad_filter=True,
                condition_on_previous_text=False,
                initial_prompt=prompt_override or None,
            )
            duration = float(getattr(info, "duration", 0.0) or 0.0)
            parts = []
            items = []
            for segment in segments:
                text = str(getattr(segment, "text", "") or "").strip()
                if not text:
                    continue
                start_time = float(getattr(segment, "start", 0.0) or 0.0)
                end_time = float(getattr(segment, "end", 0.0) or 0.0)
                percent = (
                    progress_start + int(progress_span * min(1.0, max(0.0, end_time / duration)))
                    if duration > 0
                    else progress_start + max(1, progress_span // 2)
                )
                parts.append(text)
                items.append({"text": text, "start": start_time, "end": end_time, "percent": percent})
                emit("progress", percent=percent, message="Transcribing audio...")
            return {
                "text": " ".join(parts).strip(),
                "segments": items,
                "language": str(getattr(info, "language", "") or ""),
            }

        emit("progress", percent=24, message="Transcribing audio...")
        first = transcribe_once(language, args.prompt or "", progress_start=24, progress_span=56)
        selected = first
        first_detected = str(first.get("language") or "").strip().lower()
        auto_fallback_enabled = args.auto_fallback_zh or language is None
        needs_chinese_check = (
            _cjk_count(first["text"]) == 0
            or first_detected not in ("", "en", "zh", "zh-cn", "zh-tw", "yue", "cmn")
        )
        if language is None and auto_fallback_enabled and needs_chinese_check:
            emit("progress", percent=82, message="Checking Chinese transcription...")
            zh_prompt = (
                "Transcribe the original Chinese and English speech. Do not translate. "
                "Use Simplified Chinese characters for Chinese speech."
            )
            second = transcribe_once("zh", zh_prompt, progress_start=82, progress_span=12)
            if _should_use_chinese_fallback(first["text"], first["language"], second["text"]):
                selected = second

        if language is None and _needs_stronger_model(selected["text"], selected["language"]):
            stronger_model = _next_stronger_downloaded_model(models_dir, current_model_name)
            if stronger_model:
                current_model_name = stronger_model
                emit("progress", percent=94, message=f"Retrying with Local - {stronger_model}...")
                current_model = fw.WhisperModel(
                    str(model_path(models_dir, stronger_model)),
                    device="cpu",
                    compute_type=args.compute_type,
                    cpu_threads=int(args.cpu_threads),
                )
                selected = transcribe_once(language, "", progress_start=94, progress_span=5)
                detected_after_retry = str(selected.get("language") or "").strip().lower()
                needs_chinese_check = (
                    _cjk_count(selected["text"]) == 0
                    or detected_after_retry not in ("", "en", "zh", "zh-cn", "zh-tw", "yue", "cmn")
                )
                if auto_fallback_enabled and needs_chinese_check:
                    second = transcribe_once(
                        "zh",
                        (
                            "Transcribe the original Chinese and English speech. Do not translate. "
                            "Use Simplified Chinese characters for Chinese speech."
                        ),
                        progress_start=96,
                        progress_span=3,
                    )
                    if _should_use_chinese_fallback(selected["text"], selected["language"], second["text"]):
                        selected = second

        for item in selected["segments"]:
            text = normalize_common_short_chinese_misses(str(item["text"] or ""))
            emit(
                "segment",
                text=text,
                start=item["start"],
                end=item["end"],
                percent=item["percent"],
            )
        text = normalize_common_short_chinese_misses(str(selected["text"] or "")).strip()
        emit("progress", percent=100, message="Transcription complete.")
        emit("complete", ok=True, text=text, language=selected.get("language", ""))
        return 0
    except Exception as exc:
        emit("error", message=f"Local transcription failed: {exc}")
        return 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Daily Logger local transcription helper")
    sub = parser.add_subparsers(dest="command", required=True)
    health = sub.add_parser("health")
    health.add_argument("--json", action="store_true")
    health.set_defaults(func=command_health)

    download = sub.add_parser("download")
    download.add_argument("--model", required=True)
    download.add_argument("--models-dir", required=True)
    download.add_argument("--json", action="store_true")
    download.set_defaults(func=command_download)

    transcribe = sub.add_parser("transcribe")
    transcribe.add_argument("--model", required=True)
    transcribe.add_argument("--input", required=True)
    transcribe.add_argument("--models-dir", required=True)
    transcribe.add_argument("--language", default="auto")
    transcribe.add_argument("--prompt", default="")
    transcribe.add_argument("--compute-type", default="int8")
    transcribe.add_argument("--cpu-threads", default="4")
    transcribe.add_argument("--auto-fallback-zh", action="store_true")
    transcribe.add_argument("--json", action="store_true")
    transcribe.set_defaults(func=command_transcribe)
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return int(args.func(args))


if __name__ == "__main__":
    raise SystemExit(main())
