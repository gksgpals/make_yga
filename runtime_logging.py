from __future__ import annotations

import json
import logging
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

LOG_DIR_NAME = "logs"
APP_LOG_FILE_NAME = "yga_app.log"
AI_RAW_LOG_FILE_NAME = "ai_parser_raw.jsonl"
AI_CALLS_DIR_NAME = "ai_calls"
ROOT_LOGGER_NAME = "yga"
LOG_TO_FILES_ENV = "YGA_LOG_TO_FILES"
FALSEY_VALUES = {"0", "false", "no", "off"}


def env_flag_enabled(name: str, *, default: bool = True) -> bool:
    raw_value = os.getenv(name)
    if raw_value is None:
        return default
    return raw_value.strip().lower() not in FALSEY_VALUES


def project_root(base_dir: Path | None = None) -> Path:
    return (base_dir or Path(__file__).resolve().parent).resolve()


def ensure_log_dir(base_dir: Path | None = None) -> Path:
    log_dir = project_root(base_dir) / LOG_DIR_NAME
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir


def app_log_path(base_dir: Path | None = None) -> Path:
    return ensure_log_dir(base_dir) / APP_LOG_FILE_NAME


def ai_raw_log_path(base_dir: Path | None = None) -> Path:
    return ensure_log_dir(base_dir) / AI_RAW_LOG_FILE_NAME


def ai_call_log_dir(base_dir: Path | None = None) -> Path:
    call_dir = ensure_log_dir(base_dir) / AI_CALLS_DIR_NAME
    call_dir.mkdir(parents=True, exist_ok=True)
    return call_dir


def ai_call_capture_path(request_id: str, base_dir: Path | None = None) -> Path:
    return ai_call_log_dir(base_dir) / f"{request_id}.txt"


def get_logger(name: str) -> logging.Logger:
    root_logger = logging.getLogger(ROOT_LOGGER_NAME)
    if not getattr(root_logger, "_yga_configured", False):
        root_logger.setLevel(logging.INFO)
        formatter = logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s")

        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(formatter)
        root_logger.addHandler(stream_handler)

        if env_flag_enabled(LOG_TO_FILES_ENV):
            file_handler = logging.FileHandler(app_log_path(), encoding="utf-8")
            file_handler.setFormatter(formatter)
            root_logger.addHandler(file_handler)

        root_logger.propagate = False
        setattr(root_logger, "_yga_configured", True)
    return logging.getLogger(f"{ROOT_LOGGER_NAME}.{name}")


def append_json_log(path: Path, payload: dict[str, Any]) -> None:
    if not env_flag_enabled(LOG_TO_FILES_ENV):
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    record = {
        "timestamp": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        **payload,
    }
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(record, ensure_ascii=False))
        handle.write("\n")


def write_text_log(path: Path, content: str) -> None:
    if not env_flag_enabled(LOG_TO_FILES_ENV):
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
