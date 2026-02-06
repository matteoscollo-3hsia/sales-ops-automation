from __future__ import annotations

import os
from pathlib import Path

OUTPUT_BASE_DIR_ENV = "OUTPUT_BASE_DIR"
OUTPUT_DIR_ENV = "OUTPUT_DIR"
LEAD_INPUT_PATH_ENV = "LEAD_INPUT_PATH"


def _get_env_path(name: str) -> Path | None:
    value = os.getenv(name, "").strip()
    if not value:
        return None
    return Path(value)


def get_output_base_dir() -> Path | None:
    return _get_env_path(OUTPUT_BASE_DIR_ENV)


def get_output_dir() -> Path | None:
    return _get_env_path(OUTPUT_DIR_ENV)


def get_output_root_dir() -> Path | None:
    return get_output_base_dir() or get_output_dir()


def get_lead_input_path() -> Path | None:
    return _get_env_path(LEAD_INPUT_PATH_ENV)
