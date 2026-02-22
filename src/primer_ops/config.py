from __future__ import annotations

import os
from pathlib import Path

OUTPUT_BASE_DIR_ENV = "OUTPUT_BASE_DIR"
INCLUDE_HEADINGS_ENV = "INCLUDE_HEADINGS"


def _get_env_path(name: str) -> Path | None:
    value = os.getenv(name, "").strip()
    if not value:
        return None
    return Path(value)


def get_output_base_dir() -> Path | None:
    return _get_env_path(OUTPUT_BASE_DIR_ENV)


def get_include_headings(default: bool = False) -> bool:
    value = os.getenv(INCLUDE_HEADINGS_ENV, "").strip()
    if not value:
        return default
    return value.lower() in {"1", "true", "yes", "y", "on"}
