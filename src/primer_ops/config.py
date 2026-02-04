from __future__ import annotations

import os

OUTPUT_DIR_ENV = "OUTPUT_DIR"


def get_output_dir() -> str:
    return os.getenv(OUTPUT_DIR_ENV, "").strip()
