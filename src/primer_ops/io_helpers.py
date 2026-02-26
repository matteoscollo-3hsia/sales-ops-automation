from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Iterable


def _safe_write_text(path: Path, content: str) -> None:
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    try:
        tmp_path.write_text(content, encoding="utf-8")
        tmp_path.replace(path)
    except OSError:
        return


def _safe_write_json(path: Path, payload: dict[str, Any]) -> None:
    try:
        content = json.dumps(payload, indent=2, ensure_ascii=False)
    except (TypeError, ValueError):
        return
    _safe_write_text(path, content)


def _safe_write_text_multi(paths: Iterable[Path], content: str) -> None:
    for path in paths:
        _safe_write_text(path, content)


def _safe_write_json_multi(paths: Iterable[Path], payload: dict[str, Any]) -> None:
    for path in paths:
        _safe_write_json(path, payload)
