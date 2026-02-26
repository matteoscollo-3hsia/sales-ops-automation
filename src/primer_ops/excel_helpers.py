from __future__ import annotations

import re
from typing import Any, Iterable


def _normalize(text: str) -> str:
    return " ".join(text.strip().split()).lower()


def _iter_cells(ws) -> Iterable[Any]:
    for row in ws.iter_rows():
        for cell in row:
            yield cell


def _find_anchor_exact(ws, label: str, start_row: int = 1) -> Any | None:
    target = _normalize(label)
    for cell in _iter_cells(ws):
        if cell.row < start_row:
            continue
        if isinstance(cell.value, str) and _normalize(cell.value) == target:
            return cell
    return None


def _find_anchor_exact_in_window(
    ws, label: str, start_row: int, end_row: int
) -> Any | None:
    if end_row <= start_row:
        return None
    target = _normalize(label)
    for row in ws.iter_rows(min_row=start_row, max_row=end_row - 1):
        for cell in row:
            if isinstance(cell.value, str) and _normalize(cell.value) == target:
                return cell
    return None


def _require_anchor(cell: Any | None, label: str) -> Any:
    if cell is None:
        raise SystemExit(f"ERROR: missing anchor: {label}")
    return cell


def _first_right_value(ws, row: int, col: int, limit: int = 6) -> Any | None:
    for cc in range(col + 1, col + limit + 1):
        val = ws.cell(row=row, column=cc).value
        if val is not None and str(val).strip() != "":
            return val
    return None


def _replace_placeholders(text: str, lead: dict[str, Any]) -> str:
    updated = text
    for key in sorted(lead.keys(), key=lambda k: len(str(k)), reverse=True):
        raw_value = lead[key]
        value = str(raw_value)
        updated = updated.replace(f"{{{{{key}}}}}", value)
        updated = updated.replace(f"{{{key}}}", value)
        updated = updated.replace(f"#{key}#", value)
        updated = re.sub(
            rf"<<\s*{re.escape(str(key))}\s*>>",
            value,
            updated,
            flags=re.IGNORECASE,
        )
    return updated


def _find_step_anchor(ws, step_number: int, start_row: int = 1) -> Any | None:
    pattern = re.compile(rf"^step\s*{step_number}(\D|$)", re.IGNORECASE)
    for cell in _iter_cells(ws):
        if cell.row < start_row:
            continue
        if isinstance(cell.value, str) and pattern.search(cell.value.strip()):
            return cell
    return None


def _parse_step_title(step_value: Any, step_number: int) -> tuple[str, str | None]:
    step_label = f"Step {step_number}"
    if not isinstance(step_value, str):
        return step_label, None
    raw = step_value.strip()
    if not raw:
        return step_label, None
    parts = re.split(r"\s[–\-:]\s", raw, maxsplit=1)
    if len(parts) > 1 and parts[1].strip():
        return step_label, parts[1].strip()
    return step_label, None
