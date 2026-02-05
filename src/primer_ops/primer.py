from __future__ import annotations

import json
import os
from pathlib import Path
import re
import time
from typing import Any, Iterable, Optional

from primer_ops.progress import spinner, format_seconds
from dotenv import find_dotenv, load_dotenv
from openai import APITimeoutError, OpenAI, NotFoundError, RateLimitError
from openpyxl import load_workbook

from primer_ops.config import get_output_dir

def _normalize(text: str) -> str:
    return " ".join(text.strip().split()).lower()


def _iter_cells(ws) -> Iterable[Any]:
    for row in ws.iter_rows():
        for cell in row:
            yield cell


def _find_anchor_exact(ws, label: str, start_row: int = 1) -> Optional[Any]:
    target = _normalize(label)
    for cell in _iter_cells(ws):
        if cell.row < start_row:
            continue
        if isinstance(cell.value, str) and _normalize(cell.value) == target:
            return cell
    return None


def _find_anchor_exact_in_window(
    ws, label: str, start_row: int, end_row: int
) -> Optional[Any]:
    if end_row <= start_row:
        return None
    target = _normalize(label)
    for row in ws.iter_rows(min_row=start_row, max_row=end_row - 1):
        for cell in row:
            if isinstance(cell.value, str) and _normalize(cell.value) == target:
                return cell
    return None


def _find_anchor_contains(ws, label: str, start_row: int = 1) -> Optional[Any]:
    target = _normalize(label)
    for cell in _iter_cells(ws):
        if cell.row < start_row:
            continue
        if isinstance(cell.value, str) and target in _normalize(cell.value):
            return cell
    return None


def _require_anchor(cell: Optional[Any], label: str) -> Any:
    if cell is None:
        raise SystemExit(f"ERROR: missing anchor: {label}")
    return cell


def _first_right_value(ws, row: int, col: int, limit: int = 6) -> Optional[Any]:
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


_REQUEST_TIMEOUT_SECONDS = 30 * 60
_REMINDER_LINE_RE = re.compile(r"^\s*\(here\s+copy\s+and\s+paste.*\)\s*$", re.IGNORECASE)
_REMINDER_SENTENCE = "(here copy and paste introduction from 'company and industry intro' step 1)"
_CONTEXT_HEADER_RE = re.compile(r"^\s*###\s*context\b", re.IGNORECASE)
_SECTION_HEADER_RE = re.compile(r"^\s*###\s+", re.IGNORECASE)


def _strip_human_reminders(text: str) -> str:
    if not text:
        return text
    lines = text.splitlines()
    filtered: list[str] = []
    for line in lines:
        if _REMINDER_LINE_RE.match(line):
            continue
        if _REMINDER_SENTENCE in line.lower():
            continue
        filtered.append(line)
    return "\n".join(filtered)


def _build_prev_context_block(
    prev_sheet_name: str | None, prev_sheet_output_text: str | None
) -> tuple[str, str | None, int]:
    prev_output = (prev_sheet_output_text or "").strip()
    prev_output_chars = len(prev_output)
    if not prev_output:
        return "### CONTEXT\n\n(No previous sheet context available.)\n", None, prev_output_chars
    prev_name_used = prev_sheet_name or "N/A"
    block = "### CONTEXT\n\n" + prev_output + "\n"
    return block, prev_name_used, prev_output_chars


def _post_process_prompt(
    prompt_original: str, prev_sheet_name: str | None, prev_sheet_output_text: str
) -> tuple[str, bool, str | None, int]:
    cleaned = _strip_human_reminders(prompt_original or "")
    context_block, prev_name_used, prev_output_chars = _build_prev_context_block(
        prev_sheet_name, prev_sheet_output_text
    )
    pattern = r"^###\s*CONTEXT\s*$.*?(?=^###\s+|\Z)"
    if re.search(pattern, cleaned, flags=re.MULTILINE | re.DOTALL):
        prompt_final = re.sub(pattern, context_block, cleaned, flags=re.MULTILINE | re.DOTALL)
    else:
        prompt_final = f"{context_block}\n{cleaned}"
    return prompt_final, True, prev_name_used, prev_output_chars


def _extract_output_text_from_response(response: Any) -> str | None:
    if response is None:
        return None
    if isinstance(response, dict):
        output_text = response.get("output_text")
        if isinstance(output_text, str) and output_text.strip():
            return output_text
        for item in response.get("output", []) or []:
            if isinstance(item, dict) and item.get("type") == "output_text" and item.get("text"):
                return item["text"]
    return None


def _ensure_output_text(step_entry: dict[str, Any]) -> str | None:
    existing = step_entry.get("output_text")
    if isinstance(existing, str) and existing.strip():
        return existing
    derived = _extract_output_text_from_response(step_entry.get("response"))
    if isinstance(derived, str) and derived.strip():
        step_entry["output_text"] = derived
        return derived
    return None


def _model_supports_reasoning_effort(model: str | None) -> bool:
    if not model:
        return False
    return model.strip().lower().startswith("gpt-5")


def _error_is_empty(error: Any) -> bool:
    if error is None:
        return True
    if isinstance(error, dict):
        return len(error) == 0
    if isinstance(error, str):
        return error.strip() == ""
    return False


def _step_is_completed(step_entry: dict[str, Any]) -> bool:
    if not _error_is_empty(step_entry.get("error")):
        return False
    output_text = _ensure_output_text(step_entry) or ""
    return bool(output_text.strip())


def get_initial_context(output_dir_path: Path) -> str:
    primer_path = output_dir_path / "primer_step1_company_introduction.md"
    if primer_path.exists():
        try:
            text = primer_path.read_text(encoding="utf-8")
        except OSError:
            text = ""
        lines = text.splitlines()
        if lines and lines[0].lstrip().startswith("#"):
            lines = lines[1:]
        return "\n".join(lines).strip()
    return ""


def _response_to_dict(response: Any) -> dict[str, Any]:
    if hasattr(response, "model_dump"):
        return response.model_dump()
    if hasattr(response, "dict"):
        return response.dict()
    if isinstance(response, dict):
        return response
    return {"raw": str(response)}


def _find_step1_anchor(ws, start_row: int = 1) -> Optional[Any]:
    pattern = re.compile(r"^step\s*1(\D|$)", re.IGNORECASE)
    for cell in _iter_cells(ws):
        if cell.row < start_row:
            continue
        if isinstance(cell.value, str) and pattern.search(cell.value.strip()):
            return cell
    return None


def _find_step_anchor(ws, step_number: int, start_row: int = 1) -> Optional[Any]:
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


def _is_model_not_found_error(err: Exception) -> bool:
    msg = str(err).lower()
    if isinstance(err, NotFoundError):
        return ("model" in msg) or ("model_not_found" in msg)
    status_code = getattr(err, "status_code", None)
    if status_code == 404:
        return ("model" in msg) or ("model_not_found" in msg)
    return "model_not_found" in msg


def _format_error_reason(err: Exception) -> str:
    status_code = getattr(err, "status_code", None)
    if status_code is not None:
        return f"{status_code}: {err}"
    return str(err)


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


def _confirm_continue_after_timeout() -> bool:
    prompt = "Request timed out after 30 minutes. Continue and retry? [y/N]: "
    while True:
        try:
            answer = input(prompt).strip().lower()
        except EOFError:
            return False
        if answer in ("y", "yes"):
            return True
        if answer in ("n", "no", ""):
            return False
        print("Please answer y or n.")


def _call_openai_with_retries(
    client: OpenAI,
    request_kwargs: dict[str, Any],
    *,
    max_retries: int,
    base_sleep_seconds: float,
) -> Any:
    attempt = 0
    while True:
        try:
            return client.responses.create(**request_kwargs)
        except APITimeoutError:
            print(f"Request timed out after {int(_REQUEST_TIMEOUT_SECONDS / 60)} minutes.")
            if _confirm_continue_after_timeout():
                continue
            raise SystemExit("Aborted after timeout.")
        except RateLimitError as err:
            msg = str(err).lower()
            code = getattr(err, "code", None)
            should_retry = (code == "rate_limit_exceeded") or ("rate limit reached" in msg)
            if not should_retry:
                raise
            if attempt >= max_retries:
                raise
            match = re.search(r"try again in\s+(\d+)ms", msg)
            if match:
                sleep_seconds = (int(match.group(1)) + 50) / 1000.0
            else:
                sleep_seconds = min(base_sleep_seconds * (2**attempt), 10.0)
            print(
                f"Rate limited (attempt {attempt + 1}/{max_retries}). "
                f"Sleeping {sleep_seconds:.2f}s then retrying..."
            )
            time.sleep(sleep_seconds)
            attempt += 1


def generate_primer(
    output_dir: str | None = None,
    sheet: str | None = None,
    include: str | None = None,
    exclude: str | None = None,
    resume: bool = True,
) -> None:
    env_path = find_dotenv(usecwd=True)
    load_dotenv(env_path, override=True)
    t0 = time.perf_counter()
    step = 0
    total_steps = 5

    def log_step(msg: str) -> None:
        nonlocal step
        step += 1
        print(f"[{step}/{total_steps}] {msg}", flush=True)

    resolved_output_dir = (output_dir or get_output_dir() or "").strip()
    if not resolved_output_dir:
        raise SystemExit("ERROR: OUTPUT_DIR is not set. Please set it in the .env file.")

    prompt_library_path = os.getenv("PROMPT_LIBRARY_PATH", "").strip()
    if not prompt_library_path:
        raise SystemExit("ERROR: PROMPT_LIBRARY_PATH is not set. Please set it in the .env file.")

    base_model = os.getenv("OPENAI_MODEL", "").strip() or "gpt-5.2"
    deep_model = os.getenv("OPENAI_DEEP_RESEARCH_MODEL", "").strip() or "o4-mini-deep-research"
    max_retries = int(os.getenv("OPENAI_MAX_RETRIES", "").strip() or 6)
    base_sleep_seconds = float(os.getenv("OPENAI_RETRY_BASE_SECONDS", "").strip() or 0.5)

    output_dir_path = Path(resolved_output_dir)
    output_dir_path.mkdir(parents=True, exist_ok=True)
    dossier_path = output_dir_path / "_dossier" / "lead_input.json"
    if not dossier_path.exists():
        raise SystemExit(f"ERROR: lead_input.json not found at {dossier_path}")

    log_step("Loading lead_input.json")
    lead = json.loads(dossier_path.read_text(encoding="utf-8"))
    company_name = str(lead.get("company_name", "")).strip()
    if not company_name:
        raise SystemExit("ERROR: company_name missing from lead_input.json")
    lead_plus = dict(lead)
    lead_plus["client"] = company_name
    lead_plus["company"] = company_name
    lead_plus["company_name"] = company_name
    for key in list(lead_plus.keys()):
        lead_plus[str(key).upper()] = lead_plus[key]

    prompt_path = Path(prompt_library_path)
    if not prompt_path.is_absolute():
        repo_root = Path(__file__).resolve().parents[2]
        prompt_path = repo_root / prompt_path
    log_step("Loading prompt library (Excel)")
    workbook = load_workbook(prompt_path, data_only=True)

    def is_runnable_sheet(candidate_ws) -> bool:
        return (_find_anchor_exact(candidate_ws, "Instructions") is not None) and (
            _find_anchor_exact(candidate_ws, "Prompts") is not None
        )

    def parse_sheet_filter(filter_value: str, available: list[str]) -> set[str]:
        if "," in filter_value:
            requested = [item.strip() for item in filter_value.split(",") if item.strip()]
            available_by_lower = {name.lower(): name for name in available}
            matched = set()
            for name in requested:
                match = available_by_lower.get(name.lower())
                if match:
                    matched.add(match)
            return matched
        regex = re.compile(filter_value, re.IGNORECASE)
        return {name for name in available if regex.search(name)}

    runnable_sheets = [name for name in workbook.sheetnames if is_runnable_sheet(workbook[name])]

    if sheet:
        if sheet not in workbook.sheetnames:
            raise SystemExit(f"ERROR: sheet not found: {sheet}")
        if sheet not in runnable_sheets:
            raise SystemExit(
                f"ERROR: sheet is not runnable (missing Instructions/Prompts): {sheet}"
            )
        selected_sheets = [sheet]
    else:
        selected_sheets = list(runnable_sheets)

    if sheet is None:
        if include:
            selected_sheets = [
                name for name in selected_sheets if name in parse_sheet_filter(include, selected_sheets)
            ]
        if exclude:
            excluded = parse_sheet_filter(exclude, selected_sheets)
            selected_sheets = [name for name in selected_sheets if name not in excluded]

    if not selected_sheets:
        raise SystemExit("ERROR: no runnable sheets selected.")

    print(f"Sheets to run: {', '.join(selected_sheets)}")

    primer_sections: list[str] = ["# Commercial Primer"]
    sources_path = output_dir_path / "sources.json"
    sources_payload: dict[str, Any] = {
        "prompt_library_path": str(prompt_path),
        "sheets": [],
    }
    if resume and sources_path.exists():
        try:
            loaded_payload = json.loads(sources_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            loaded_payload = None
        if isinstance(loaded_payload, dict):
            sources_payload = loaded_payload
            sources_payload["prompt_library_path"] = str(prompt_path)
            if not isinstance(sources_payload.get("sheets"), list):
                sources_payload["sheets"] = []

    first_sheet_written = False
    prev_sheet_output_text = get_initial_context(output_dir_path)
    prev_sheet_name: str | None = "seed:intro" if prev_sheet_output_text.strip() else None
    client = OpenAI(timeout=_REQUEST_TIMEOUT_SECONDS)

    sheets_by_name: dict[str, dict[str, Any]] = {}
    for sheet_entry in sources_payload.get("sheets", []) or []:
        if isinstance(sheet_entry, dict) and isinstance(sheet_entry.get("name"), str):
            sheets_by_name[sheet_entry["name"]] = sheet_entry

    for sheet_index, sheet_name in enumerate(selected_sheets, start=1):
        ws = workbook[sheet_name]
        log_step(f"Parsing sheet anchors and building prompt ({sheet_name})")
        instructions_cell = _require_anchor(_find_anchor_exact(ws, "Instructions"), "Instructions")
        web_search_cell = _require_anchor(
            _find_anchor_exact(ws, "Web Search", start_row=instructions_cell.row),
            "Web Search",
        )
        web_search_value = _first_right_value(ws, web_search_cell.row, web_search_cell.column)
        if web_search_value is None:
            raise SystemExit("ERROR: missing anchor: Web Search")
        deep_research_cell = _require_anchor(
            _find_anchor_exact(ws, "Deep Research", start_row=instructions_cell.row),
            "Deep Research",
        )
        deep_research_value = _first_right_value(
            ws, deep_research_cell.row, deep_research_cell.column, limit=20
        )
        if deep_research_value is None:
            raise SystemExit("ERROR: missing anchor: Deep Research")

        reasoning_effort: str | None = None
        gpt_model_cell = _find_anchor_exact(ws, "GPT Model", start_row=instructions_cell.row)
        if gpt_model_cell is not None:
            gpt_model_value = _first_right_value(ws, gpt_model_cell.row, gpt_model_cell.column)
            if isinstance(gpt_model_value, str):
                normalized = gpt_model_value.strip()
                if normalized == "Thinking - Reasoning: Extended":
                    reasoning_effort = "xhigh"
                elif normalized == "Thinking":
                    reasoning_effort = "high"
                elif normalized == "Auto":
                    reasoning_effort = None

        web_search_enabled = _normalize(str(web_search_value)).startswith("enable")
        deep_research_requested = _normalize(str(deep_research_value)).startswith("enable")
        if deep_research_requested and not web_search_enabled:
            raise SystemExit(
                "ERROR: Deep Research is Enable but Web Search is not Enable. Deep Research needs a data source."
            )

        sheet_entry = sheets_by_name.get(sheet_name)
        if sheet_entry is None:
            sheet_entry = {
                "name": sheet_name,
                "web_search": web_search_enabled,
                "deep_research_requested": deep_research_requested,
                "deep_research_effective": False,
                "deep_research_error_reason": None,
                "steps": [],
            }
            sources_payload.setdefault("sheets", [])
            sources_payload["sheets"].append(sheet_entry)
            sheets_by_name[sheet_name] = sheet_entry
        else:
            sheet_entry["web_search"] = web_search_enabled
            sheet_entry["deep_research_requested"] = deep_research_requested
            sheet_entry.setdefault("steps", [])

        primer_sections.append(f"## {sheet_name}")

        prompts_cell = _require_anchor(_find_anchor_exact(ws, "Prompts"), "Prompts")
        current_sheet_output_sections: list[str] = []
        steps_by_number: dict[int, dict[str, Any]] = {}
        for existing_step in sheet_entry.get("steps", []) or []:
            if not isinstance(existing_step, dict):
                continue
            raw_number = existing_step.get("step_number")
            if isinstance(raw_number, int):
                steps_by_number[raw_number] = existing_step
            else:
                try:
                    step_num = int(raw_number)
                except (TypeError, ValueError):
                    continue
                steps_by_number[step_num] = existing_step
        step_number = 1
        while True:
            step_cell = _find_step_anchor(ws, step_number, start_row=prompts_cell.row)
            if step_cell is None:
                break
            step_label, step_title_clean = _parse_step_title(step_cell.value, step_number)
            if step_title_clean:
                heading = f"### {step_label} — {step_title_clean}"
            else:
                heading = f"### {step_label}"

            existing_step_entry = steps_by_number.get(step_number)
            if existing_step_entry is None:
                step_entry = {
                    "step_number": step_number,
                    "title": step_title_clean or step_label,
                    "prompt_original": None,
                    "prompt_final": None,
                    "injected_prev_sheet_context": None,
                    "prev_sheet_name_used": None,
                    "prev_sheet_output_chars": None,
                    "output_text": "",
                    "model": None,
                    "web_tool_type": None,
                    "request_used": None,
                    "response": None,
                    "deep_research_effective": False,
                    "deep_research_error_reason": None,
                    "error": None,
                }
                sheet_entry["steps"].append(step_entry)
                steps_by_number[step_number] = step_entry
            else:
                step_entry = existing_step_entry
                step_entry.setdefault("prompt_original", None)
                step_entry.setdefault("prompt_final", None)
                step_entry.setdefault("injected_prev_sheet_context", None)
                step_entry.setdefault("prev_sheet_name_used", None)
                step_entry.setdefault("prev_sheet_output_chars", None)
                step_entry.setdefault("output_text", "")
                step_entry.setdefault("model", None)
                step_entry.setdefault("web_tool_type", None)
                step_entry.setdefault("request_used", None)
                step_entry.setdefault("response", None)
                step_entry.setdefault("deep_research_effective", False)
                step_entry.setdefault("deep_research_error_reason", None)
                step_entry.setdefault("error", None)
            try:
                next_step_cell = _find_step_anchor(ws, step_number + 1, start_row=step_cell.row + 1)
                end_row = next_step_cell.row if next_step_cell is not None else ws.max_row + 1
                suggested_prompt_cell = _find_anchor_exact_in_window(
                    ws, "Suggested Prompt", start_row=step_cell.row, end_row=end_row
                )
                if suggested_prompt_cell is None:
                    raise ValueError(f"ERROR: missing anchor: Suggested Prompt (step {step_number})")
                suggested_prompt_value = _first_right_value(
                    ws, suggested_prompt_cell.row, suggested_prompt_cell.column, limit=20
                )
                if not isinstance(suggested_prompt_value, str) or not suggested_prompt_value.strip():
                    raise ValueError("ERROR: missing anchor: Suggested Prompt")

                prompt_original = _replace_placeholders(suggested_prompt_value, lead_plus)
                context_text = prev_sheet_output_text
                if current_sheet_output_sections:
                    sheet_context = "\n\n".join(current_sheet_output_sections).strip()
                    if sheet_context:
                        if context_text:
                            context_text = f"{context_text}\n\n{sheet_context}"
                        else:
                            context_text = sheet_context
                (
                    prompt_final,
                    injected_prev_sheet_context,
                    prev_sheet_name_used,
                    prev_sheet_output_chars,
                ) = _post_process_prompt(prompt_original, prev_sheet_name, context_text)
                step_entry["title"] = step_title_clean or step_label
                step_entry["prompt_original"] = prompt_original
                step_entry["prompt_final"] = prompt_final
                step_entry["injected_prev_sheet_context"] = injected_prev_sheet_context
                step_entry["prev_sheet_name_used"] = prev_sheet_name_used
                step_entry["prev_sheet_output_chars"] = prev_sheet_output_chars

                output_text = ""
                error_info: dict[str, str] | None = None
                call_label = f"[sheet {sheet_index}/{len(selected_sheets)}][step {step_number}]"

                if resume and existing_step_entry is not None and _step_is_completed(existing_step_entry):
                    output_text = _ensure_output_text(step_entry) or ""
                    if output_text:
                        step_entry["output_text"] = output_text.strip()
                    elif step_entry.get("output_text") is None:
                        step_entry["output_text"] = ""
                    print(f"{call_label} SKIP completed step")
                    if (not first_sheet_written) and step_number == 1 and sheet_index == 1:
                        first_sheet_written = True
                else:
                    model = base_model
                    web_tool_type: str | None = "web_search" if web_search_enabled else None
                    deep_research_effective = False
                    deep_research_error_reason: str | None = None
                    request_used: dict[str, Any] = {}
                    response: Any | None = None

                    call_start = time.perf_counter()
                    print(f"{call_label} Calling OpenAI...")
                    if deep_research_requested:
                        request_kwargs = {"model": deep_model, "input": prompt_final}
                        if web_search_enabled:
                            request_kwargs["tools"] = [{"type": "web_search_preview"}]
                        try:
                            response = _call_openai_with_retries(
                                client,
                                request_kwargs,
                                max_retries=max_retries,
                                base_sleep_seconds=base_sleep_seconds,
                            )
                            model = deep_model
                            web_tool_type = "web_search_preview" if web_search_enabled else None
                            deep_research_effective = True
                            request_used = {
                                "model": request_kwargs.get("model"),
                                "tools": request_kwargs.get("tools"),
                                "reasoning": request_kwargs.get("reasoning"),
                            }
                        except (RateLimitError, NotFoundError) as err:
                            if isinstance(err, RateLimitError):
                                deep_research_error_reason = _format_error_reason(err)
                            elif not _is_model_not_found_error(err):
                                raise
                            else:
                                deep_research_error_reason = _format_error_reason(err)
                            model = base_model
                            web_tool_type = "web_search" if web_search_enabled else None
                            request_kwargs = {"model": model, "input": prompt_final}
                            if web_search_enabled:
                                request_kwargs["tools"] = [{"type": "web_search"}]
                            if reasoning_effort is not None and _model_supports_reasoning_effort(model):
                                request_kwargs["reasoning"] = {"effort": reasoning_effort}
                            request_used = {
                                "model": request_kwargs.get("model"),
                                "tools": request_kwargs.get("tools"),
                                "reasoning": request_kwargs.get("reasoning"),
                            }
                            try:
                                response = _call_openai_with_retries(
                                    client,
                                    request_kwargs,
                                    max_retries=max_retries,
                                    base_sleep_seconds=base_sleep_seconds,
                                )
                            except RateLimitError as err_fallback:
                                error_info = {
                                    "type": type(err_fallback).__name__,
                                    "message": str(err_fallback),
                                }
                    else:
                        request_kwargs = {"model": model, "input": prompt_final}
                        if web_search_enabled:
                            request_kwargs["tools"] = [{"type": "web_search"}]
                        if reasoning_effort is not None and _model_supports_reasoning_effort(model):
                            request_kwargs["reasoning"] = {"effort": reasoning_effort}
                        request_used = {
                            "model": request_kwargs.get("model"),
                            "tools": request_kwargs.get("tools"),
                            "reasoning": request_kwargs.get("reasoning"),
                        }
                        try:
                            response = _call_openai_with_retries(
                                client,
                                request_kwargs,
                                max_retries=max_retries,
                                base_sleep_seconds=base_sleep_seconds,
                            )
                        except RateLimitError as err:
                            error_info = {"type": type(err).__name__, "message": str(err)}

                    effort_effective = None if deep_research_effective else reasoning_effort
                    call_elapsed = time.perf_counter() - call_start
                    print(f"{call_label} Done in {format_seconds(call_elapsed)}")

                    print(
                        " ".join(
                            [
                                f"model={model}",
                                f"effort={effort_effective}",
                                f"web_search={web_search_enabled}",
                                f"deep_research_requested={deep_research_requested}",
                                f"deep_research_effective={deep_research_effective}",
                                f"web_tool_type={web_tool_type}",
                            ]
                        )
                    )

                    if response is not None:
                        output_text = getattr(response, "output_text", None) or ""
                        if not output_text:
                            response_dict = _response_to_dict(response)
                            for item in response_dict.get("output", []) or []:
                                if item.get("type") == "output_text" and item.get("text"):
                                    output_text = item["text"]
                                    break

                    step_entry["model"] = model
                    step_entry["web_tool_type"] = web_tool_type
                    step_entry["request_used"] = request_used
                    step_entry["response"] = _response_to_dict(response) if response is not None else None
                    step_entry["deep_research_effective"] = deep_research_effective
                    step_entry["deep_research_error_reason"] = deep_research_error_reason
                    step_entry["error"] = error_info
                    step_entry["output_text"] = output_text.strip() if output_text else ""

                    if deep_research_effective:
                        sheet_entry["deep_research_effective"] = True
                    if deep_research_error_reason and sheet_entry["deep_research_error_reason"] is None:
                        sheet_entry["deep_research_error_reason"] = deep_research_error_reason

                    if (not first_sheet_written) and step_number == 1:
                        primer_path = output_dir_path / "primer_step1_company_introduction.md"
                        sources_path = output_dir_path / "sources_step1.json"
                        if output_text and output_text.strip():
                            primer_path.write_text(
                                "# Company Introduction\n\n" + output_text.strip() + "\n",
                                encoding="utf-8",
                            )
                        legacy_sources_payload = {
                            "prompt": prompt_final,
                            "model": model,
                            "reasoning_effort_requested": reasoning_effort,
                            "reasoning_effort_effective": effort_effective,
                            "web_search": web_search_enabled,
                            "deep_research_requested": deep_research_requested,
                            "deep_research_effective": deep_research_effective,
                            "deep_research_error_reason": deep_research_error_reason,
                            "web_tool_type": web_tool_type,
                            "request_used": request_used,
                            "response": _response_to_dict(response) if response is not None else None,
                            "error": error_info,
                        }
                        if not output_text or not output_text.strip():
                            legacy_sources_payload["error"] = legacy_sources_payload["error"] or {}
                            legacy_sources_payload["error"]["message"] = "No output returned for step 1."
                        sources_path.write_text(
                            json.dumps(legacy_sources_payload, indent=2, ensure_ascii=False),
                            encoding="utf-8",
                        )
                        first_sheet_written = True

                primer_sections.append(heading)
                if output_text:
                    primer_sections.append(output_text.strip())
                    current_sheet_output_sections.append(output_text.strip())
                elif error_info:
                    primer_sections.append(f"Error: {error_info.get('message')}")
                else:
                    primer_sections.append("Error: No output returned.")
            except Exception as err:
                step_entry["error"] = {"type": type(err).__name__, "message": str(err)}
                primer_sections.append(heading)
                primer_sections.append(f"Error: {err}")

            primer_content = "\n\n".join(primer_sections).strip() + "\n"
            _safe_write_text(output_dir_path / "primer.md", primer_content)
            _safe_write_json(output_dir_path / "sources.json", sources_payload)
            step_number += 1

        prev_sheet_output_text = "\n\n".join(current_sheet_output_sections).strip()
        prev_sheet_name = sheet_name

    if not sources_payload["sheets"]:
        raise SystemExit("ERROR: no runnable sheets executed.")

    log_step("Saving outputs")
    primer_path = output_dir_path / "primer.md"
    primer_content = "\n\n".join(primer_sections).strip() + "\n"
    _safe_write_text(primer_path, primer_content)

    sources_path = output_dir_path / "sources.json"
    _safe_write_json(sources_path, sources_payload)

    print(f"Saved: {primer_path}")
    print(f"Saved: {sources_path}")
    elapsed = time.perf_counter() - t0
    print(f"Done in {format_seconds(elapsed)}", flush=True)
