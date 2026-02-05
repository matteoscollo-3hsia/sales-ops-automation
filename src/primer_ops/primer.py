from __future__ import annotations

import json
import os
from pathlib import Path
import re
import time
from typing import Any, Iterable, Optional

from primer_ops.progress import spinner, format_seconds
from dotenv import find_dotenv, load_dotenv
from openai import OpenAI, NotFoundError, RateLimitError
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


def _replace_placeholders(text: str, company_name: str) -> str:
    return (
        text.replace("{{client}}", company_name)
        .replace("{{company_name}}", company_name)
        .replace("{company_name}", company_name)
        .replace("#client#", company_name)
        .replace("#company_name#", company_name)
        .replace("#company#", company_name)
        .replace("<<COMPANY_NAME>>", company_name)
    )


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


def generate_primer(output_dir: str | None = None, sheet: str = "Company and Industry Intro") -> None:
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

    prompt_path = Path(prompt_library_path)
    if not prompt_path.is_absolute():
        repo_root = Path(__file__).resolve().parents[2]
        prompt_path = repo_root / prompt_path
    log_step("Loading prompt library (Excel)")
    workbook = load_workbook(prompt_path, data_only=True)
    if sheet not in workbook.sheetnames:
        raise SystemExit(f"ERROR: sheet not found: {sheet}")
    ws = workbook[sheet]

    log_step("Parsing sheet anchors and building prompt")
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

    prompts_cell = _require_anchor(_find_anchor_exact(ws, "Prompts"), "Prompts")
    step1_cell = _require_anchor(_find_step1_anchor(ws, start_row=prompts_cell.row), "Step 1")
    suggested_prompt_cell = _require_anchor(
        _find_anchor_exact(ws, "Suggested Prompt", start_row=step1_cell.row),
        "Suggested Prompt",
    )
    suggested_prompt_value = _first_right_value(
        ws, suggested_prompt_cell.row, suggested_prompt_cell.column
    )
    if not isinstance(suggested_prompt_value, str) or not suggested_prompt_value.strip():
        raise SystemExit("ERROR: missing anchor: Suggested Prompt")

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

    prompt = _replace_placeholders(suggested_prompt_value, company_name)

    web_search_enabled = _normalize(str(web_search_value)).startswith("enable")
    deep_research_requested = _normalize(str(deep_research_value)).startswith("enable")
    if deep_research_requested and not web_search_enabled:
        raise SystemExit(
            "ERROR: Deep Research is Enable but Web Search is not Enable. Deep Research needs a data source."
        )

    client = OpenAI()
    model = base_model
    web_tool_type: str | None = "web_search" if web_search_enabled else None
    deep_research_effective = False
    deep_research_error_reason: str | None = None
    request_used: dict[str, Any] = {}
    response: Any | None = None
    error_info: dict[str, str] | None = None

    log_step("Calling OpenAI (this can take a bit)")
    with spinner("Waiting for OpenAI response"):
        if deep_research_requested:
            request_kwargs = {"model": deep_model, "input": prompt}
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
            except Exception as err:
                if isinstance(err, RateLimitError):
                    deep_research_error_reason = _format_error_reason(err)
                elif not _is_model_not_found_error(err):
                    raise
                else:
                    deep_research_error_reason = _format_error_reason(err)
                model = base_model
                web_tool_type = "web_search" if web_search_enabled else None
                request_kwargs = {"model": model, "input": prompt}
                if web_search_enabled:
                    request_kwargs["tools"] = [{"type": "web_search"}]
                if reasoning_effort is not None:
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
                    error_info = {"type": type(err_fallback).__name__, "message": str(err_fallback)}
        else:
            request_kwargs = {"model": model, "input": prompt}
            if web_search_enabled:
                request_kwargs["tools"] = [{"type": "web_search"}]
            if reasoning_effort is not None:
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


    if response is None:
        log_step("Saving outputs")
        sources_path = output_dir_path / "sources_step1.json"
        sources_payload = {
            "prompt": prompt,
            "model": model,
            "reasoning_effort_requested": reasoning_effort,
            "reasoning_effort_effective": effort_effective,
            "web_search": web_search_enabled,
            "deep_research_requested": deep_research_requested,
            "deep_research_effective": deep_research_effective,
            "deep_research_error_reason": deep_research_error_reason,
            "web_tool_type": web_tool_type,
            "request_used": request_used,
            "error": error_info,
        }
        sources_path.write_text(
            json.dumps(sources_payload, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        raise SystemExit(
            "ERROR: OpenAI rate limit exceeded after retries. "
            "Please try again later."
        )

    output_text = getattr(response, "output_text", None)
    if output_text is None:
        response_dict = _response_to_dict(response)
        output_text = ""
        for item in response_dict.get("output", []) or []:
            if item.get("type") == "output_text" and item.get("text"):
                output_text = item["text"]
                break

    log_step("Saving outputs")
    primer_path = output_dir_path / "primer_step1_company_introduction.md"
    primer_path.write_text(
        "# Company Introduction\n\n" + (output_text or "").strip() + "\n",
        encoding="utf-8",
    )

    sources_path = output_dir_path / "sources_step1.json"
    sources_payload = {
        "prompt": prompt,
        "model": model,
        "reasoning_effort_requested": reasoning_effort,
        "reasoning_effort_effective": effort_effective,
        "web_search": web_search_enabled,
        "deep_research_requested": deep_research_requested,
        "deep_research_effective": deep_research_effective,
        "deep_research_error_reason": deep_research_error_reason,
        "web_tool_type": web_tool_type,
        "request_used": request_used,
        "response": _response_to_dict(response),
        "error": error_info,
    }
    sources_path.write_text(json.dumps(sources_payload, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"Saved: {primer_path}")
    print(f"Saved: {sources_path}")
    elapsed = time.perf_counter() - t0
    print(f"Done in {format_seconds(elapsed)}", flush=True)
