from __future__ import annotations

import json
import os
from pathlib import Path
import re
import time
from typing import Any, Iterable, Optional

from primer_ops.progress import spinner, format_seconds
from dotenv import find_dotenv, load_dotenv
from openai import OpenAI
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


def generate_primer(output_dir: str | None = None, sheet: str = "Company and Industry Intro") -> None:
    env_path = find_dotenv(usecwd=True)
    load_dotenv(env_path, override=True)
    # TEMP DEBUG
    print(f"ENV_FILE={env_path}")
    print(f"PROMPT_LIBRARY_PATH={os.getenv('PROMPT_LIBRARY_PATH','').strip()}")
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
    log_step("Loading prompt library (Excel)")
    if not prompt_path.is_absolute():
        repo_root = Path(__file__).resolve().parents[2]
        prompt_path = repo_root / prompt_path
    # TEMP DEBUG
    print(f"PROMPT_LIBRARY_ABS={prompt_path}")
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
    deep_research_enabled = _normalize(str(deep_research_value)).startswith("enable")
    if deep_research_enabled and not web_search_enabled:
        raise SystemExit(
            "ERROR: Deep Research is Enable but Web Search is not Enable. Deep Research needs a data source."
        )

    model = deep_model if deep_research_enabled else base_model

    client = OpenAI()
    request_kwargs = {"model": model, "input": prompt}
    if web_search_enabled:
        request_kwargs["tools"] = [
            {"type": "web_search_preview" if deep_research_enabled else "web_search"}
        ]
    if (not deep_research_enabled) and reasoning_effort is not None:
        request_kwargs["reasoning"] = {"effort": reasoning_effort}

    print(
        " ".join(
            [
                f"model={model}",
                f"effort={reasoning_effort}",
                f"web_search={web_search_enabled}",
                f"deep_research={deep_research_enabled}",
            ]
        )
    )
    log_step("Calling OpenAI (this can take a bit)")
    with spinner("Waiting for OpenAI response"):
        response = client.responses.create(**request_kwargs)


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
        "reasoning_effort": reasoning_effort,
        "web_search": web_search_enabled,
        "deep_research": deep_research_enabled,
        "web_tool_type": (
            "web_search_preview" if deep_research_enabled else "web_search"
        )
        if web_search_enabled
        else None,
        "response": _response_to_dict(response),
    }
    sources_path.write_text(json.dumps(sources_payload, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"Saved: {primer_path}")
    print(f"Saved: {sources_path}")
    elapsed = time.perf_counter() - t0
    print(f"Done in {format_seconds(elapsed)}", flush=True)
