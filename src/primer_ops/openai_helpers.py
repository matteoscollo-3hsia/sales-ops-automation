from __future__ import annotations

import re
import time
from typing import Any

from openai import APITimeoutError, NotFoundError, OpenAI, RateLimitError

_REQUEST_TIMEOUT_SECONDS = 30 * 60
_URL_RE = re.compile(r"https?://[^\s)>\"]+", re.IGNORECASE)


def _extract_output_text_from_response(response: Any) -> str | None:
    if response is None:
        return None
    if isinstance(response, dict):
        output_text = response.get("output_text")
        if isinstance(output_text, str) and output_text.strip():
            return output_text
        for item in response.get("output", []) or []:
            text = _extract_output_text_from_item(item)
            if text:
                return text
        return None
    output_text = getattr(response, "output_text", None)
    if isinstance(output_text, str) and output_text.strip():
        return output_text
    output_items = getattr(response, "output", None)
    if isinstance(output_items, list):
        for item in output_items:
            text = _extract_output_text_from_item(item)
            if text:
                return text
    return None


def _extract_output_text_from_item(item: Any) -> str | None:
    if item is None:
        return None
    if isinstance(item, dict):
        item_type = item.get("type")
        if item_type == "output_text":
            text = item.get("text")
            if isinstance(text, str) and text.strip():
                return text
        content = item.get("content")
        if isinstance(content, list):
            for part in content:
                text = _extract_output_text_from_item(part)
                if text:
                    return text
    else:
        item_type = getattr(item, "type", None)
        if item_type == "output_text":
            text = getattr(item, "text", None)
            if isinstance(text, str) and text.strip():
                return text
        content = getattr(item, "content", None)
        if isinstance(content, list):
            for part in content:
                text = _extract_output_text_from_item(part)
                if text:
                    return text
    return None


def _extract_urls_from_text(text: str) -> list[str]:
    if not text:
        return []
    urls: list[str] = []
    for match in _URL_RE.findall(text):
        cleaned = match.rstrip(").,]")
        if cleaned:
            urls.append(cleaned)
    return urls


def _extract_citations_from_response(
    response: Any, output_text: str | None = None
) -> list[str]:
    urls: list[str] = []

    def add_url(value: Any) -> None:
        if isinstance(value, str):
            cleaned = value.strip()
            if cleaned:
                urls.append(cleaned)

    def walk(obj: Any) -> None:
        if obj is None:
            return
        if isinstance(obj, dict):
            add_url(obj.get("url"))
            add_url(obj.get("source_url"))
            add_url(obj.get("source"))
            annotations = obj.get("annotations")
            if isinstance(annotations, list):
                for item in annotations:
                    walk(item)
            content = obj.get("content")
            if isinstance(content, list):
                for item in content:
                    walk(item)
        else:
            add_url(getattr(obj, "url", None))
            add_url(getattr(obj, "source_url", None))
            add_url(getattr(obj, "source", None))
            annotations = getattr(obj, "annotations", None)
            if isinstance(annotations, list):
                for item in annotations:
                    walk(item)
            content = getattr(obj, "content", None)
            if isinstance(content, list):
                for item in content:
                    walk(item)

    output_items = None
    if isinstance(response, dict):
        output_items = response.get("output")
    else:
        output_items = getattr(response, "output", None)
    if isinstance(output_items, list):
        for item in output_items:
            walk(item)
    if output_text:
        urls.extend(_extract_urls_from_text(output_text))

    seen: set[str] = set()
    deduped: list[str] = []
    for url in urls:
        if url in seen:
            continue
        seen.add(url)
        deduped.append(url)
    return deduped


def _ensure_response_text(step_entry: dict[str, Any]) -> str | None:
    existing = step_entry.get("response_text")
    if isinstance(existing, str) and existing.strip():
        return existing
    legacy = step_entry.get("output_text")
    if isinstance(legacy, str) and legacy.strip():
        step_entry["response_text"] = legacy
        return legacy
    derived = _extract_output_text_from_response(step_entry.get("response"))
    if isinstance(derived, str) and derived.strip():
        step_entry["response_text"] = derived
        return derived
    return None


def _model_supports_reasoning_effort(model: str | None) -> bool:
    if not model:
        return False
    return model.strip().lower().startswith("gpt-5")


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
            print(
                f"Request timed out after {int(_REQUEST_TIMEOUT_SECONDS / 60)} minutes."
            )
            if _confirm_continue_after_timeout():
                continue
            raise SystemExit("Aborted after timeout.")
        except RateLimitError as err:
            msg = str(err).lower()
            code = getattr(err, "code", None)
            should_retry = (code == "rate_limit_exceeded") or (
                "rate limit reached" in msg
            )
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
