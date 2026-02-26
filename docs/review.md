# Code Review Cleanup Plan

## Context

Full review of the sales-ops-automation repo to clean up structure, remove dead code, lint/format, improve style, and add proper documentation.

---

## Phase 1: Safe Removals (no behavior change)

| Action | File | Reason |
|--------|------|--------|
| Delete | `scripts/__init__.py` | Not a package — scripts are run standalone via `python scripts/foo.py`. The `__init__.py` misleads tooling into treating `scripts/` as an importable package. |
| Delete | `src/create_lead_input.py` | Redundant with `run.py create-input` CLI — it's a thin wrapper that just calls `run_create_input()`, which is already exposed through the CLI. Dead entry point. |
| Delete | `requirements.txt` | Duplicates `pyproject.toml` dependency list; project uses `uv` exclusively (`uv sync` reads `pyproject.toml`). Having both risks drift. |
| Update | `.gitattributes` | Remove `requirements.txt text eol=lf` line — references a file that no longer exists. |
| Untrack | `out/docx_spec.json` | Generated artifact from `extract_docx_spec.py` — should not be committed. Can always be regenerated. |
| Update | `.gitignore` | Add `out/` to prevent generated output artifacts from being committed again. |
| Update | `pyproject.toml` | Remove `jinja2` and `python-pptx` — zero imports of either package in the entire codebase. Dead dependencies increase install time and attack surface. |
| Remove | `primer.py` `_find_anchor_contains` | Defined but never called anywhere in the codebase. Dead code that adds confusion. |

## Phase 2: Relocate Smoke Tests to `tests/`

**Why:** Smoke tests in `src/` are invisible to `pytest` (configured to look in `tests/`), can't use pytest fixtures like `monkeypatch` or `tmp_path`, and require manual `python src/smoke_test_foo.py` invocation. Moving them under `tests/` makes them run automatically in CI and during development.

**`src/smoke_test_output_resolution.py`** → `tests/test_output_resolution.py`

- Remove `main()` and `if __name__` block — pytest discovers `test_*` functions directly.
- Replace manual `os.environ` save/restore with `monkeypatch` fixture — safer, auto-reverts on test exit even if test crashes.
- Test functions were already pytest-compatible (`test_*` naming).

**`src/smoke_test_openai.py`** → deleted

- Originally moved to `tests/integration/test_openai_smoke.py`, but removed entirely — requires real API key and network access, not suitable for automated test suite without a proper integration test infrastructure.

**`src/smoke_test_docx.py`** → `tests/test_docx_rendering.py`

- Extract `_assert_heading_styles`, `_assert_inline_markdown_render`, `_assert_numeric_heading_normalization` into proper `test_*` functions — each assertion group becomes an independent test case with its own pass/fail status.
- Drop `main()`, `_resolve_latest_md_path`, golden comparison — these require runtime state (a previously generated primer and a non-existent golden file). Not suitable for automated tests.
- Remove `scripts.compare_docx_structure` import — only used by the dropped golden comparison code.

**`pyproject.toml`** updates:

- Add `"scripts"` to `pythonpath` — allows any test that needs to import script modules.

## Phase 3: Lint & Format

**Why:** Consistent style reduces cognitive load during code review and prevents style-related merge conflicts. Automating with `ruff` ensures future contributions stay consistent.

- Remove unused import `spinner` from `primer.py` — imported from `progress` but never referenced. Dead import.
- Replace `Optional[Any]` with `Any | None` in `primer.py` — file already uses `from __future__ import annotations`, so the modern union syntax works and is more readable. Removes need for `Optional` import.
- Fix missing blank line before `class LeadInput` in `lead_input.py` — PEP 8 requires two blank lines before top-level class definitions.
- Sort imports in all `src/primer_ops/*.py` files — stdlib → third-party → local, separated by blank lines. Standard Python convention (PEP 8 / isort).
- Add `ruff` to dev dependencies — fast, Rust-based linter+formatter. Run `ruff check --fix` + `ruff format` on entire project for consistent style.

## Phase 4: Style & Refactoring

### 4a. Extract helper modules from `primer.py` (1210 → ~1050 lines in primer.py)

**Why:** `primer.py` at 1210 lines mixes Excel cell navigation, OpenAI API plumbing, file I/O, and business logic. Extracting cohesive helper modules improves readability, makes individual concerns testable in isolation, and reduces merge conflicts when multiple people edit the file.

**New file: `src/primer_ops/excel_helpers.py`** — Excel/worksheet utilities

- `_normalize`, `_iter_cells`, `_find_anchor_exact`, `_find_anchor_exact_in_window`, `_require_anchor`, `_first_right_value`, `_replace_placeholders`, `_find_step_anchor`, `_parse_step_title`
- **Reason:** These are all pure functions that operate on openpyxl worksheets. They have no dependency on OpenAI, file I/O, or business logic.

**New file: `src/primer_ops/openai_helpers.py`** — OpenAI response parsing + API calls

- `_REQUEST_TIMEOUT_SECONDS`, `_URL_RE` constants
- `_extract_output_text_from_response`, `_extract_output_text_from_item`, `_extract_urls_from_text`, `_extract_citations_from_response`, `_ensure_response_text`
- `_model_supports_reasoning_effort`, `_is_model_not_found_error`, `_format_error_reason`
- `_confirm_continue_after_timeout`, `_call_openai_with_retries`
- **Reason:** All OpenAI-specific logic — response parsing, retry/rate-limit handling, model detection. Self-contained and independently testable.

**New file: `src/primer_ops/io_helpers.py`** — File I/O utilities

- `_safe_write_text`, `_safe_write_json`, `_safe_write_text_multi`, `_safe_write_json_multi`
- **Reason:** Atomic write operations used by both primer generation and tests. Small, pure utility module.

**`primer.py` retains:** `generate_primer()`, prompt processing (`_strip_human_reminders`, `_post_process_prompt`, `_build_prev_context_block`), output resolution (`resolve_*`), sources sanitization (`_sanitize_sources_payload`, coercion helpers), `get_initial_context`, `_is_verbose`, `_resolve_template_path`

### 4b. Minor style simplifications

- `_is_verbose()` → use `any()` generator expression — more idiomatic, eliminates explicit loop + early return.
- `_error_is_empty()` → use `not error` / `not error.strip()` idioms — more Pythonic truthiness checks.

## Phase 5: Simplify Configuration

**Why:** Multiple env vars and JSON keys that do the same thing add confusion without providing value. The codebase had `OUTPUT_DIR` (legacy) alongside `OUTPUT_BASE_DIR`, `LEAD_INPUT_PATH` as an env-only fallback, and `output_dir` in `lead_input.json` duplicating `client_output_dir`. Removing these reduces the number of concepts users need to understand.

| Action | File | Reason |
|--------|------|--------|
| Remove | `config.py` `OUTPUT_DIR_ENV`, `get_output_dir()`, `get_output_root_dir()`, `LEAD_INPUT_PATH_ENV`, `get_lead_input_path()` | Legacy/duplicate env vars. `OUTPUT_BASE_DIR` is the only persistent output config needed. Lead input is resolved via `--lead-input` flag or `./lead_input.json` default. |
| Update | `primer.py` `_extract_output_dir_override` | Remove `output_dir` key lookup — only `client_output_dir` is supported in `lead_input.json`. Eliminates ambiguity about which key to use. |
| Update | `primer.py` `resolve_lead_input_path` | Remove `LEAD_INPUT_PATH` env fallback — use `--lead-input` flag or `./lead_input.json` default. Flags and JSON files are the two config mechanisms; env vars for paths add a third that's easy to forget. |
| Update | `primer.py` `resolve_output_dir`, `resolve_output_targets` | Remove `get_output_dir()` fallback — only `get_output_base_dir()` is used. |
| Update | `lead_input.py` | Remove `get_output_dir` import and fallback in `run_create_input`. |
| Update | `cli.py` | Update help strings to remove references to `OUTPUT_DIR` and `LEAD_INPUT_PATH`. |
| Delete | `tests/integration/test_openai_smoke.py` | Requires real API key, no CI infrastructure to support it. Removed entirely rather than maintaining dead test code. |
| Update | `pyproject.toml` | Remove `integration` pytest marker — no integration tests remain. |

## Phase 6: README Rewrite

**Why:** The existing README is 25 lines with minimal information. New contributors need to understand prerequisites, setup steps, CLI flags, env vars, and project structure. A complete README reduces onboarding friction.

Expand to include:

- **Header**: What the tool does (1-2 sentences)
- **Prerequisites**: Python 3.10+, uv, OpenAI API key
- **Quickstart**: Clone → `uv sync` → configure `.env` → create lead input → generate primer
- **CLI reference**: All commands and flags
- **Path resolution**: Existing precedence docs (keep)
- **Environment variables**: Full table of config vars
- **Scripts**: Document standalone utility scripts
- **Project structure**: File tree with descriptions
- **Tests**: How to run (including integration vs. unit)


---

## Files Modified/Created/Deleted

**Deleted:** `scripts/__init__.py`, `src/create_lead_input.py`, `requirements.txt`, `src/smoke_test_*.py` (3 files), `tests/integration/` (removed entirely)

**Created:** `src/primer_ops/excel_helpers.py`, `src/primer_ops/openai_helpers.py`, `src/primer_ops/io_helpers.py`, `tests/test_output_resolution.py`, `tests/test_docx_rendering.py`, `docs/review.md`

**Modified:** `pyproject.toml`, `.gitignore`, `.gitattributes`, `src/primer_ops/primer.py`, `src/primer_ops/config.py`, `src/primer_ops/lead_input.py`, `src/primer_ops/cli.py`, `README.md`, `tests/test_primer_headings.py`, all `src/primer_ops/*.py` (import sorting/formatting)
