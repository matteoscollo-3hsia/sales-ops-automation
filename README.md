# codex-playground

Automated commercial primer generation from Excel prompt libraries using OpenAI, with Word (DOCX) output.

## Prerequisites

- Python 3.10+
- [uv](https://docs.astral.sh/uv/) (package manager)
- OpenAI API key

## Quickstart

```bash
# Install dependencies
uv sync

# Configure environment
cp .env.example .env
# Edit .env: set OPENAI_API_KEY, PROMPT_LIBRARY_PATH, OUTPUT_BASE_DIR, PRIMER_WORD_TEMPLATE_PATH

# Create a lead input interactively
uv run python run.py create-input

# Generate a primer
uv run python run.py generate-primer --lead-input path/to/lead_input.json
```

## CLI Reference

### `create-input`

Create a lead input file interactively.

```
uv run python run.py create-input [--lead-output PATH] [--company-name NAME]
```

| Flag | Description |
|------|-------------|
| `--lead-output` | Path to write `lead_input.json` (overrides default placement) |
| `--company-name` | Company name used to place file under client repo layout |

### `generate-primer`

Generate a commercial primer from an Excel prompt library.

```
uv run python run.py generate-primer [OPTIONS]
```

| Flag | Description |
|------|-------------|
| `--output-dir` | Override output directory (used as final output folder) |
| `--lead-input` | Path to `lead_input.json` |
| `--sheet` | Run a single Excel sheet by name |
| `--include` | Regex or comma-separated list of sheet names to include |
| `--exclude` | Regex or comma-separated list of sheet names to exclude |
| `--resume` / `--no-resume` | Resume from existing `sources.json` (default: enabled) |
| `--include-headings` | Include sheet/step headings in `primer.md` (default: disabled) |

## Path Resolution

### Lead input

1. `--lead-input` (if provided)
2. `LEAD_INPUT_PATH` (env)
3. `./lead_input.json`

### Output directory

1. `--output-dir` (if provided, treated as final output folder)
2. `lead_input.json`: `client_output_dir` or `output_dir`
3. `OUTPUT_BASE_DIR/<company_folder>`
4. `OUTPUT_DIR/<company_folder>` (legacy)

## Environment Variables

| Variable | Description |
|----------|-------------|
| `OPENAI_API_KEY` | OpenAI API key |
| `OPENAI_MODEL` | Model to use (default: `gpt-5.2`) |
| `OPENAI_DEEP_RESEARCH_MODEL` | Deep research model (default: `o4-mini-deep-research`) |
| `PROMPT_LIBRARY_PATH` | Path to the Excel prompt library |
| `OUTPUT_BASE_DIR` | Base directory for per-client output repos |
| `OUTPUT_DIR` | Legacy output directory |
| `LEAD_INPUT_PATH` | Default lead input path |
| `PRIMER_WORD_TEMPLATE_PATH` | Path to the Word template for DOCX output |
| `INCLUDE_HEADINGS` | Include headings in output (`1`/`true` to enable) |

## Scripts

Standalone utility scripts in `scripts/`:

| Script | Description |
|--------|-------------|
| `compare_docx_structure.py` | Compare two DOCX files and report structural differences |
| `extract_docx_spec.py` | Extract style/structure spec from a DOCX template as JSON |

## Project Structure

```
codex-playground/
├── run.py                          # CLI entry point
├── pyproject.toml                  # Project config and dependencies
├── src/
│   └── primer_ops/
│       ├── __init__.py
│       ├── cli.py                  # Argument parsing and subcommands
│       ├── client_repo.py          # Client directory layout management
│       ├── config.py               # Environment variable helpers
│       ├── excel_helpers.py        # Excel/worksheet anchor and cell utilities
│       ├── io_helpers.py           # Atomic file write utilities
│       ├── lead_input.py           # Lead input model and interactive wizard
│       ├── openai_helpers.py       # OpenAI API calls, retries, response parsing
│       ├── primer.py               # Core primer generation orchestration
│       ├── progress.py             # Spinner and time formatting
│       └── render_docx.py          # Markdown → DOCX rendering engine
├── scripts/
│   ├── compare_docx_structure.py
│   └── extract_docx_spec.py
├── tests/
│   ├── test_docx_rendering.py      # DOCX heading/inline/normalization tests
│   ├── test_output_resolution.py   # Output path resolution tests
│   ├── test_primer_headings.py     # End-to-end primer generation test
│   ├── unit/
│   │   └── test_markdown_normalize.py
│   └── integration/
│       └── test_openai_smoke.py    # Requires OPENAI_API_KEY
└── docs/
    └── review.md                   # Code review cleanup plan
```

## Tests

```bash
# Run all tests (excluding integration)
uv run pytest -m "not integration"

# Run all tests including integration (requires OPENAI_API_KEY)
uv run pytest

# Run with verbose output
uv run pytest -v
```
