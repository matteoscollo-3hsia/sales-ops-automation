# codex-playground
Repository to test Sales Automation

## CLI
Create lead input:
`python run.py create-input --lead-output path/to/lead_input.json`

Generate primer:
`python run.py generate-primer --lead-input path/to/lead_input.json --output-dir path/to/output`

Lead input resolution precedence:
- `--lead-input` (if provided)
- `LEAD_INPUT_PATH` (env)
- `./lead_input.json`

Output resolution precedence:
- `--output-dir` (if provided, treated as final output folder)
- `lead_input.json`: `client_output_dir` or `output_dir`
- `OUTPUT_BASE_DIR/<company_folder>`
- `OUTPUT_DIR/<company_folder>` (legacy)

## Tests
Run:
`uv run pytest`
