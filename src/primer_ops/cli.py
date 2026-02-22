from __future__ import annotations

import argparse

from primer_ops.lead_input import run_create_input
from primer_ops.primer import generate_primer


def main() -> int:
    parser = argparse.ArgumentParser(prog="primer-ops")
    subparsers = parser.add_subparsers(dest="command", required=True)

    create_parser = subparsers.add_parser(
        "create-input", help="Create lead input interactively"
    )
    create_parser.add_argument(
        "--lead-output",
        default=None,
        help="Path to write lead_input.json (overrides --company-name placement)",
    )
    create_parser.add_argument(
        "--company-name",
        "--client",
        dest="company_name",
        default=None,
        help="Company name used to place lead_input.json under the client repo layout",
    )
    gen_parser = subparsers.add_parser("generate-primer", help="Generate primer")
    gen_parser.add_argument(
        "--output-dir",
        default=None,
        help="Override output directory (use as final output folder)",
    )
    gen_parser.add_argument(
        "--lead-input",
        default=None,
        help="Path to lead_input.json (default: LEAD_INPUT_PATH env or ./lead_input.json)",
    )
    gen_parser.add_argument(
        "--sheet",
        default=None,
        help="Excel sheet name (default: run all runnable sheets)",
    )
    gen_parser.add_argument(
        "--include",
        default=None,
        help="Regex or comma-separated list of sheet names to include",
    )
    gen_parser.add_argument(
        "--exclude",
        default=None,
        help="Regex or comma-separated list of sheet names to exclude",
    )
    gen_parser.add_argument(
        "--resume",
        default=True,
        action=argparse.BooleanOptionalAction,
        help="Resume from sources.json if present (default: True).",
    )
    gen_parser.add_argument(
        "--include-headings",
        default=None,
        action="store_true",
        help="Include sheet/step headings in primer.md (default: disabled).",
    )

    args = parser.parse_args()

    if args.command == "create-input":
        run_create_input(lead_output=args.lead_output, company_name=args.company_name)
        return 0

    if args.command == "generate-primer":
        generate_primer(
            output_dir=args.output_dir,
            lead_input=args.lead_input,
            sheet=args.sheet,
            include=args.include,
            exclude=args.exclude,
            resume=args.resume,
            include_headings=args.include_headings,
        )
        return 0

    parser.print_help()
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
