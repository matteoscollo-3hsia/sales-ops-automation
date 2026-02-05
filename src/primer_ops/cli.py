from __future__ import annotations

import argparse

from primer_ops.lead_input import run_create_input
from primer_ops.primer import generate_primer


def main() -> int:
    parser = argparse.ArgumentParser(prog="primer-ops")
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("create-input", help="Create lead input interactively")
    gen_parser = subparsers.add_parser("generate-primer", help="Generate primer")
    gen_parser.add_argument(
        "--output-dir",
        default=None,
        help="Override OUTPUT_DIR from .env",
    )
    gen_parser.add_argument(
        "--sheet",
        default="Company and Industry Intro",
        help="Excel sheet name (default: Company and Industry Intro)",
    )

    args = parser.parse_args()

    if args.command == "create-input":
        run_create_input()
        return 0

    if args.command == "generate-primer":
        generate_primer(output_dir=args.output_dir, sheet=args.sheet)
        return 0

    parser.print_help()
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
