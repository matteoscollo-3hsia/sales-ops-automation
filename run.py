from __future__ import annotations

import sys
from pathlib import Path


def main() -> int:
    root = Path(__file__).resolve().parent
    src_path = root / "src"
    sys.path.insert(0, str(src_path))

    from primer_ops.cli import main as cli_main

    return cli_main()


if __name__ == "__main__":
    raise SystemExit(main())
