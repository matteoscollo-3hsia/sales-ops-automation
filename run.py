from __future__ import annotations

import os
import sys
from pathlib import Path

from dotenv import find_dotenv, load_dotenv


def main() -> int:
    root = Path(__file__).resolve().parent
    src_path = root / "src"
    sys.path.insert(0, str(src_path))

    env_path = find_dotenv(usecwd=True)
    load_dotenv(env_path, override=False)
    os.environ.setdefault(
        "PRIMER_WORD_TEMPLATE_PATH", "templates/Commercial_Primer_Template.docx"
    )

    from primer_ops.cli import main as cli_main

    return cli_main()


if __name__ == "__main__":
    raise SystemExit(main())
