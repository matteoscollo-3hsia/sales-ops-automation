from pathlib import Path
import os

from dotenv import load_dotenv
from openai import OpenAI


def main() -> None:
    # Always load .env from repo root
    root = Path(__file__).resolve().parents[1]
    load_dotenv(root / ".env", override=True)

    model = os.getenv("OPENAI_MODEL", "gpt-5.2")

    client = OpenAI()
    resp = client.responses.create(
        model=model,
        input="Reply with exactly: OK",
    )

    print(resp.output_text)


if __name__ == "__main__":
    main()
