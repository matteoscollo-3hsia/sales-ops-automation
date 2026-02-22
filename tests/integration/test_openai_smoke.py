from __future__ import annotations

import os
from pathlib import Path

import pytest
from dotenv import load_dotenv
from openai import OpenAI


@pytest.mark.integration
def test_openai_connection() -> None:
    root = Path(__file__).resolve().parents[2]
    load_dotenv(root / ".env", override=True)

    model = os.getenv("OPENAI_MODEL", "gpt-5.2")

    client = OpenAI()
    resp = client.responses.create(
        model=model,
        input="Reply with exactly: OK",
    )

    assert resp.output_text.strip() == "OK"
