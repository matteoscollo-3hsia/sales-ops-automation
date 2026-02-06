from __future__ import annotations

import json
from pathlib import Path

from dotenv import find_dotenv, load_dotenv
from pydantic import BaseModel, Field, ValidationError

class LeadInput(BaseModel):
    company_name: str = Field(min_length=1)
    company_website: str = Field(default="")
    hq_country: str = Field(default="")
    industry: str = Field(default="")
    revenue_mln: float = Field(ge=0)
    primary_contact_name: str = Field(default="")
    primary_contact_role: str = Field(default="")


def prompt_str(label: str, required: bool = False) -> str:
    while True:
        val = input(f"{label}: ").strip()
        if required and not val:
            print("  -> Required field. Please enter a value.")
            continue
        return val


def prompt_float(label: str) -> float:
    while True:
        raw = input(f"{label} (number): ").strip().replace(",", ".")
        try:
            return float(raw)
        except ValueError:
            print("  -> Please enter a valid number (e.g., 75 or 75.5).")


def run_create_input(lead_output: str | None = None) -> None:
    env_path = find_dotenv(usecwd=True)
    load_dotenv(env_path, override=True)
    out_path = Path(lead_output) if lead_output else Path("lead_input.json")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    print("\n=== Lead Input Wizard (writes lead_input.json to the specified path) ===\n")

    data = {
        "company_name": prompt_str("Company name", required=True),
        "company_website": prompt_str("Company website (optional)", required=False),
        "hq_country": prompt_str("HQ country (optional)", required=False),
        "industry": prompt_str("Industry (optional)", required=False),
        "revenue_mln": prompt_float("Revenue in EUR (mln)"),
        "primary_contact_name": prompt_str("Primary contact name (optional)", required=False),
        "primary_contact_role": prompt_str("Primary contact role (optional)", required=False),
    }

    try:
        lead = LeadInput(**data)
    except ValidationError as e:
        print("\nValidation error:\n")
        print(e)
        raise SystemExit(1)

    out_path.write_text(
        json.dumps(lead.model_dump(), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    print(f"\nSaved: {out_path}\n")
