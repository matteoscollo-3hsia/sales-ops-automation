from __future__ import annotations

import json
from pathlib import Path

from primer_ops.client_repo import ensure_client_repo, sanitize_folder_name
from primer_ops.primer import (
    _safe_write_text_multi,
    resolve_lead_input_path,
    resolve_output_dir,
    resolve_output_targets,
)


def test_sanitize_folder_name() -> None:
    assert sanitize_folder_name('Acme<>:"/\\|?*  Corp') == "Acme Corp"
    assert sanitize_folder_name("Foo . ") == "Foo"
    assert sanitize_folder_name("Bar...") == "Bar"
    assert sanitize_folder_name("  Mega   Corp  ") == "Mega Corp"


def test_ensure_client_repo_paths(tmp_path: Path) -> None:
    repo = ensure_client_repo(tmp_path, 'Acme<>:"/\\|?*  Corp')
    assert repo["repo_root"].exists()
    assert repo["dossier_dir"].exists()
    assert repo["latest_dir"].exists()
    assert repo["runs_dir"].exists()
    assert repo["lead_input_path"].parent == repo["dossier_dir"]


def test_client_repo_output_targets_and_writes(tmp_path: Path, monkeypatch) -> None:
    base_dir = tmp_path / "base"
    base_dir.mkdir()

    monkeypatch.setenv("OUTPUT_BASE_DIR", str(base_dir))
    monkeypatch.delenv("OUTPUT_DIR", raising=False)

    lead = {"company_name": "Acme Corp"}
    targets = resolve_output_targets(None, lead)

    repo_root = targets["repo_root"]
    latest_dir = targets["latest_dir"]
    run_dir = targets["run_dir"]

    assert repo_root == base_dir / "Acme Corp"
    assert (repo_root / "_dossier").exists()
    assert latest_dir.exists()
    assert run_dir.exists()

    primer_content = "OK"
    _safe_write_text_multi(
        [latest_dir / "primer.md", run_dir / "primer.md"], primer_content
    )
    assert (latest_dir / "primer.md").read_text(encoding="utf-8") == primer_content
    assert (run_dir / "primer.md").read_text(encoding="utf-8") == primer_content


def test_output_dir_override_skips_client_repo(tmp_path: Path, monkeypatch) -> None:
    base_dir = tmp_path / "base"
    base_dir.mkdir()
    override_dir = tmp_path / "override"

    monkeypatch.setenv("OUTPUT_BASE_DIR", str(base_dir))
    monkeypatch.delenv("OUTPUT_DIR", raising=False)

    lead = {"company_name": "Acme Corp"}
    targets = resolve_output_targets(str(override_dir), lead)

    assert targets["repo_root"] is None
    assert override_dir.exists()
    assert not (base_dir / "Acme Corp").exists()

    primer_content = "OK"
    _safe_write_text_multi([targets["output_dir"] / "primer.md"], primer_content)
    assert (override_dir / "primer.md").read_text(encoding="utf-8") == primer_content
    assert not (base_dir / "Acme Corp" / "latest" / "primer.md").exists()


def test_lead_override_skips_client_repo(tmp_path: Path, monkeypatch) -> None:
    base_dir = tmp_path / "base"
    base_dir.mkdir()
    override_dir = tmp_path / "lead_override"

    monkeypatch.setenv("OUTPUT_BASE_DIR", str(base_dir))
    monkeypatch.delenv("OUTPUT_DIR", raising=False)

    lead = {"company_name": "Acme Corp", "client_output_dir": str(override_dir)}
    targets = resolve_output_targets(None, lead)

    assert targets["repo_root"] is None
    assert override_dir.exists()
    assert not (base_dir / "Acme Corp").exists()


def test_lead_input_resolution_independent_from_output_dir(
    tmp_path: Path, monkeypatch
) -> None:
    lead_path = tmp_path / "lead_input.json"
    lead_path.write_text(json.dumps({"company_name": "Acme"}), encoding="utf-8")

    monkeypatch.delenv("LEAD_INPUT_PATH", raising=False)
    assert resolve_lead_input_path(None) == Path("lead_input.json")

    monkeypatch.setenv("LEAD_INPUT_PATH", str(lead_path))
    assert resolve_lead_input_path(None) == lead_path

    override_path = tmp_path / "override" / "lead.json"
    assert resolve_lead_input_path(str(override_path)) == override_path

    override_output = tmp_path / "output_override"
    lead = {"company_name": "Acme"}
    assert resolve_output_dir(str(override_output), lead) == override_output
    assert resolve_lead_input_path(None) == lead_path
