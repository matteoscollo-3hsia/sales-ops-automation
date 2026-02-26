from __future__ import annotations

from pathlib import Path

from primer_ops.render_docx import render_primer_docx
from dump_docx_runs import assert_run_thresholds, summarize_docx_runs


def test_inline_markdown_runs_are_preserved(tmp_path: Path) -> None:
    md_path = tmp_path / "inline.md"
    docx_path = tmp_path / "inline.docx"
    md_path.write_text(
        (
            "Paragraph with **bold**, *italic*, `code`, and "
            "[link](https://example.com).\n\n"
            "| Left | Right |\n"
            "| --- | --- |\n"
            "| **table bold** | *table italic* |\n"
        ),
        encoding="utf-8",
    )

    render_primer_docx(str(md_path), str(docx_path), None)
    summary = summarize_docx_runs(docx_path)
    assert_run_thresholds(summary, min_bold=2, min_italic=2, min_code=1)

    paragraph_blob = "\n".join(summary["paragraph_text"])
    assert "Paragraph with bold, italic, code, and link." in paragraph_blob

    table_blob = "\n".join(summary["table_cell_text"])
    assert "table bold" in table_blob
    assert "table italic" in table_blob
