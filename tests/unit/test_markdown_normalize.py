from __future__ import annotations

from primer_ops.render_docx import normalize_markdown_for_docx


def test_normalize_numeric_headings() -> None:
    text = (
        "\n"
        "1. Company Introduction\n"
        "\n"
        "Paragraph text.\n\n"
        "2.1 Ownership & Governance\n"
        "\n"
        "More text.\n\n"
        "1) List item one\n"
        "2) List item two\n"
    )
    normalized = normalize_markdown_for_docx(text)
    assert "## 1. Company Introduction" in normalized
    assert "### 2.1 Ownership & Governance" in normalized
    assert "1) List item one" in normalized
    assert "2) List item two" in normalized


def test_does_not_convert_ordered_list() -> None:
    text = "1. First item\n2. Second item\n3. Third item\n"
    normalized = normalize_markdown_for_docx(text)
    assert "## 1. First item" not in normalized
    assert "### 2. Second item" not in normalized


def test_converts_numeric_heading_without_blank_after() -> None:
    text = "1. Company Introduction\nThe TJX Companies, Inc. operates stores.\n"
    normalized = normalize_markdown_for_docx(text)
    assert normalized.startswith("## 1. Company Introduction\n")


def test_heading_followed_by_bullets_converts() -> None:
    text = "3. Triggers\n- Item one\n- Item two\n"
    normalized = normalize_markdown_for_docx(text)
    assert normalized.startswith("## 3. Triggers\n")


def test_normalize_table_separator_columns() -> None:
    text = (
        "| A | B | C | D | E | F |\n"
        "| --- | --- | --- | --- | --- |\n"
        "| 1 | 2 | 3 | 4 | 5 | 6 |\n"
    )
    normalized = normalize_markdown_for_docx(text)
    lines = normalized.splitlines()
    sep_line = lines[1]
    sep_cells = [cell.strip() for cell in sep_line.strip("|").split("|")]
    assert len(sep_cells) == 6
