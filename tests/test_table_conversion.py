# SPDX-FileCopyrightText: 2026 Jason Hastain
# SPDX-License-Identifier: Apache-2.0
"""Unit tests for markdown-table -> HTML conversion used in outbound Teams messages."""

from mcp_teams_server.teams import (
    contains_markdown_table,
    convert_markdown_tables_to_html,
)

BASIC_TABLE = """Here are the results:

| Step | Table | Pattern |
| --- | --- | --- |
| Accounts | Accounts.primaryEmail | Accounts_<id>@mask.example |
| People | People.emailAddress | People_<id>@mask.example |

That's all."""


# ---------------------------------------------------------------------------
# Detection
# ---------------------------------------------------------------------------

def test_detects_pipe_table():
    assert contains_markdown_table(BASIC_TABLE) is True


def test_detects_table_without_outer_pipes():
    content = "a | b\n--- | ---\n1 | 2"
    assert contains_markdown_table(content) is True


def test_no_false_positive_on_plain_pipes():
    # A pipe in prose is not a table (no separator row).
    assert contains_markdown_table("use a | b for the OR operator") is False


def test_no_false_positive_on_empty_string():
    assert contains_markdown_table("") is False


# ---------------------------------------------------------------------------
# Conversion
# ---------------------------------------------------------------------------

def test_converts_to_html_table():
    html = convert_markdown_tables_to_html(BASIC_TABLE)
    assert "<table" in html
    assert "<thead>" in html and "<tbody>" in html
    assert "<th align=\"left\">Step</th>" in html
    assert "<td>Accounts</td>" in html
    # Markdown pipe/separator syntax must be gone.
    assert "| --- |" not in html
    assert "| Step |" not in html
    # Surrounding prose preserved.
    assert "Here are the results:" in html
    assert "That's all." in html


def test_escapes_angle_brackets_in_cells():
    content = "| Col |\n| --- |\n| Accounts_<id>@mask.example |"
    html = convert_markdown_tables_to_html(content)
    assert "Accounts_&lt;id&gt;@mask.example" in html
    assert "<id>" not in html


def test_ragged_row_is_padded():
    content = "| A | B | C |\n| --- | --- | --- |\n| 1 | 2 |"
    html = convert_markdown_tables_to_html(content)
    # Row with 2 cells against a 3-col header should still yield 3 <td>.
    assert html.count("<td>") == 3


def test_escaped_pipe_in_cell_preserved():
    content = "| Op |\n| --- |\n| a \\| b |"
    html = convert_markdown_tables_to_html(content)
    assert "a | b" in html


def test_content_without_table_unchanged():
    content = "Just a normal message.\nNo tables here."
    assert convert_markdown_tables_to_html(content) == content


def test_preexisting_html_left_intact():
    content = "<p>hello</p>\n<b>world</b>"
    assert convert_markdown_tables_to_html(content) == content
