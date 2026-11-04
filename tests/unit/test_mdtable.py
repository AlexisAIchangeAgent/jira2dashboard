"""Markdown table reader (params\\*.md and the docs share this format)."""
import pytest

from jira2dashboard.params.mdtable import MarkdownTableError, Table, parse_table, parse_tables

SIMPLE = """# Title

Some text with a | pipe that is not a table.

| A | B |
|---|---|
| 1 | two words |
| `x` | |
"""


def test_parse_table_reads_header_and_rows():
    table = parse_table(SIMPLE)
    assert table == Table(header=("A", "B"), rows=(("1", "two words"), ("x", "")))


def test_backticks_are_removed_from_cells():
    assert parse_table("| A |\n|---|\n| `code` and `more` |\n").rows == (("code and more",),)


def test_alignment_separators_are_accepted():
    assert parse_table("| A | B |\n|:---|---:|\n| 1 | 2 |\n").rows == (("1", "2"),)


def test_accents_and_slashes_are_kept():
    assert parse_table("| S |\n|---|\n| En test / recette |\n| Dépriorisé |\n").rows == (
        ("En test / recette",), ("Dépriorisé",))


def test_parse_tables_returns_every_table_in_order():
    text = "| A |\n|---|\n| 1 |\n\ntext\n\n| B | C |\n|---|---|\n| 2 | 3 |\n"
    tables = parse_tables(text)
    assert [t.header for t in tables] == [("A",), ("B", "C")]


def test_a_line_starting_with_a_pipe_without_separator_is_not_a_table():
    assert parse_tables("| just a line |\nplain text\n") == []


def test_row_with_the_wrong_number_of_cells_is_refused_with_its_line_number():
    with pytest.raises(MarkdownTableError, match="line 4: 3 cells, expected 2"):
        parse_table("| A | B |\n|---|---|\n| 1 | 2 |\n| 1 | 2 | 3 |\n")


def test_parse_table_needs_exactly_one_table():
    with pytest.raises(MarkdownTableError, match="no table"):
        parse_table("nothing here")
    with pytest.raises(MarkdownTableError, match="2 tables"):
        parse_table("| A |\n|---|\n| 1 |\n\n| B |\n|---|\n| 2 |\n")
