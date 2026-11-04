"""Read Markdown tables. Pure. Inline code backticks are removed from the cells."""
import re
from dataclasses import dataclass


class MarkdownTableError(ValueError):
    """A table is malformed (wrong number of cells) or the file does not hold exactly one."""


@dataclass(frozen=True)
class Table:
    header: tuple[str, ...]
    rows: tuple[tuple[str, ...], ...]


_SEPARATOR_CELL = re.compile(r":?-{3,}:?")


def _cells(line: str) -> tuple[str, ...]:
    body = line.strip()
    if body.startswith("|"):
        body = body[1:]
    if body.endswith("|"):
        body = body[:-1]
    return tuple(cell.strip().replace("`", "") for cell in body.split("|"))


def _is_separator(cells: tuple[str, ...]) -> bool:
    return all(_SEPARATOR_CELL.fullmatch(c) for c in cells)


def _build(block: list[tuple[int, str]]) -> Table | None:
    """A block of consecutive lines starting with '|' is a table only if its 2nd line is a separator."""
    if len(block) < 2 or not _is_separator(_cells(block[1][1])):
        return None
    header = _cells(block[0][1])
    rows = []
    for number, line in block[2:]:
        cells = _cells(line)
        if len(cells) != len(header):
            raise MarkdownTableError(f"line {number}: {len(cells)} cells, expected {len(header)}")
        rows.append(cells)
    return Table(header, tuple(rows))


def parse_tables(text: str) -> list[Table]:
    """Every table of the text, in order."""
    tables: list[Table] = []
    block: list[tuple[int, str]] = []
    for number, line in enumerate(text.splitlines() + [""], start=1):
        if line.lstrip().startswith("|"):
            block.append((number, line))
            continue
        table = _build(block) if block else None
        if table:
            tables.append(table)
        block = []
    return tables


def parse_table(text: str) -> Table:
    """The one table of the text (a parameters file holds exactly one)."""
    tables = parse_tables(text)
    if not tables:
        raise MarkdownTableError("no table found")
    if len(tables) > 1:
        raise MarkdownTableError(f"{len(tables)} tables found, expected one")
    return tables[0]
