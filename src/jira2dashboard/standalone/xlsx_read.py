"""Read the cells of sheets of an .xlsx or .xlsm with the standard library only (no Excel: the file is a zip of XML).

It reads what Excel saves: shared strings (rich text runs joined), inline strings, numbers, booleans, error values and
formula results. A date comes back as its Excel serial number (the caller knows which columns hold dates). A workbook
open in Excel can be read: the last saved version is read.
A part with a DOCTYPE is refused: Excel never writes one, and without a DTD no entity can be declared, so neither
external entities (XXE) nor entity expansion (billion laughs) can reach the standard XML parser.
"""
import posixpath
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree

_REF = re.compile(r"([A-Z]+)(\d+)")


class NotAWorkbook(ValueError):
    """The file is not an Excel workbook this reader accepts."""


def _xml(package: zipfile.ZipFile, part: str):
    data = package.read(part)
    if b"<!DOCTYPE" in data.upper():
        raise NotAWorkbook(f"{part} declares a DOCTYPE: refused")
    if data[:2] in (b"\xff\xfe", b"\xfe\xff") or b"\x00" in data[:8]:       # UTF-16: the check above cannot see a DOCTYPE in it
        raise NotAWorkbook(f"{part} is not UTF-8: refused")
    return ElementTree.fromstring(data)


def _local(tag: str) -> str:
    """The name of a tag without its namespace (transitional and strict OOXML use different namespaces)."""
    return tag.rsplit("}", 1)[-1]


def _children(element, name: str):
    return [child for child in element if _local(child.tag) == name]


def _text(element) -> str:
    """The text of a string item: <t>, or the <t> of every rich text run <r>, joined."""
    return "".join(node.text or "" for node in element.iter() if _local(node.tag) == "t")


def column_index(letters: str) -> int:
    """A -> 1, AA -> 27."""
    index = 0
    for letter in letters:
        index = index * 26 + ord(letter) - 64
    return index


def _sheet_paths(package: zipfile.ZipFile) -> dict[str, str]:
    """{sheet name: part path} from xl/workbook.xml and its relationships."""
    workbook = _xml(package, "xl/workbook.xml")
    rels = _xml(package, "xl/_rels/workbook.xml.rels")
    targets = {rel.get("Id"): rel.get("Target") for rel in rels}
    paths = {}
    for sheets in _children(workbook, "sheets"):
        for sheet in _children(sheets, "sheet"):
            rel_id = next(value for key, value in sheet.attrib.items() if _local(key) == "id")
            target = targets[rel_id]
            paths[sheet.get("name")] = target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target)
    return paths


def _shared_strings(package: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in package.namelist():
        return []
    root = _xml(package, "xl/sharedStrings.xml")
    return [_text(item) for item in _children(root, "si")]


def _value(cell, shared: list[str]):
    kind = cell.get("t", "n")
    value = next((child for child in cell if _local(child.tag) == "v"), None)
    if kind == "inlineStr":
        inline = next((child for child in cell if _local(child.tag) == "is"), None)
        return _text(inline) if inline is not None else ""
    if value is None or value.text is None:
        return None
    text = value.text
    if kind == "s":
        return shared[int(text)]
    if kind == "b":
        return text == "1"
    if kind in ("str", "e", "d"):
        return text
    number = float(text)
    return int(number) if number.is_integer() and "." not in text and "E" not in text.upper() else number


def read_sheets(path: Path, names=None) -> dict[str, list[list]]:
    """{sheet name: rows}, a row being the list of its cell values (None where a cell is empty), row 1 first.
    `names`: the sheets to read (default: all). A sheet that is not in the workbook is not in the result."""
    with zipfile.ZipFile(path) as package:
        paths = _sheet_paths(package)
        shared = _shared_strings(package)
        result = {}
        for name, part in paths.items():
            if names is not None and name not in names:
                continue
            root = _xml(package, part)
            rows: dict[int, dict[int, object]] = {}
            for data in _children(root, "sheetData"):
                for row in _children(data, "row"):
                    number = int(row.get("r"))
                    cells = rows.setdefault(number, {})
                    for position, cell in enumerate(_children(row, "c"), start=1):
                        ref = cell.get("r")
                        column = column_index(_REF.fullmatch(ref).group(1)) if ref else position
                        cells[column] = _value(cell, shared)
            last = max(rows, default=0)
            grid = []
            for number in range(1, last + 1):
                cells = rows.get(number, {})
                width = max(cells, default=0)
                grid.append([cells.get(c) for c in range(1, width + 1)])
            result[name] = grid
    return result
