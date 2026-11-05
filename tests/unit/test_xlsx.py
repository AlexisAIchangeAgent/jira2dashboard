"""The xlsx of the Python path, written and read back with the standard library only (no Excel)."""
import datetime as dt
import struct
import zipfile
import zlib
from xml.etree import ElementTree

import pytest

from jira2dashboard.standalone import xlsx_read as R
from jira2dashboard.standalone import xlsx_write as W

SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="640" height="380"><rect width="10" height="10"/></svg>'


def test_column_letters_both_ways():
    for index, letters in ((1, "A"), (26, "Z"), (27, "AA"), (52, "AZ"), (703, "AAA")):
        assert W.column_letter(index) == letters and R.column_index(letters) == index


def test_dates_are_excel_serial_numbers():
    assert W.excel_serial(dt.datetime(1899, 12, 31)) == 1
    assert W.excel_serial(dt.datetime(2026, 11, 3, 18, 0)) == pytest.approx(46329.75)
    assert W.excel_serial(dt.date(2026, 11, 3)) == 46329


def test_values_come_back_as_they_were_written(tmp_path):
    rows = [("Key", 3, 2.5, True, dt.datetime(2026, 11, 3, 18, 0), dt.date(2026, 6, 1), None, " spaced ", "Terminé"),
            ("<&>", 0, -1.25, False, None, None, "", "a\x07b", "x" * 40000)]
    path = W.write_xlsx(tmp_path / "out.xlsx", [W.Sheet("Data", ("A", "B", "C", "D", "E", "F", "G", "H", "I"), rows),
                                                 W.Sheet("Empty")])
    sheets = R.read_sheets(path)
    assert list(sheets) == ["Data", "Empty"] and sheets["Empty"] == []
    header, first, second = sheets["Data"]
    assert header == ["A", "B", "C", "D", "E", "F", "G", "H", "I"]
    assert first == ["Key", 3, 2.5, True, pytest.approx(46329.75), 46174, None, " spaced ", "Terminé"]
    assert second[:6] == ["<&>", 0, -1.25, False, None, None]
    assert second[7] == "ab" and len(second[8]) == W.MAX_TEXT          # no control character, Excel's length limit


def test_only_the_sheets_asked_for_are_read(tmp_path):
    path = W.write_xlsx(tmp_path / "out.xlsx", [W.Sheet("One", ("A",), [(1,)]), W.Sheet("Two", ("B",), [(2,)])])
    assert R.read_sheets(path, {"Two", "Missing"}) == {"Two": [["B"], [2]]}


def test_a_header_row_is_bold_frozen_and_filtered():
    sheet = W.xlsx_parts([W.Sheet("Data", ("Key", "Count"), [("A-1", 1), ("A-2", 2)])])["xl/worksheets/sheet1.xml"]
    text = sheet.decode()
    assert 'state="frozen"' in text and '<autoFilter ref="A1:B3"/>' in text
    assert f'<c r="A1" t="inlineStr" s="{W.STYLE_HEADER}">' in text


def test_every_part_is_well_formed_xml():
    parts = W.xlsx_parts([W.Sheet("Run", ("Item", "Value"), [("a", 1)]),
                          W.Sheet("Charts", rows=[("charts",)], pictures=[W.Picture("velocity", "Velocity", SVG)])])
    for name, content in parts.items():
        if name.endswith((".xml", ".rels", ".svg")):
            ElementTree.fromstring(content)


def test_a_picture_is_an_svg_with_a_png_fallback_as_office_writes_it():
    parts = W.xlsx_parts([W.Sheet("Charts", pictures=[W.Picture("velocity", "Velocity", SVG),
                                                      W.Picture("flow_time", "Flow Time", SVG)])])
    assert parts["xl/media/image1.svg"] == SVG.encode() and parts["xl/media/image2.svg"] == SVG.encode()
    drawing = parts["xl/drawings/drawing1.xml"].decode()
    assert drawing.count("<xdr:pic>") == 2 and W.SVG_EXTENSION in drawing and 'name="velocity"' in drawing
    rels = parts["xl/drawings/_rels/drawing1.xml.rels"].decode()
    assert 'Target="../media/image1.png"' in rels and 'Target="../media/image1.svg"' in rels
    assert b'Extension="svg" ContentType="image/svg+xml"' in parts["[Content_Types].xml"]
    assert b'<drawing r:id="rId1"/>' in parts["xl/worksheets/sheet1.xml"]


def test_the_png_fallback_is_a_valid_png():
    data = W.blank_png(3, 2)
    assert data.startswith(b"\x89PNG\r\n\x1a\n")
    position, kinds = 8, []
    while position < len(data):
        length, kind = struct.unpack(">I4s", data[position:position + 8])
        body = data[position + 8:position + 8 + length]
        (crc,) = struct.unpack(">I", data[position + 8 + length:position + 12 + length])
        assert crc == zlib.crc32(kind + body) & 0xFFFFFFFF
        kinds.append(kind)
        if kind == b"IHDR":
            assert struct.unpack(">II", body[:8]) == (3, 2)
        position += 12 + length
    assert kinds == [b"IHDR", b"IDAT", b"IEND"]


def test_sheet_names_follow_the_rules_of_excel():
    for names in (("Same", "Same"), ("x" * 32,), ("a/b",)):
        with pytest.raises(ValueError):
            W.xlsx_parts([W.Sheet(n) for n in names])


def test_a_target_open_in_excel_is_left_as_it_was(tmp_path, monkeypatch):
    target = tmp_path / "jira2dashboard.xlsx"
    target.write_bytes(b"the previous results")

    def refused(*_args):
        raise PermissionError(13, "The process cannot access the file")
    monkeypatch.setattr(W.os, "replace", refused)
    with pytest.raises(W.XlsxLocked, match="close jira2dashboard.xlsx first"):
        W.write_xlsx(target, [W.Sheet("Run")])
    assert target.read_bytes() == b"the previous results" and list(tmp_path.iterdir()) == [target]


def excel_like(tmp_path, sheet_xml: str, shared: str = "", shared_part: str = ""):
    """A workbook as Excel saves it: shared strings with rich text, a relative or absolute sheet target."""
    path = tmp_path / "saved.xlsm"
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("xl/workbook.xml", f'<workbook xmlns="{main}" xmlns:r="{rel}"><sheets>'
                                            '<sheet name="Config" sheetId="1" r:id="rId7"/></sheets></workbook>')
        package.writestr("xl/_rels/workbook.xml.rels",
                         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                         f'<Relationship Id="rId7" Type="{rel}/worksheet" Target="/xl/worksheets/sheet3.xml"/>'
                         '</Relationships>')
        package.writestr("xl/worksheets/sheet3.xml", f'<worksheet xmlns="{main}"><sheetData>{sheet_xml}</sheetData>'
                                                     "</worksheet>")
        if shared or shared_part:
            package.writestr("xl/sharedStrings.xml", shared_part or f'<sst xmlns="{main}">{shared}</sst>')
    return path


def test_a_workbook_saved_by_excel_is_read(tmp_path):
    path = excel_like(tmp_path,
                      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>'
                      '<row r="3"><c r="A3" t="s"><v>2</v></c><c r="C3" s="5"><v>46329.75</v></c>'
                      '<c r="D3" t="b"><v>1</v></c><c r="E3" t="e"><v>#N/A</v></c><c r="F3" t="str"><f>A1</f>'
                      '<v>Parameter</v></c><c r="G3"><v>60</v></c><c r="H3"><v>1E-3</v></c></row>',
                      '<si><t>Parameter</t></si><si><r><t>Va</t></r><r><rPr><b/></rPr><t>lue</t></r></si>'
                      '<si><t xml:space="preserve">AsOf </t></si>')
    assert R.read_sheets(path) == {"Config": [["Parameter", "Value"], [],
                                              ["AsOf ", None, 46329.75, True, "#N/A", "Parameter", 60, 0.001]]}


def test_a_part_that_declares_a_doctype_is_refused(tmp_path):
    path = excel_like(tmp_path, '<row r="1"><c r="A1" t="s"><v>0</v></c></row>',
                      shared_part='<!DOCTYPE sst [<!ENTITY lol "lol">]><sst><si><t>&lol;</t></si></sst>')
    with pytest.raises(R.NotAWorkbook, match="DOCTYPE"):
        R.read_sheets(path)
