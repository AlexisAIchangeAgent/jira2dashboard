"""Write an .xlsx with the standard library only: sheets of values, and SVG pictures (no Excel, no VBA, no formula).

Cells: text (inline strings), numbers, booleans, dates (Excel serial numbers with a date style). A sheet with a header
gets a bold header, a frozen first row and a filter. A picture is an SVG with a blank PNG beside it, as Office writes
them: Excel 365 draws the SVG (shape type msoGraphic), an old Excel would show the blank PNG.
The file is written next to the target, then renamed over it: a target still open in Excel is left as it was.
"""
import datetime as dt
import os
import re
import struct
import zipfile
import zlib
from dataclasses import dataclass, field
from pathlib import Path
from xml.sax.saxutils import escape

EPOCH = dt.datetime(1899, 12, 30)
EMU_PER_PIXEL = 9525
MAX_TEXT = 32767                                     # Excel's limit for one cell
_ILLEGAL = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
SVG_EXTENSION = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
STYLE_DATETIME, STYLE_DATE, STYLE_HEADER = 1, 2, 3

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


class XlsxLocked(OSError):
    """The target is open in another program (Excel): it was not replaced."""


@dataclass(frozen=True)
class Picture:
    name: str
    title: str
    svg: str
    width: int = 640                                 # pixels
    height: int = 380


@dataclass
class Sheet:
    name: str
    header: tuple[str, ...] = ()
    rows: list[tuple] = field(default_factory=list)
    pictures: list[Picture] = field(default_factory=list)
    pictures_per_row: int = 2


def column_letter(index: int) -> str:
    """1 -> A, 27 -> AA."""
    letters = ""
    while index:
        index, rest = divmod(index - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def excel_serial(value: dt.datetime | dt.date) -> float:
    if not isinstance(value, dt.datetime):
        value = dt.datetime(value.year, value.month, value.day)
    return (value - EPOCH).total_seconds() / 86400


def clean_text(value: str) -> str:
    """Text Excel accepts in a cell: no control character, at most 32767 characters."""
    return _ILLEGAL.sub("", value)[:MAX_TEXT]


def _cell(ref: str, value, header: bool = False) -> str:
    if value is None or value == "":
        return ""
    style = f' s="{STYLE_HEADER}"' if header else ""
    if isinstance(value, bool):
        return f'<c r="{ref}" t="b"{style}><v>{int(value)}</v></c>'
    if isinstance(value, dt.datetime):
        return f'<c r="{ref}" s="{STYLE_DATETIME}"><v>{excel_serial(value)!r}</v></c>'
    if isinstance(value, dt.date):
        return f'<c r="{ref}" s="{STYLE_DATE}"><v>{excel_serial(value)!r}</v></c>'
    if isinstance(value, (int, float)):
        if value != value or value in (float("inf"), float("-inf")):        # NaN or infinite: not a number Excel takes
            return ""
        return f'<c r="{ref}"{style}><v>{value!r}</v></c>'
    text = escape(clean_text(str(value)))
    space = ' xml:space="preserve"' if text != text.strip() else ""
    return f'<c r="{ref}" t="inlineStr"{style}><is><t{space}>{text}</t></is></c>'


def _width(header: str, rows: list[tuple], column: int) -> float:
    longest = max([len(header)] + [len(str(r[column])) for r in rows[:200] if column < len(r) and r[column] is not None])
    return max(10.0, min(60.0, longest + 2.0))


def _sheet_xml(sheet: Sheet, drawing: bool) -> str:
    lines = [f'{XML_HEAD}<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_R}">']
    width = max([len(sheet.header)] + [len(r) for r in sheet.rows] or [0])
    if sheet.header:
        lines.append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft"'
                     ' state="frozen"/></sheetView></sheetViews>')
        cols = "".join(f'<col min="{n}" max="{n}" width="{_width(h, sheet.rows, n - 1)}" customWidth="1"/>'
                       for n, h in enumerate(sheet.header, start=1))
        lines.append(f"<cols>{cols}</cols>")
    lines.append("<sheetData>")
    number = 0
    if sheet.header:
        number = 1
        cells = "".join(_cell(f"{column_letter(c)}1", h, header=True) for c, h in enumerate(sheet.header, start=1))
        lines.append(f'<row r="1">{cells}</row>')
    for row in sheet.rows:
        number += 1
        cells = "".join(_cell(f"{column_letter(c)}{number}", v) for c, v in enumerate(row, start=1))
        lines.append(f'<row r="{number}">{cells}</row>')
    lines.append("</sheetData>")
    if sheet.header and width:
        lines.append(f'<autoFilter ref="A1:{column_letter(width)}{max(number, 1)}"/>')
    if drawing:
        lines.append('<drawing r:id="rId1"/>')
    lines.append("</worksheet>")
    return "\n".join(lines)


def blank_png(width: int, height: int) -> bytes:
    """A white PNG (zlib + CRC): the fallback beside an SVG picture."""
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    row = b"\x00" + b"\xff\xff\xff" * width
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(row * height, 9)) + chunk(b"IEND", b""))


def _attr(text: str) -> str:
    """Text for an XML attribute in double quotes: escape() leaves the double quote alone."""
    return escape(text, {'"': "&quot;"})


def _drawing_xml(sheet: Sheet, first_rel: int) -> str:
    """Pictures in a grid, `pictures_per_row` across; a default column is 64 px wide, a default row 20 px high."""
    anchors = []
    for n, picture in enumerate(sheet.pictures):
        across, down = n % sheet.pictures_per_row, n // sheet.pictures_per_row
        col = across * (picture.width // 64 + 1)
        row = 2 + down * (picture.height // 20 + 1)
        cx, cy = picture.width * EMU_PER_PIXEL, picture.height * EMU_PER_PIXEL
        png_rel, svg_rel = first_rel + 2 * n, first_rel + 2 * n + 1
        anchors.append(
            f'<xdr:oneCellAnchor><xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row>'
            f'<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="{cx}" cy="{cy}"/><xdr:pic><xdr:nvPicPr>'
            f'<xdr:cNvPr id="{n + 2}" name="{_attr(picture.name)}" descr="{_attr(picture.title)}"/>'
            f'<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill>'
            f'<a:blip r:embed="rId{png_rel}"><a:extLst><a:ext uri="{SVG_EXTENSION}"><asvg:svgBlip '
            f'xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId{svg_rel}"/></a:ext>'
            f'</a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm>'
            f'<a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'</xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>')
    return (f'{XML_HEAD}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{NS_R}">'
            + "".join(anchors) + "</xdr:wsDr>")


def _rels(targets: list[tuple[str, str]]) -> str:
    items = "".join(f'<Relationship Id="rId{n}" Type="{REL}{kind}" Target="{target}"/>'
                    for n, (kind, target) in enumerate(targets, start=1))
    return f'{XML_HEAD}<Relationships xmlns="{NS_REL}">{items}</Relationships>'


STYLES = (f'{XML_HEAD}<styleSheet xmlns="{NS_MAIN}">'
          '<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm"/>'
          '<numFmt numFmtId="165" formatCode="yyyy-mm-dd"/></numFmts>'
          '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
          '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
          '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
          '</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
          '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
          '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
          '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')


def xlsx_parts(sheets: list[Sheet]) -> dict[str, bytes]:
    """{part name: content} of the package (pure: tests read it without writing a file)."""
    names = [s.name for s in sheets]
    if len(set(names)) != len(names) or any(len(n) > 31 or re.search(r"[\[\]:*?/\\]", n) for n in names):
        raise ValueError(f"sheet names must be unique, 31 characters at most, without []:*?/\\ : {names}")
    parts: dict[str, bytes] = {}
    overrides = ['<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.'
                 'spreadsheetml.sheet.main+xml"/>',
                 '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.'
                 'spreadsheetml.styles+xml"/>']
    drawings = 0
    media = 0
    for n, sheet in enumerate(sheets, start=1):
        has_drawing = bool(sheet.pictures)
        parts[f"xl/worksheets/sheet{n}.xml"] = _sheet_xml(sheet, has_drawing).encode("utf-8")
        overrides.append(f'<Override PartName="/xl/worksheets/sheet{n}.xml" ContentType="application/'
                         'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        if not has_drawing:
            continue
        drawings += 1
        parts[f"xl/worksheets/_rels/sheet{n}.xml.rels"] = _rels(
            [("drawing", f"../drawings/drawing{drawings}.xml")]).encode("utf-8")
        parts[f"xl/drawings/drawing{drawings}.xml"] = _drawing_xml(sheet, 1).encode("utf-8")
        overrides.append(f'<Override PartName="/xl/drawings/drawing{drawings}.xml" ContentType="application/'
                         'vnd.openxmlformats-officedocument.drawing+xml"/>')
        targets = []
        for picture in sheet.pictures:
            media += 1
            parts[f"xl/media/image{media}.png"] = blank_png(picture.width, picture.height)
            parts[f"xl/media/image{media}.svg"] = picture.svg.encode("utf-8")
            targets += [("image", f"../media/image{media}.png"), ("image", f"../media/image{media}.svg")]
        parts[f"xl/drawings/_rels/drawing{drawings}.xml.rels"] = _rels(targets).encode("utf-8")
    sheet_list = "".join(f'<sheet name="{_attr(s.name)}" sheetId="{n}" r:id="rId{n}"/>'
                         for n, s in enumerate(sheets, start=1))
    parts["xl/workbook.xml"] = (f'{XML_HEAD}<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_R}"><sheets>{sheet_list}</sheets>'
                                "</workbook>").encode("utf-8")
    parts["xl/_rels/workbook.xml.rels"] = _rels(
        [("worksheet", f"worksheets/sheet{n}.xml") for n in range(1, len(sheets) + 1)]
        + [("styles", "styles.xml")]).encode("utf-8")
    parts["xl/styles.xml"] = STYLES.encode("utf-8")
    parts["_rels/.rels"] = _rels([("officeDocument", "xl/workbook.xml")]).encode("utf-8")
    parts["[Content_Types].xml"] = (
        f'{XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/><Default Extension="svg" ContentType="image/svg+xml"/>'
        + "".join(overrides) + "</Types>").encode("utf-8")
    return parts


def write_xlsx(path: Path, sheets: list[Sheet]) -> Path:
    """Write the workbook at `path`. Raises XlsxLocked when the old file is open in Excel (it stays as it was)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    parts = xlsx_parts(sheets)
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as package:
        for name in ("[Content_Types].xml", "_rels/.rels"):
            package.writestr(name, parts.pop(name))
        for name, content in parts.items():
            package.writestr(name, content)
    try:
        os.replace(temporary, path)
    except PermissionError as err:
        temporary.unlink(missing_ok=True)
        raise XlsxLocked(f"close {path.name} first: it is open in another program, it was not replaced") from err
    return path
