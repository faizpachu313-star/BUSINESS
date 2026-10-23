"""Shrink an openpyxl-written .xlsx and embed cached formula results.

Why: the finished workbook has to travel through the Google Drive connector as a
base64 string, so every kilobyte matters. 31 daily sheets x 60 identical formula
rows compress badly as written by openpyxl; shared formulas and dropping empty
attributes cut the file to a fraction of the size without changing a single
formula result.

Cached results are copied in from a LibreOffice-recalculated twin of the file so
viewers that do not calculate (phone previews, Drive preview, pandas) still show
numbers. Excel recalculates everything on open anyway (fullCalcOnLoad).
"""
import re
import zipfile
from xml.sax.saxutils import escape

import openpyxl

_ROW_FORMULA = re.compile(
    r'<c r="D(\d+)" s="(\d+)"><f>(IF\(AND\(ISNUMBER\(B\1\),.*?)</f><v ?/></c>')
_FORMULA_CELL = re.compile(r'<c r="([A-Z]+)(\d+)"( s="\d+")?><f( t="shared"[^>/]*)?(?:/>|>(.*?)</f>)(?:<v ?/>)?</c>')


def _share_row_formulas(xml):
    rows = [int(m.group(1)) for m in _ROW_FORMULA.finditer(xml)]
    if not rows:
        return xml
    first, last = min(rows), max(rows)

    def repl(m):
        r, s, formula = int(m.group(1)), m.group(2), m.group(3)
        if r == first:
            return (f'<c r="D{r}" s="{s}"><f t="shared" ref="D{first}:D{last}" si="0">'
                    f'{formula}</f></c>')
        return f'<c r="D{r}" s="{s}"><f t="shared" si="0"/></c>'

    return _ROW_FORMULA.sub(repl, xml)


_DEFAULT_NOISE = (
    re.compile(r'<dimension [^>]*/>'),
    re.compile(r'<selection [^>]*/>'),
    re.compile(r'<outlinePr [^>]*/>'),
    re.compile(r' (?:showDropDown|showInputMessage)="0"'),
    re.compile(r' baseColWidth="8"'),
)


def _tidy(xml):
    xml = xml.replace(' t="n" />', '/>').replace(' />', '/>').replace('<v/>', '')
    for pattern in _DEFAULT_NOISE:      # attributes/elements that only restate Excel's defaults
        xml = pattern.sub('', xml)
    return xml


def _inject_cached(xml, values):
    """values: {"D5": 12.5, ...} computed by LibreOffice for this sheet."""
    def repl(m):
        col, row, style, shared, formula = m.groups()
        v = values.get(f"{col}{row}")
        head = f'<c r="{col}{row}"{style or ""}'
        f_el = (f'<f{shared}/>' if formula is None else
                f'<f{shared or ""}>{formula}</f>')
        if v is None or v == "":
            return f'{head}>{f_el}</c>'
        if isinstance(v, bool):
            return f'{head} t="b">{f_el}<v>{int(v)}</v></c>'
        if isinstance(v, (int, float)):
            return f'{head}>{f_el}<v>{v!r}</v></c>'
        return f'{head} t="str">{f_el}<v>{escape(str(v))}</v></c>'

    return _FORMULA_CELL.sub(repl, xml)


_INLINE = re.compile(r'<c r="([A-Z]+\d+)"( s="\d+")? t="inlineStr"><is><t(?: [^>]*)?>(.*?)</t></is></c>', re.S)


def _to_shared_strings(xml, table):
    """Replace inline strings by indexes into one workbook-wide table (stored once, not per sheet)."""
    def repl(m):
        ref, style, text = m.groups()
        idx = table.setdefault(text, len(table))
        return f'<c r="{ref}"{style or ""} t="s"><v>{idx}</v></c>'
    return _INLINE.sub(repl, xml)


def _shared_strings_part(table):
    items = "".join(f"<si><t xml:space=\"preserve\">{t}</t></si>" for t in table)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'count="{len(table)}" uniqueCount="{len(table)}">{items}</sst>').encode("utf-8")


def compact(src, dst, recalculated=None):
    """Write a compact copy of src to dst.

    recalculated: path of a LibreOffice-recalculated copy of src, used to embed
    cached formula results (optional).
    """
    cached = {}
    if recalculated:
        wb = openpyxl.load_workbook(recalculated, data_only=True)
        for ws in wb.worksheets:
            cached[ws.title] = {c.coordinate: c.value for row in ws.iter_rows() for c in row
                                if c.value is not None}
        sheet_order = wb.sheetnames
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED,
                                                      compresslevel=9) as zout:
        if recalculated:
            wbxml = zin.read("xl/workbook.xml").decode("utf-8")
            rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            rid_by_name = {n: r for n, r in re.findall(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wbxml)}
            target_by_rid = {}
            for m in re.finditer(r"<Relationship [^>]*>", rels):
                tag = m.group(0)
                rid = re.search(r'Id="([^"]+)"', tag).group(1)
                tgt = re.search(r'Target="([^"]+)"', tag).group(1)
                target_by_rid[rid] = tgt.lstrip("/").replace("xl/", "", 1)
            path_to_name = {"xl/" + target_by_rid[rid]: name for name, rid in rid_by_name.items()}
        strings = {}
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.decode("utf-8").replace(
                    "</Types>", '<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>').encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = data.decode("utf-8").replace(
                    "</Relationships>", '<Relationship Id="rIdSst" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'
                ).encode("utf-8")
            if item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                xml = data.decode("utf-8")
                xml = _to_shared_strings(xml, strings)
                xml = _tidy(_share_row_formulas(xml))
                if recalculated:
                    xml = _inject_cached(xml, cached.get(path_to_name.get(item.filename), {}))
                data = xml.encode("utf-8")
            zout.writestr(item.filename, data)
        zout.writestr("xl/sharedStrings.xml", _shared_strings_part(strings))
