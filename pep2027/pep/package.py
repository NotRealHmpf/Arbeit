"""Nachbearbeitung des .xlsx-Pakets:
- Formelfehler nach der LibreOffice-Durchrechnung finden
- berechnete Werte als Cache in die (von openpyxl geschriebenen) Formelzellen und Diagramme uebernehmen,
  damit die Datei auch in der geschuetzten Ansicht/Vorschau Zahlen zeigt (Excel rechnet beim Oeffnen neu)
- .xlsm erzeugen: vbaProject.bin einbetten, Inhaltstypen und Beziehungen ergaenzen
"""
import re
import zipfile

from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PNS = "http://schemas.openxmlformats.org/package/2006/relationships"
CTNS = "http://schemas.openxmlformats.org/package/2006/content-types"
CNS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
N, C = "{%s}" % NS, "{%s}" % CNS
ERRORS = {"#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#NULL!", "#NUM!", "#N/A"}
LO_ERR = re.compile(r"^Err:\d+$")


def check_errors(calc_path, allow_na_cells=()):
    """-> Liste 'Blatt!Zelle Wert' aller Fehlerwerte (#N/A in erlaubten Bereichen ausgenommen)."""
    wb = load_workbook(calc_path, data_only=True)
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and (v in ERRORS or LO_ERR.match(v)):
                    if v == "#N/A" and (ws.title, c.column_letter) in allow_na_cells:
                        continue
                    out.append(f"{ws.title}!{c.coordinate} {v}")
    return out


def dump(el):
    return etree.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def fmt_num(v):
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, int) or (isinstance(v, float) and v.is_integer()):
        return str(int(v))
    return repr(float(v))


def set_cell_cache(c, value):
    v = c.find(N + "v")
    if v is None:
        v = etree.SubElement(c, N + "v")
    if value is None:
        value = ""
    if hasattr(value, "toordinal"):          # Datum -> serielle Zahl
        import datetime as dt
        base = dt.datetime(1899, 12, 30)
        if not isinstance(value, dt.datetime):
            value = dt.datetime(value.year, value.month, value.day)
        value = (value - base).total_seconds() / 86400
    if isinstance(value, bool):
        c.set("t", "b")
        v.text = "1" if value else "0"
    elif isinstance(value, (int, float)):
        c.attrib.pop("t", None)
        v.text = fmt_num(value)
    elif isinstance(value, str) and value in ERRORS:
        c.set("t", "e")
        v.text = value
    else:
        c.set("t", "str")
        v.text = str(value)


def ref_values(values_wb, ref):
    sheet, rng = ref.rsplit("!", 1)
    sheet = sheet.strip("'").replace("''", "'")
    ws = values_wb[sheet]
    c1, r1, c2, r2 = range_boundaries(rng.replace("$", ""))
    return [ws.cell(r, c).value for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


def fill_chart_cache(chart_xml, values_wb):
    root = etree.fromstring(chart_xml)
    for ref_el in list(root.iter(C + "numRef", C + "strRef")):
        f = ref_el.find(C + "f")
        vals = ref_values(values_wb, f.text)
        for old in ref_el.findall(C + "numCache") + ref_el.findall(C + "strCache"):
            ref_el.remove(old)
        is_text = any(isinstance(v, str) and v not in ERRORS for v in vals)
        if ref_el.tag == C + "numRef" and is_text:
            ref_el.tag = C + "strRef"
        if ref_el.tag == C + "strRef":
            cache = etree.SubElement(ref_el, C + "strCache")
            etree.SubElement(cache, C + "ptCount").set("val", str(len(vals)))
            for i, v in enumerate(vals):
                if v is None:
                    continue
                pt = etree.SubElement(cache, C + "pt")
                pt.set("idx", str(i))
                etree.SubElement(pt, C + "v").text = str(v)
        else:
            cache = etree.SubElement(ref_el, C + "numCache")
            etree.SubElement(cache, C + "formatCode").text = "General"
            etree.SubElement(cache, C + "ptCount").set("val", str(len(vals)))
            for i, v in enumerate(vals):
                if not isinstance(v, (int, float)) or isinstance(v, bool):
                    continue
                pt = etree.SubElement(cache, C + "pt")
                pt.set("idx", str(i))
                etree.SubElement(pt, C + "v").text = fmt_num(v)
    return dump(root)


def _sheet_parts(parts):
    wb = etree.fromstring(parts["xl/workbook.xml"])
    rels = etree.fromstring(parts["xl/_rels/workbook.xml.rels"])
    target = {r.get("Id"): r.get("Target") for r in rels}
    out = {}
    for sh in wb.find(N + "sheets"):
        t = target[sh.get("{%s}id" % RNS)]
        out[sh.get("name")] = t.lstrip("/") if t.startswith("/") else "xl/" + t
    return out


def inject_values(raw_path, calc_path, out_path):
    values_wb = load_workbook(calc_path, data_only=True)
    zin = zipfile.ZipFile(raw_path)
    parts = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    order = [i.filename for i in zin.infolist()]
    n_cells, charts = 0, set()
    for name, part in _sheet_parts(parts).items():
        ws = values_wb[name]
        root = etree.fromstring(parts[part])
        for c in root.iter(N + "c"):
            if c.find(N + "f") is None:
                continue
            set_cell_cache(c, ws[c.get("r")].value)
            n_cells += 1
        parts[part] = dump(root)
        srels = part.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"
        if srels in parts:
            for r in etree.fromstring(parts[srels]):
                t = r.get("Target")
                if "drawings/drawing" in t:
                    d = "xl/" + t.replace("../", "") if not t.startswith("/") else t.lstrip("/")
                    drels = d.replace("xl/drawings/", "xl/drawings/_rels/") + ".rels"
                    if drels in parts:
                        for dr in etree.fromstring(parts[drels]):
                            tt = dr.get("Target")
                            charts.add("xl/" + tt.replace("../", "") if not tt.startswith("/") else tt.lstrip("/"))
    for ch in charts:
        if ch in parts:
            parts[ch] = fill_chart_cache(parts[ch], values_wb)
    # Excel soll beim Oeffnen trotzdem vollstaendig neu rechnen
    wbx = parts["xl/workbook.xml"].decode("utf-8")
    if "fullCalcOnLoad" not in wbx:
        wbx = re.sub(r"<calcPr([^>]*)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wbx)
    parts["xl/workbook.xml"] = wbx.encode("utf-8")
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in order:
            zout.writestr(name, parts[name])
    return n_cells


def make_xlsm(xlsx_path, out_path, vba_bin):
    """vbaProject.bin einbetten und Paket als Excel-Arbeitsmappe mit Makros kennzeichnen."""
    zin = zipfile.ZipFile(xlsx_path)
    parts = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    order = [i.filename for i in zin.infolist()]
    ct = parts["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
    if 'Extension="bin"' not in ct:
        ct = ct.replace("<Default ", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ', 1)
    parts["[Content_Types].xml"] = ct.encode("utf-8")
    rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    if "vbaProject" not in rels:
        rels = rels.replace("</Relationships>",
                            '<Relationship Id="rIdVBA1" Type="http://schemas.microsoft.com/office/2006/relationships/'
                            'vbaProject" Target="vbaProject.bin"/></Relationships>')
    parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    parts["xl/vbaProject.bin"] = vba_bin
    if "xl/vbaProject.bin" not in order:
        order.append("xl/vbaProject.bin")
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in order:
            zout.writestr(name, parts[name])
