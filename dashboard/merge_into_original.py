"""Setzt die von build_sheets.py erzeugten Blaetter in die Original-Arbeitsmappe ein,
ohne die Original-Blaetter anzufassen (Kommentare, Diagramme, SharePoint-Metadaten
usw. bleiben erhalten). Dazu werden die XML-Teile direkt im .xlsx-Paket ergaenzt.

Reihenfolge danach:  Uebersicht, Soll-Ist, Grafik <Bezirk> ..., <Original-Blaetter>, Daten, Zielzustand, Einstellungen
"""
import re
import zipfile
from copy import deepcopy

from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PNS = "http://schemas.openxmlformats.org/package/2006/relationships"
CTNS = "http://schemas.openxmlformats.org/package/2006/content-types"
N = "{%s}" % NS
CT_SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
REL_SHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
REL_CALCCHAIN = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"

TAIL_SHEETS = ("Daten", "Zielzustand", "Einstellungen")   # kommen ans Ende, alle anderen neuen nach vorne


def xml(data):
    return etree.fromstring(data)


def dump(el):
    return etree.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def part_number(names, prefix, suffix):
    nums = [int(m.group(1)) for n in names for m in [re.fullmatch(re.escape(prefix) + r"(\d+)" + re.escape(suffix), n)] if m]
    return max(nums, default=0)


# ------------------------------------------------------------------ Styles
def merge_styles(orig_xml, new_xml):
    o, n = xml(orig_xml), xml(new_xml)

    def sect(root, name, create_before=None):
        el = root.find(N + name)
        if el is None and create_before is not None:
            el = etree.Element(N + name)
            anchor = None
            for cand in create_before:
                anchor = root.find(N + cand)
                if anchor is not None:
                    break
            if anchor is not None:
                anchor.addprevious(el)
            else:
                root.append(el)
        return el

    def fix_count(el):
        el.set("count", str(len(el)))

    # numFmts
    numfmt_map = {}
    o_nf = sect(o, "numFmts", ["fonts"])
    used = {int(x.get("numFmtId")) for x in o_nf}
    nxt = max(used | {163}) + 1
    n_nf = n.find(N + "numFmts")
    if n_nf is not None:
        for nf in n_nf:
            old = int(nf.get("numFmtId"))
            same = [x for x in o_nf if x.get("formatCode") == nf.get("formatCode")]
            if same:
                numfmt_map[old] = int(same[0].get("numFmtId"))
                continue
            new_el = deepcopy(nf)
            new_el.set("numFmtId", str(nxt))
            o_nf.append(new_el)
            numfmt_map[old] = nxt
            nxt += 1
    fix_count(o_nf)

    maps = {}
    for name, child in (("fonts", "font"), ("fills", "fill"), ("borders", "border")):
        o_s, n_s = o.find(N + name), n.find(N + name)
        base = len(o_s)
        maps[name] = {i: base + i for i in range(len(n_s))}
        for el in n_s:
            o_s.append(deepcopy(el))
        fix_count(o_s)

    o_xf, n_xf = o.find(N + "cellXfs"), n.find(N + "cellXfs")
    base = len(o_xf)
    xf_map = {}
    for i, xf in enumerate(n_xf):
        e = deepcopy(xf)
        nid = int(e.get("numFmtId", "0"))
        e.set("numFmtId", str(numfmt_map.get(nid, nid)))
        e.set("fontId", str(maps["fonts"][int(e.get("fontId", "0"))]))
        e.set("fillId", str(maps["fills"][int(e.get("fillId", "0"))]))
        e.set("borderId", str(maps["borders"][int(e.get("borderId", "0"))]))
        e.set("xfId", "0")
        o_xf.append(e)
        xf_map[i] = base + i
    fix_count(o_xf)

    dxf_map = {}
    n_dx = n.find(N + "dxfs")
    if n_dx is not None and len(n_dx):
        o_dx = sect(o, "dxfs", ["tableStyles", "colors", "extLst"])
        base = len(o_dx)
        for i, d in enumerate(n_dx):
            o_dx.append(deepcopy(d))
            dxf_map[i] = base + i
        fix_count(o_dx)
    return dump(o), xf_map, dxf_map


def remap_sheet(sheet_xml, xf_map, dxf_map, rel_map=None):
    root = xml(sheet_xml)
    for el in root.iter(N + "c", N + "row"):
        s = el.get("s")
        if s is not None:
            el.set("s", str(xf_map[int(s)]))
    for col in root.iter(N + "col"):
        s = col.get("style")
        if s is not None:
            col.set("style", str(xf_map[int(s)]))
    for rule in root.iter(N + "cfRule"):
        d = rule.get("dxfId")
        if d is not None:
            rule.set("dxfId", str(dxf_map[int(d)]))
    return root


def rels_root(data):
    return xml(data)


# ------------------------------------------------------------------ Hauptfunktion
def merge(orig_path, new_path, out_path):
    zo = zipfile.ZipFile(orig_path)
    zn = zipfile.ZipFile(new_path)
    parts = {i.filename: zo.read(i.filename) for i in zo.infolist()}
    order = [i.filename for i in zo.infolist()]
    new_parts = {i.filename: zn.read(i.filename) for i in zn.infolist()}

    styles, xf_map, dxf_map = merge_styles(parts["xl/styles.xml"], new_parts["xl/styles.xml"])
    parts["xl/styles.xml"] = styles

    # --- Nummerierung fuer neue Teile
    sheet_base = part_number(parts, "xl/worksheets/sheet", ".xml")
    drawing_base = part_number(parts, "xl/drawings/drawing", ".xml")
    chart_base = part_number(parts, "xl/charts/chart", ".xml")

    chart_ren = {}
    for name in new_parts:
        m = re.fullmatch(r"xl/charts/chart(\d+)\.xml", name)
        if m:
            chart_ren[name] = f"xl/charts/chart{chart_base + int(m.group(1))}.xml"
    drawing_ren = {}
    for name in new_parts:
        m = re.fullmatch(r"xl/drawings/drawing(\d+)\.xml", name)
        if m:
            drawing_ren[name] = f"xl/drawings/drawing{drawing_base + int(m.group(1))}.xml"

    def norm_target(t, base_dir):
        if t.startswith("/"):
            return t[1:]
        # relativ zum Ordner des Quellteils
        segs = base_dir.split("/") + t.split("/")
        out = []
        for s in segs:
            if s == "..":
                out.pop()
            elif s and s != ".":
                out.append(s)
        return "/".join(out)

    added_ct = []
    # Charts kopieren
    for old, new in chart_ren.items():
        # Zahlenformat der Beschriftungen nicht an die Quelle koppeln (sonst zeigt Excel z. B. auch Nullen)
        parts[new] = re.sub(rb'<numFmt formatCode="([^"]*)"/>', rb'<numFmt formatCode="\1" sourceLinked="0"/>', new_parts[old])
        order.append(new)
        added_ct.append(("/" + new, CT_CHART))
    # Drawings + deren Rels
    for old, new in drawing_ren.items():
        parts[new] = new_parts[old]
        order.append(new)
        added_ct.append(("/" + new, CT_DRAWING))
        rel_old = old.replace("xl/drawings/", "xl/drawings/_rels/") + ".rels"
        if rel_old in new_parts:
            r = rels_root(new_parts[rel_old])
            for rel in r:
                tgt = norm_target(rel.get("Target"), "xl/drawings")
                rel.set("Target", "/" + chart_ren[tgt])
            rel_new = new.replace("xl/drawings/", "xl/drawings/_rels/") + ".rels"
            parts[rel_new] = dump(r)
            order.append(rel_new)

    # --- Neue Blaetter
    nwb = xml(new_parts["xl/workbook.xml"])
    nrels = rels_root(new_parts["xl/_rels/workbook.xml.rels"])
    nrel_target = {r.get("Id"): norm_target(r.get("Target"), "xl") for r in nrels}
    new_sheets = []   # (name, partname, old_index)
    for idx, sh in enumerate(nwb.find(N + "sheets")):
        rid = sh.get("{%s}id" % RNS)
        new_sheets.append((sh.get("name"), nrel_target[rid], idx))

    owb = xml(parts["xl/workbook.xml"])
    orels = rels_root(parts["xl/_rels/workbook.xml.rels"])
    o_sheets_el = owb.find(N + "sheets")
    orig_sheet_els = list(o_sheets_el)
    max_sheet_id = max(int(s.get("sheetId")) for s in orig_sheet_els)
    used_rids = {r.get("Id") for r in orels}

    front = [s for s in new_sheets if s[0] not in TAIL_SHEETS]
    tail = [s for s in new_sheets if s[0] in TAIL_SHEETS]
    final_order = [("new", s) for s in front] + [("orig", e) for e in orig_sheet_els] + [("new", s) for s in tail]

    old_new_index = {}  # Index in new.xlsx -> Index im Ergebnis
    for pos, (kind, item) in enumerate(final_order):
        if kind == "new":
            old_new_index[item[2]] = pos
    orig_index_shift = {i: len(front) + i for i in range(len(orig_sheet_els))}

    for el in list(o_sheets_el):
        o_sheets_el.remove(el)
    k = 0
    for kind, item in final_order:
        if kind == "orig":
            o_sheets_el.append(item)
            continue
        name, partname, _ = item
        k += 1
        new_part = f"xl/worksheets/sheet{sheet_base + k}.xml"
        rid = f"rId{1000 + k}"
        assert rid not in used_rids
        max_sheet_id += 1
        sh = etree.SubElement(o_sheets_el, N + "sheet")
        sh.set("name", name)
        sh.set("sheetId", str(max_sheet_id))
        sh.set("{%s}id" % RNS, rid)
        rel = etree.SubElement(orels, "{%s}Relationship" % PNS)
        rel.set("Id", rid)
        rel.set("Type", REL_SHEET)
        rel.set("Target", f"worksheets/sheet{sheet_base + k}.xml")
        # Blatt-XML: Stile umnummerieren
        root = remap_sheet(new_parts[partname], xf_map, dxf_map)
        sv = root.find(f"{N}sheetViews/{N}sheetView")
        if sv is not None:
            if name == front[0][0]:
                sv.set("tabSelected", "1")
            elif "tabSelected" in sv.attrib:
                del sv.attrib["tabSelected"]
        parts[new_part] = dump(root)
        order.append(new_part)
        added_ct.append(("/" + new_part, CT_SHEET))
        # Blatt-Rels (Drawing)
        rel_old = partname.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"
        if rel_old in new_parts:
            r = rels_root(new_parts[rel_old])
            for rr in r:
                tgt = norm_target(rr.get("Target"), "xl/worksheets")
                if tgt in drawing_ren:
                    rr.set("Target", "/" + drawing_ren[tgt])
                else:
                    raise RuntimeError(f"Unerwartete Beziehung in {rel_old}: {tgt}")
            rel_new = f"xl/worksheets/_rels/sheet{sheet_base + k}.xml.rels"
            parts[rel_new] = dump(r)
            order.append(rel_new)

    # --- Defined Names (localSheetId verschieben, neue Druckbereiche/Filter ergaenzen)
    o_dn = owb.find(N + "definedNames")
    if o_dn is None:
        o_dn = etree.Element(N + "definedNames")
        o_sheets_el.addnext(o_dn)
    for dn in o_dn:
        if dn.get("localSheetId") is not None:
            dn.set("localSheetId", str(orig_index_shift[int(dn.get("localSheetId"))]))
    n_dn = nwb.find(N + "definedNames")
    if n_dn is not None:
        for dn in n_dn:
            e = deepcopy(dn)
            if e.get("localSheetId") is not None:
                e.set("localSheetId", str(old_new_index[int(e.get("localSheetId"))]))
            o_dn.append(e)

    # --- Ansicht: Uebersicht aktiv, kein anderes Blatt ausgewaehlt
    for wv in owb.iter(N + "workbookView"):
        wv.set("activeTab", "0")
        wv.set("firstSheet", "0")
    calc = owb.find(N + "calcPr")
    if calc is None:
        calc = etree.SubElement(owb, N + "calcPr")
    calc.set("fullCalcOnLoad", "1")
    for i in range(1, sheet_base + 1):
        p = f"xl/worksheets/sheet{i}.xml"
        if p in parts and b'tabSelected="1"' in parts[p]:
            parts[p] = parts[p].replace(b' tabSelected="1"', b"", 1)

    # --- calcChain entfernen (Excel baut sie neu auf)
    for rel in list(orels):
        if rel.get("Type") == REL_CALCCHAIN:
            orels.remove(rel)
    parts.pop("xl/calcChain.xml", None)
    order = [o_ for o_ in order if o_ != "xl/calcChain.xml"]

    parts["xl/workbook.xml"] = dump(owb)
    parts["xl/_rels/workbook.xml.rels"] = dump(orels)

    # --- Content Types
    ct = xml(parts["[Content_Types].xml"])
    for ov in list(ct):
        if ov.get("PartName") == "/xl/calcChain.xml":
            ct.remove(ov)
    for pn, typ in added_ct:
        ov = etree.SubElement(ct, "{%s}Override" % CTNS)
        ov.set("PartName", pn)
        ov.set("ContentType", typ)
    parts["[Content_Types].xml"] = dump(ct)

    seen = set()
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in ["[Content_Types].xml"] + order:
            if name in seen or name not in parts:
                continue
            seen.add(name)
            zout.writestr(name, parts[name])
    return {"sheets_added": len(new_sheets), "charts": len(chart_ren), "drawings": len(drawing_ren)}


if __name__ == "__main__":
    import sys
    print(merge(sys.argv[1], sys.argv[2], sys.argv[3]))
