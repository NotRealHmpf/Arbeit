"""Sicherungsplanheft: wer hat es gemacht, wer fehlt?

Gleicht die Mitarbeiterliste (Spalten Personennummer, Name als "Nachname, Vorname")
mit einem oder mehreren Sicherungsplanheft-Exporten ab (.xls aus dem Lernportal, intern
eine HTML-Tabelle mit der Spalte "Name" als "Vorname Nachname"). Wer in mindestens einem
Export mit Status "Abgeschlossen" steht, gilt als erledigt.

Aufruf:
    python abgleich.py Mitarbeiter.xlsx Ausgabe.xlsx Export1.xls [Export2.xls ...]
"""
import io
import re
import sys
import unicodedata
from collections import defaultdict
from html.parser import HTMLParser
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font

FEHLT, ERLEDIGT = "Fehlt", "Erledigt"


class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.rows, self._row, self._cell = [], None, None

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
        elif tag in ("td", "th"):
            self._cell = []

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._cell is not None:
            self._row.append("".join(self._cell).strip())
            self._cell = None
        elif tag == "tr" and self._row is not None:
            self.rows.append(self._row)
            self._row = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def read_rows(path):
    """Zeilen einer Tabelle als Listen von Texten (HTML-.xls oder echtes .xlsx)."""
    data = Path(path).read_bytes()
    if data[:2] == b"PK":
        ws = load_workbook(io.BytesIO(data), read_only=True, data_only=True).active
        return [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    parser = _TableParser()
    parser.feed(data.decode("utf-8", errors="replace"))
    return parser.rows


def norm(text):
    """Vergleichsform: klein, ä=ae, ohne Akzente, nur Buchstaben/Ziffern."""
    text = text.strip().lower()
    for a, b in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss")):
        text = text.replace(a, b)
    text = "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))
    return " ".join(re.sub(r"[^a-z0-9]+", " ", text).split())


def sort_key(text):
    text = unicodedata.normalize("NFKD", text.casefold().replace("ß", "ss"))
    return "".join(c for c in text if not unicodedata.combining(c))


def keys_for(list_name):
    """'Nachname, Vorname (Zweitname)' -> Vergleichsformen als 'Vorname Nachname'."""
    aliases = [a.strip() for a in re.findall(r"\(([^)]*)\)", list_name) if a.strip()]
    base = " ".join(re.sub(r"\([^)]*\)", " ", list_name).split())
    if "," in base:
        last, first = (x.strip() for x in base.split(",", 1))
        names = [f"{first} {last}"]
    else:
        first = " ".join(base.split()[:-1])
        names = [base]
    names += [f"{first} {a}" for a in aliases]
    return [norm(n) for n in names if norm(n)]


def main(mitarbeiter, ausgabe, exporte):
    # Mitarbeiter je Personennummer (doppelte Einträge zählen einmal)
    rows = read_rows(mitarbeiter)
    header = [h.lower() for h in rows[0]]
    i_pnr, i_name = header.index("personennummer"), header.index("name")
    people = {}
    for r in rows[1:]:
        name = r[i_name] if i_name < len(r) else ""
        if not name:
            continue
        pnr = r[i_pnr] if i_pnr < len(r) else ""
        p = people.setdefault(pnr or name, dict(pnr=pnr, name=name, list_names=[], seen_as=set()))
        p["list_names"].append(name)

    exact, unordered = defaultdict(set), defaultdict(set)
    for pid, p in people.items():
        for ln in p["list_names"]:
            for k in keys_for(ln):
                exact[k].add(pid)
                unordered[tuple(sorted(k.split()))].add(pid)

    unassigned = []
    for path in exporte:
        rows = read_rows(path)
        header = [h.lower() for h in rows[0]]
        i_name = header.index("name")
        i_status = header.index("status") if "status" in header else None
        for r in rows[1:]:
            name = r[i_name] if i_name < len(r) else ""
            if not name:
                continue
            if i_status is not None and r[i_status] != "Abgeschlossen":
                unassigned.append((name, Path(path).name, f"Status „{r[i_status]}“"))
                continue
            k = norm(name)
            ids = exact.get(k) or unordered.get(tuple(sorted(k.split()))) or set()
            if len(ids) == 1:
                people[next(iter(ids))]["seen_as"].add(name)
            else:
                reason = "mehrere Mitarbeiter möglich" if ids else "nicht in der Mitarbeiterliste"
                unassigned.append((name, Path(path).name, reason))

    wb = Workbook()
    ws = wb.active
    ws.title = "Abgleich"
    ws.append(["Personennummer", "Name", "Status"])
    order = sorted(people.values(), key=lambda p: (bool(p["seen_as"]), sort_key(p["name"])))
    for p in order:
        ws.append([p["pnr"], p["name"], ERLEDIGT if p["seen_as"] else FEHLT])
        notes = []
        other = [n for n in p["seen_as"] if norm(n) != keys_for(p["name"])[0]]
        if other:
            notes.append("Im Sicherungsplanheft als " + ", ".join(f"„{n}“" for n in sorted(other)) + " eingetragen.")
        if len(p["list_names"]) > 1:
            notes.append(f"Personennummer steht {len(p['list_names'])}× in der Mitarbeiterliste: "
                         + ", ".join(f"„{n}“" for n in p["list_names"]) + ".")
        if notes:
            c = Comment(" ".join(notes), "Abgleich")
            c.width, c.height = 320, 90
            ws.cell(row=ws.max_row, column=2).comment = c
    style(ws, [18, max(30, min(45, max(len(p["name"]) for p in order) + 2)), 11])
    for cell in ws["A"][1:]:
        cell.number_format = "@"

    if unassigned:
        ws2 = wb.create_sheet("Nicht zugeordnet")
        ws2.append(["Name im Sicherungsplanheft", "Datei", "Grund"])
        for row in sorted(unassigned, key=lambda x: sort_key(x[0])):
            ws2.append(list(row))
        style(ws2, [32, 45, 30])

    wb.save(ausgabe)
    fehlt = sum(1 for p in people.values() if not p["seen_as"])
    print(f"Mitarbeiter: {len(people)}  Erledigt: {len(people) - fehlt}  Fehlt: {fehlt}  "
          f"Nicht zugeordnet: {len(unassigned)}")
    for name, datei, grund in unassigned:
        print(f"  nicht zugeordnet: {name} ({datei}): {grund}")


def style(ws, widths):
    for row in ws.iter_rows():
        for cell in row:
            cell.font = Font(name="Arial", size=10, bold=cell.row == 1)
    for col, w in zip("ABC", widths):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


if __name__ == "__main__":
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2], sys.argv[3:])
