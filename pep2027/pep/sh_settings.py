"""Blatt Einstellungen: Parameter, Rotation, Teams, Status/Legende, Auswahllisten, Feiertage Hessen."""
import datetime as dt

from openpyxl.styles import Alignment

from . import model as M
from .ctx import S_SET
from .style import (BOX, INPUT_FILL, add_table, dv_list, note, put, style_range, title, widths)

EASTER = [  # Osterformel (Meeus/Jones/Butcher), gilt fuer jedes Jahr
    ("a", "MOD(Jahr,19)"), ("b", "INT(Jahr/100)"), ("c", "MOD(Jahr,100)"), ("d", "INT(B{b}/4)"),
    ("e", "MOD(B{b},4)"), ("f", "INT((B{b}+8)/25)"), ("g", "INT((B{b}-B{f}+1)/3)"),
    ("h", "MOD(19*B{a}+B{b}-B{d}-B{g}+15,30)"), ("i", "INT(B{c}/4)"), ("k", "MOD(B{c},4)"),
    ("l", "MOD(32+2*B{e}+2*B{i}-B{h}-B{k},7)"), ("m", "INT((B{a}+11*B{h}+22*B{l})/451)"),
    ("Monat", "INT((B{h}+B{l}-7*B{m}+114)/31)"), ("Tag", "MOD(B{h}+B{l}-7*B{m}+114,31)+1"),
]


def section(ws, row, text):
    put(ws, f"B{row}", text, size=12, bold=True, color="FFFFFF", bg="404040")
    for c in "CDEFGH":
        put(ws, f"{c}{row}", None, bg="404040")


def header(ws, row, cols, labels):
    for c, lab in zip(cols, labels):
        put(ws, f"{c}{row}", lab, bold=True, bg=M.HEADER_FILL, border=BOX, wrap=True)


def build(ctx):
    wb = ctx.wb
    ws = wb[S_SET]
    ws.sheet_properties.tabColor = "404040"
    widths(ws, {"A": 2, "B": 34, "C": 22, "D": 24, "E": 26, "F": 16, "G": 14, "H": 16, "I": 50})
    title(ws, "B1", "Einstellungen")
    note(ws, "B2", "Gelbe Felder sind Eingabefelder. Alle anderen Blätter rechnen mit diesen Werten.")

    # ------------------------------------------------------------ Planung
    section(ws, 3, "Planung")
    rows = [
        (4, "Planungsjahr", M.YEAR, "0", "Jahr", "Alle Kalender, Feiertage und Auswertungen folgen diesem Jahr."),
        (5, "Start der Rotation", dt.date(*M.ROTATION_START), "DD.MM.YYYY", "RotStart",
         "Vorgabe 01.01.2027 (Freitag). Die Rotation läuft über Jahresgrenzen weiter."),
        (6, "Mindestbesetzung Tagdienst (MA je Arbeitstag)", M.SOLL_TAGDIENST, "0", "SollTag",
         "Annahme – bitte an den tatsächlichen Bedarf anpassen."),
        (7, "Mindestreserve (Ausfallreserve, MA)", M.MINDESTRESERVE, "0", "Reserve",
         "Annahme: Über Soll + Reserve gilt als Überbesetzung, darunter als kritisch."),
    ]
    for r, label, val, nf, nm, hint in rows:
        put(ws, f"B{r}", label, border=BOX)
        put(ws, f"C{r}", val, bg=INPUT_FILL, nf=nf, h="center", border=BOX, bold=True)
        note(ws, f"D{r}", hint)
        ctx.name(nm, f"{S_SET}!$C${r}")
    put(ws, "B8", "Prüfung Rotationsstart", border=BOX)
    put(ws, "C8", '=IF(WEEKDAY(RotStart,2)=5,"✓ Freitag","⚠ kein Freitag – Blöcke verschieben sich")',
        border=BOX, h="center")
    put(ws, "B9", "Blattschutz Jahreskalender", border=BOX)
    put(ws, "C9", "Kennwort: pep", border=BOX, h="center")
    note(ws, "D9", "Das Kalenderraster ist immer geschützt (nur Formeln) – Einträge über die Tabellen oder mit Makros per "
         "Doppelklick/Rechtsklick.")

    # ------------------------------------------------------------ Sollbesetzung Bereitschaft
    section(ws, 11, "Sollbesetzung Bereitschaft (je Tag)")
    header(ws, 12, "BC", ["Funktion", "Soll"])
    names = {"Meister": "SollM", "AV": "SollAV", "TVT-Fahrer": "SollTVT", "Handwerker": "SollHW"}
    for i, rolle in enumerate(M.ROLLEN):
        r = 13 + i
        put(ws, f"B{r}", rolle, border=BOX)
        put(ws, f"C{r}", M.SOLL[rolle], bg=INPUT_FILL, h="center", border=BOX, nf="0")
        ctx.name(names[rolle], f"{S_SET}!$C${r}")
    put(ws, "B17", "Summe", bold=True, border=BOX)
    put(ws, "C17", "=SUM(C13:C16)", bold=True, h="center", border=BOX)
    ctx.name("lstRollen", f"{S_SET}!$B$13:$B$16")
    ctx.name("SollRollen", f"{S_SET}!$C$13:$C$16")
    ctx.name("SollBereitschaft", f"{S_SET}!$C$17")

    # ------------------------------------------------------------ Rotation
    section(ws, 19, "Rotationsfolge Nachtdienst / Bereitschaft (fest, wird nicht automatisch verändert)")
    header(ws, 20, "BCD", ["Block", "Zeitraum", "Team"])
    for i, (span, team) in enumerate(M.ROTATION):
        r = 21 + i
        put(ws, f"B{r}", i + 1, h="center", border=BOX)
        put(ws, f"C{r}", span, h="center", border=BOX)
        put(ws, f"D{r}", team, h="center", border=BOX, bold=True, bg=M.TEAM_COLOR[team],
            color=M.TEAM_FONT[team])
    note(ws, "E21", "Block 1 beginnt am Rotationsstart (Freitag).", wrap=False)
    note(ws, "E22", "Danach fortlaufend: Fr–So, Mo–Do im Wechsel.")
    note(ws, "E23", "Nach Mo–Do: Freitag frei (Vorschlag).")
    note(ws, "E24", "Nach Fr–So: Montag und Dienstag frei (Vorschlag).")
    ctx.name("RotTeam", f"{S_SET}!$D$21:$D$30")
    ctx.name("RotZeitraum", f"{S_SET}!$C$21:$C$30")
    dv_list(ws, "D21:D30", "lstTeams")

    # ------------------------------------------------------------ Teams
    section(ws, 32, "Teams")
    header(ws, 33, "BCD", ["Team", "Kürzel", "Farbe"])
    for i, t in enumerate(M.TEAMS):
        r = 34 + i
        put(ws, f"B{r}", t, border=BOX, bold=True)
        put(ws, f"C{r}", M.TEAM_SHORT[t], border=BOX, h="center")
        put(ws, f"D{r}", "#" + M.TEAM_COLOR[t], border=BOX, h="center", bg=M.TEAM_COLOR[t], color=M.TEAM_FONT[t])
    ctx.name("lstTeams", f"{S_SET}!$B$34:$B$38")
    ctx.name("TeamKurz", f"{S_SET}!$C$34:$C$38")

    # ------------------------------------------------------------ Gruppen
    section(ws, 40, "Gruppen im Jahreskalender (Reihenfolge)")
    header(ws, 41, "BCD", ["Gruppe", "Überschrift", "Farbe"])
    for i, g in enumerate(M.GROUPS):
        r = 42 + i
        put(ws, f"B{r}", g, border=BOX)
        put(ws, f"C{r}", M.GROUP_TITLE[g], border=BOX)
        put(ws, f"D{r}", "#" + M.GROUP_COLOR[g], border=BOX, h="center", bg=M.GROUP_COLOR[g], color=M.GROUP_FONT[g])
    ctx.name("lstGruppen", f"{S_SET}!$B$42:$B$48")
    ctx.name("GruppenTitel", f"{S_SET}!$C$42:$C$48")

    # ------------------------------------------------------------ Funktionen
    section(ws, 50, "Funktionen (Mitarbeiterstamm)")
    header(ws, 51, "BCD", ["Funktion", "Kategorie", "Bereitschaftsrolle"])
    kat = {"Meister": "Team", "AV": "Team", "TVT-Fahrer": "Team", "Handwerker": "Team",
           "Disponent": "Ohne Team", "Zusätzliche MA": "Ressource", "AZUBI": "Ressource"}
    for i, f in enumerate(M.FUNKTIONEN):
        r = 52 + i
        put(ws, f"B{r}", f, border=BOX)
        put(ws, f"C{r}", kat[f], border=BOX)
        put(ws, f"D{r}", "JA" if f in M.ROLLEN else "NEIN", border=BOX, h="center")
    ctx.name("lstFunktionen", f"{S_SET}!$B$52:$B$58")
    ctx.name("FunktionKat", f"{S_SET}!$C$52:$C$58")

    # ------------------------------------------------------------ Status / Legende (tblArten)
    section(ws, 60, "Status, Kalendertexte und Eingabekürzel (Legende)")
    cols = ["Code", "Bezeichnung", "Kalendertext", "Eingabekürzel", "Kategorie", "Darstellung", "Beschreibung"]
    letters = "BCDEFGH"
    for c, lab in zip(letters, cols):
        put(ws, f"{c}61", lab, bold=True)
    for i, s in enumerate(M.STATUS):
        r = 62 + i
        code, bez, txt, inp, kat_, bg, fg, b, u, it, descr = s
        put(ws, f"B{r}", code, h="center")
        put(ws, f"C{r}", bez)
        put(ws, f"D{r}", txt, h="center", bg=INPUT_FILL)
        put(ws, f"E{r}", inp, bg=INPUT_FILL if inp else None)
        put(ws, f"F{r}", kat_)
        put(ws, f"G{r}", txt if txt and txt != "!" else ("Text" if txt == "!" else ""), h="center", bg=bg,
            color=fg, bold=b, italic=it, underline="single" if u else None)
        put(ws, f"H{r}", descr)
    last = 62 + len(M.STATUS) - 1
    add_table(ws, "tblArten", f"B61:H{last}", style="TableStyleLight1")
    ctx.name("StatusText", "tblArten[Kalendertext]")
    note(ws, f"B{last + 1}", "Kalendertext und Eingabekürzel sind änderbar. Mehrere Kürzel mit Komma trennen. "
         "Eingabe „-“ im Kalender = regulärer Tagdienst (hebt Abwesenheit/Freistellung auf).")

    # ------------------------------------------------------------ Auswahllisten
    r0 = last + 3
    section(ws, r0, "Auswahllisten")
    lists = [
        ("B", "Abwesenheitsart", M.ARTEN_AUSWAHL, "lstArtenAuswahl"),
        ("C", "Ressourcenart", M.RES_ARTEN, "lstResArten"),
        ("D", "JA/NEIN", ["JA", "NEIN"], "lstJaNein"),
        ("E", "Monat", M.MONATE, "lstMonate"),
        ("F", "Monat kurz", M.MON_KURZ, "lstMonKurz"),
        ("G", "Wochentag", M.WT, "lstWT"),
        ("H", "Bereitschaftsrolle", M.ROLLEN, "lstVertrFunktion"),
    ]
    for c, lab, vals, nm in lists:
        put(ws, f"{c}{r0 + 1}", lab, bold=True, bg=M.HEADER_FILL, border=BOX)
        for i, v in enumerate(vals):
            put(ws, f"{c}{r0 + 2 + i}", v, border=BOX)
        ctx.name(nm, f"{S_SET}!${c}${r0 + 2}:${c}${r0 + 1 + len(vals)}")

    # ------------------------------------------------------------ Feiertage Hessen
    f0 = r0 + 16
    section(ws, f0, "Feiertage Hessen (automatisch für das Planungsjahr)")
    fcols = ["Feiertag", "Datum", "Wochentag", "Berücksichtigen", "Datum aktiv", "Gesetzlich (HE)"]
    for c, lab in zip("BCDEFG", fcols):
        put(ws, f"{c}{f0 + 1}", lab, bold=True)
    easter_row0 = f0 + len(M.FEIERTAGE) + 5
    for i, (nm, formula, legal) in enumerate(M.FEIERTAGE):
        r = f0 + 2 + i
        put(ws, f"B{r}", nm)
        put(ws, f"C{r}", "=" + formula.replace("Ostern", "Ostern"), nf="DD.MM.YYYY", h="center")
        put(ws, f"D{r}", f'=INDEX(lstWT,WEEKDAY(C{r},2))', h="center")
        put(ws, f"E{r}", "JA" if legal else "NEIN", h="center", bg=INPUT_FILL)
        put(ws, f"F{r}", f'=IF(E{r}="JA",C{r},0)', nf="DD.MM.YYYY;;", h="center")
        put(ws, f"G{r}", "JA" if legal else "NEIN", h="center")
    flast = f0 + 1 + len(M.FEIERTAGE)
    add_table(ws, "tblFeiertage", f"B{f0 + 1}:G{flast}", style="TableStyleLight1")
    dv_list(ws, f"E{f0 + 2}:E{flast}", "lstJaNein")
    ctx.name("FT_Name", "tblFeiertage[Feiertag]")
    ctx.name("FT_Datum", "tblFeiertage[Datum]")
    ctx.name("FT_Aktiv", "tblFeiertage[Datum aktiv]")
    note(ws, f"B{flast + 1}", "Gesetzliche Feiertage in Hessen sind mit JA vorbelegt. Weitere Tage (z. B. Heiligabend, "
         "Silvester) bei Bedarf auf JA setzen – sie gelten dann als arbeitsfreie Feiertage.")

    # ------------------------------------------------------------ Osterberechnung
    e0 = flast + 3
    section(ws, e0, "Osterberechnung (Gauß/Meeus) – nicht ändern")
    rows = {}
    for i, (lab, _) in enumerate(EASTER):
        rows[lab] = e0 + 1 + i
    for lab, formula in EASTER:
        r = rows[lab]
        put(ws, f"C{r}", lab, h="right", color="595959")
        put(ws, f"B{r}", "=" + formula.format(**rows), h="center", color="595959")
    re_ = e0 + 1 + len(EASTER)
    put(ws, f"C{re_}", "Ostersonntag", h="right", bold=True)
    put(ws, f"B{re_}", f"=DATE(Jahr,B{rows['Monat']},B{rows['Tag']})", nf="DD.MM.YYYY", h="center", bold=True)
    ctx.name("Ostern", f"{S_SET}!$B${re_}")
    ws.freeze_panes = "A3"
    return ws
