"""Eingabe- und Stammdatenblaetter: Einstellungen, Mitarbeiterstamm, Team- und
Funktionsmatrix, Qualifikationen, Urlaub und Abwesenheiten, Vertretungen, Zeitbuchungen."""
from datetime import date, time

from pp_base import (ABW_ARTEN, ABW_CODES, CAP_ABW, CAP_FM, CAP_MA, CAP_Q, CAP_VERT, CAP_ZB,
                     DATE_FMT, DUTY_CODES, FKT_KURZ, FUNKTIONEN, GREY, HEAD, HEAD_DARK, INPUT_BG,
                     MONATE, MUTED, NMAX, PFLICHT, ROTATION, SCHULFERIEN, SOLL, TEAMFUNKTIONEN,
                     TEAMS, THIN, U_NAMES, UNTERRICHTE, YEAR, L, col_widths, fill, header_row,
                     put, style, team_cf, title, weekday_formula)

X = "_xlfn."


# ====================================================================== Einstellungen
def build_settings(book, ws):
    title(ws, "Einstellungen – Personal- und Einsatzplanung",
          "Gelb hinterlegte Zellen sind Eingaben. Alle übrigen Werte werden berechnet. "
          "Werte mit „Annahme“ bitte fachlich prüfen.")
    col_widths(ws, {"A": 2, "B": 40, "C": 14, "D": 46, "E": 3, "F": 27, "G": 12, "H": 14,
                    "I": 12, "J": 20, "K": 3, "L": 22})
    for c in range(13, 26):
        ws.column_dimensions[L(c)].width = 11

    def section(r, text):
        put(ws, f"B{r}", text, size=11, bold=True, color="FFFFFF", bg=HEAD_DARK)
        for c in "CD":
            ws[f"{c}{r}"].fill = fill(HEAD_DARK)

    def kv(r, label, value, name=None, fmt=None, inp=False, note=None):
        put(ws, f"B{r}", label, border=THIN)
        c = put(ws, f"C{r}", value, border=THIN, h="center", fmt=fmt,
                bg=INPUT_BG if inp else None, bold=inp)
        if note:
            put(ws, f"D{r}", note, size=8, color=MUTED, italic=True, wrap=True)
        if name:
            book.name_ws(name, ws, f"C{r}")
        return c

    section(4, "Allgemein")
    kv(5, "Planungsjahr", YEAR, "Jahr", inp=True, note="Steuert Kalender, Feiertage und alle Auswertungen.")
    kv(6, "Jahresbeginn", "=DATE(Jahr,1,1)", "Jahresbeginn", DATE_FMT)
    kv(7, "Jahresende", "=DATE(Jahr,12,31)", "Jahresende", DATE_FMT)
    kv(8, "Tage im Jahr", "=Jahresende-Jahresbeginn+1", "TageJahr")
    kv(9, "Stichtag für Auswertungen", "=MAX(Jahresbeginn,MIN(Jahresende,TODAY()))", "Stichtag",
       DATE_FMT, inp=True,
       note="Standard: heute (innerhalb des Planungsjahres). Kann mit einem festen Datum überschrieben werden.")
    kv(10, "Ostersonntag (berechnet)", "=FLOOR(DATE(Jahr,5,DAY(MINUTE(Jahr/38)/2+56)),7)-34",
       "Ostersonntag", DATE_FMT, note="Grundlage der beweglichen Feiertage (gültig 1900–2203).")

    section(12, "Rotation Nachtdienst / Bereitschaft")
    kv(13, "Rotationsbeginn", date(2027, 1, 1), "RotStart", DATE_FMT, inp=True,
       note="Erster Block der Rotationsfolge (Fr–So). Rotation läuft danach fortlaufend weiter.")
    kv(14, "Prüfung Rotationsbeginn",
       '=IF(WEEKDAY(RotStart,2)=5,"✔ Freitag – passt zur Folge","⚠ Rotationsbeginn muss ein Freitag sein")')
    put(ws, "B15", "Die Rotationsfolge (rechts) ist fest vorgegeben und wird nicht automatisch verändert.",
        size=8, color=MUTED, italic=True)

    section(17, "Dienstzeiten und Zeitkonten")
    kv(18, "Tagdienst Beginn", time(6, 45), "TD_Beginn", "hh:mm", inp=True)
    kv(19, "Tagdienst Ende", time(15, 30), "TD_Ende", "hh:mm", inp=True)
    kv(20, "Pause Tagdienst", time(0, 45), "TD_Pause", "hh:mm", inp=True, note="Annahme.")
    kv(21, "Stunden je Tagdienst", "=(TD_Ende-TD_Beginn-TD_Pause)*24", "StdTag", "0.00")
    kv(22, "Bereitschaft / Nachtdienst Beginn", time(15, 30), "BD_Beginn", "hh:mm", inp=True)
    kv(23, "Bereitschaft / Nachtdienst Ende (Folgetag)", time(6, 45), "BD_Ende", "hh:mm", inp=True)
    kv(24, "Dauer Bereitschaft (Std.)", "=(1-BD_Beginn+BD_Ende)*24", "StdBereitschaft", "0.00")
    kv(25, "Anrechnung Bereitschaft auf Arbeitszeit", 0.6, "AnrBereitschaft", "0%", inp=True,
       note="Annahme 60 % (≈ 9,15 Std. je Nacht): 7 Nächte je 5-Wochen-Zyklus gleichen die entfallenen Tagdienste "
            "und die 3 Freistellungstage genau aus (Zeitkonto ausgeglichen).")
    kv(26, "Angerechnete Std. je Nachtdienst-/Bereitschaftstag", "=StdBereitschaft*AnrBereitschaft",
       "StdNacht", "0.00")
    kv(27, "Sollstunden je Arbeitstag", 8, "SollStdTag", "0.00", inp=True,
       note="Annahme 40-Std.-Woche (entspricht 06:45–15:30 abzgl. 45 min Pause).")

    section(29, "Sollbesetzung, Personalbedarf und Ampel")
    kv(30, "Sollbesetzung Nachtteam: Meister", SOLL["Meister"], "SollM", inp=True)
    kv(31, "Sollbesetzung Nachtteam: AV", SOLL["AV"], "SollA", inp=True)
    kv(32, "Sollbesetzung Nachtteam: TVT-Fahrer", SOLL["TVT-Fahrer"], "SollT", inp=True)
    kv(33, "Sollbesetzung Nachtteam: Handwerker", SOLL["Handwerker"], "SollH", inp=True)
    kv(34, "Sollbesetzung Nachtteam gesamt", "=SollM+SollA+SollT+SollH", "SollNacht")
    kv(35, "Mindestbesetzung Tagdienst (Personen)", 12, "MinTag", inp=True,
       note="Annahme. Unterschreitung = Ampel Rot (Unterbesetzung).")
    kv(36, "Gewünschte Ausfallreserve Tagdienst (Personen)", 2, "PufferTag", inp=True,
       note="Soll erreicht, aber Reserve kleiner als dieser Wert = Ampel Gelb (kritisch).")

    section(38, "Qualifikationen")
    kv(39, "Grün ab Restlaufzeit über (Tage)", 180, "QualGelbTage", inp=True,
       note="90–180 Tage = Gelb, unter 90 Tage = Rot, abgelaufen = Dunkelrot.")
    kv(40, "Rot / Vorwarnung unter (Tage)", 90, "QualWarnTage", inp=True,
       note="Markierung 90 Tage vor Ablauf.")
    kv(41, "Fehlender Nachweis gilt als ungültig", "NEIN", "StrengeQual", inp=True,
       note="NEIN: Pflichtqualifikation ohne Eintrag blockiert nicht. Auf JA stellen, sobald alle "
            "Schulungsdaten erfasst sind.")
    book.dv_list(ws, "C41", "=JaNein")

    section(43, "Technik")
    kv(44, "Plätze im Planungsraster (Mitarbeiter)", NMAX, "NMax",
       note="Feste Größe des Rasters (Blatt PEP 2027, Hilfsberechnung).")

    # ---------------------------------------------------------------- Listen rechts
    def tbl_title(r, c, text):
        put(ws, f"{c}{r}", text, size=11, bold=True)

    tbl_title(4, "F", "Teams")
    header_row(ws, 5, 6, ["Team", "Kürzel", "Farbe (Hex)", "Nr"], inputs=(0, 1), height=18)
    for i, (t, k, col) in enumerate(TEAMS):
        r = 6 + i
        put(ws, f"F{r}", t, bold=True, bg=col, border=THIN)
        put(ws, f"G{r}", k, h="center", border=THIN)
        put(ws, f"H{r}", col, h="center", border=THIN)
        put(ws, f"I{r}", i + 1, h="center", border=THIN)
    book.table(ws, "tbl_Teams", "F5:I10")
    book.name_ws("Teams", ws, "F6:F10")
    book.name_ws("TeamKurz", ws, "G6:G10")

    tbl_title(12, "F", "Rotationsfolge (fest – nicht ändern)")
    header_row(ws, 13, 6, ["Block", "Blocktyp", "Team"], height=18)
    for i, (typ, t) in enumerate(ROTATION):
        r = 14 + i
        put(ws, f"F{r}", i + 1, h="center", border=THIN)
        put(ws, f"G{r}", typ, h="center", border=THIN)
        put(ws, f"H{r}", t, bold=True, h="center", border=THIN)
    team_cf(book, ws, "H14:H23", "H14")
    book.table(ws, "tbl_Rotation", "F13:H23")
    book.name_ws("RotTeams", ws, "H14:H23")
    book.dv_list(ws, "H14:H23", "=Teams")

    tbl_title(25, "F", "Teamfunktionen")
    header_row(ws, 26, 6, ["Funktion", "Kürzel", "Soll Nachtteam"], height=18)
    soll_name = {"Meister": "=SollM", "AV": "=SollA", "TVT-Fahrer": "=SollT", "Handwerker": "=SollH",
                 "Reserve": "–"}
    for i, f in enumerate(TEAMFUNKTIONEN):
        r = 27 + i
        put(ws, f"F{r}", f, border=THIN)
        put(ws, f"G{r}", FKT_KURZ[f], h="center", border=THIN)
        put(ws, f"H{r}", soll_name[f], h="center", border=THIN)
    book.table(ws, "tbl_Funktionen", "F26:H31")
    book.name_ws("Teamfunktionen", ws, "F27:F31")
    book.name_ws("Funktionen4", ws, "F27:F30")
    book.name_ws("FktKurz", ws, "G27:G31")

    tbl_title(33, "F", "Abwesenheitsarten")
    header_row(ws, 34, 6, ["Art", "Code", "Hinweis"], height=18)
    hints = {"Unterricht": "Thema wählen: FDL → FDL, Kabel → KAB, sonst UR",
             "JAZ": "wird vom JAZ-Konto abgebucht", "Krankheit": "hat Vorrang vor allen Arten"}
    for i, (a, code) in enumerate(ABW_ARTEN):
        r = 35 + i
        put(ws, f"F{r}", a, border=THIN)
        put(ws, f"G{r}", code, h="center", border=THIN, bold=True)
        put(ws, f"H{r}", hints.get(a, ""), size=8, border=THIN)
    book.table(ws, "tbl_AbwArten", "F34:H41")
    book.name_ws("AbwArten", ws, "F35:F41")
    book.name_ws("AbwArtCode", ws, "G35:G41")

    tbl_title(43, "F", "Status-Codes Abwesenheit (Priorität aufsteigend)")
    header_row(ws, 44, 6, ["Prio", "Code", "Bedeutung", "Farbe", "zählt als"], height=18)
    for i, (code, text, bg, fc, b) in enumerate(ABW_CODES):
        r = 45 + i
        put(ws, f"F{r}", i + 1, h="center", border=THIN)
        put(ws, f"G{r}", code, h="center", bold=True, bg=bg, color=fc, border=THIN)
        put(ws, f"H{r}", text, size=8, border=THIN)
        put(ws, f"I{r}", bg, size=8, h="center", border=THIN)
        put(ws, f"J{r}", "abwesend (Soll gutgeschrieben)", size=8, border=THIN)
    book.table(ws, "tbl_Codes", "F44:J53")
    book.name_ws("AbwCodes", ws, "G45:G53")
    book.name_ws("AbwCodeText", ws, "H45:H53")

    put(ws, "F55", "Dienst-Codes (automatisch)", size=11, bold=True)
    for i, (code, text, bg, fc, b) in enumerate(DUTY_CODES):
        r = 56 + i
        put(ws, f"F{r}", text, size=8, border=THIN)
        put(ws, f"G{r}", code, h="center", bold=True, bg=bg, color=fc, border=THIN)

    tbl_title(62, "F", "Feiertage Hessen (automatisch aus Planungsjahr)")
    header_row(ws, 63, 6, ["Feiertag", "Datum", "Wochentag", "Hinweis", "Kurzname Kalender"], inputs=(4,), height=18)
    kurz = ["Neujahr", "Karfreitag", "Ostern", "Oster-montag", "Tag der Arbeit", "Christi Himmelfahrt", "Pfingsten",
            "Pfingst-montag", "Fron-leichnam", "Tag der Dt. Einheit", "1. Weih-nachtstag", "2. Weih-nachtstag"]
    ft = [("Neujahr", "=DATE(Jahr,1,1)"), ("Karfreitag", "=Ostersonntag-2"),
          ("Ostersonntag", "=Ostersonntag"), ("Ostermontag", "=Ostersonntag+1"),
          ("Tag der Arbeit", "=DATE(Jahr,5,1)"), ("Christi Himmelfahrt", "=Ostersonntag+39"),
          ("Pfingstsonntag", "=Ostersonntag+49"), ("Pfingstmontag", "=Ostersonntag+50"),
          ("Fronleichnam", "=Ostersonntag+60"), ("Tag der Deutschen Einheit", "=DATE(Jahr,10,3)"),
          ("1. Weihnachtstag", "=DATE(Jahr,12,25)"), ("2. Weihnachtstag", "=DATE(Jahr,12,26)")]
    for i, (n, f) in enumerate(ft):
        r = 64 + i
        put(ws, f"F{r}", n, border=THIN, color="CC0000", bold=True)
        put(ws, f"G{r}", f, fmt=DATE_FMT, h="center", border=THIN)
        put(ws, f"H{r}", "=" + weekday_formula(f"G{r}"), h="center", border=THIN)
        put(ws, f"I{r}", "gesetzlich (HE)", size=8, border=THIN)
        put(ws, f"J{r}", kurz[i], size=8, border=THIN)
    book.table(ws, "tbl_Feiertage", "F63:J75")
    book.name_ws("FT_Kurz", ws, "J64:J75")
    book.name_ws("FT_Name", ws, "F64:F75")
    book.name_ws("FT_Datum", ws, "G64:G75")

    tbl_title(77, "F", "Sondertage (kein gesetzlicher Feiertag)")
    header_row(ws, 78, 6, ["Sondertag", "Datum", "Wochentag"], height=18)
    for i, (n, f) in enumerate([("Heiligabend", "=DATE(Jahr,12,24)"), ("Silvester", "=DATE(Jahr,12,31)")]):
        r = 79 + i
        put(ws, f"F{r}", n, border=THIN)
        put(ws, f"G{r}", f, fmt=DATE_FMT, h="center", border=THIN)
        put(ws, f"H{r}", "=" + weekday_formula(f"G{r}"), h="center", border=THIN)
    book.table(ws, "tbl_Sondertage", "F78:H80")
    book.name_ws("SO_Name", ws, "F79:F80")
    book.name_ws("SO_Datum", ws, "G79:G80")

    tbl_title(82, "F", "Schulferien Hessen (Eingabe, je Jahr pflegen)")
    header_row(ws, 83, 6, ["Bezeichnung", "Von", "Bis"], inputs=(0, 1, 2), height=18)
    for i in range(10):
        r = 84 + i
        if i < len(SCHULFERIEN):
            n, a, b = SCHULFERIEN[i]
            vals = (n, date(*a), date(*b))
        else:
            vals = (None, None, None)
        for j, v in enumerate(vals):
            put(ws, f"{'FGH'[j]}{r}", v, border=THIN, fmt=DATE_FMT if j else None, h="center" if j else None)
    book.table(ws, "tbl_Schulferien", "F83:H93")
    book.name_ws("SF_Name", ws, "F84:F93")
    book.name_ws("SF_Von", ws, "G84:G93")
    book.name_ws("SF_Bis", ws, "H84:H93")
    put(ws, "F94", "Quelle 2027: graue Markierung im Bereitschaftskalender 2027 (kalenderpedia.de).",
        size=8, color=MUTED, italic=True)

    # Unterrichte
    tbl_title(4, "L", "Unterrichte / Qualifikationen")
    header_row(ws, 5, 12, ["Unterricht", "Gültigkeit (Monate)"], inputs=(1,), height=30)
    for i, (u, m) in enumerate(UNTERRICHTE):
        r = 6 + i
        put(ws, f"L{r}", u, border=THIN)
        put(ws, f"M{r}", m, h="center", border=THIN)
    book.table(ws, "tbl_Unterrichte", "L5:M18")
    book.name_ws("Unterrichte", ws, "L6:L18")
    book.name_ws("U_Monate", ws, "M6:M18")
    put(ws, "N6", "Gültigkeitsdauer = Annahme, bitte prüfen.", size=8, color=MUTED, italic=True)

    tbl_title(20, "L", "Pflichtqualifikationen je Funktion (JA = für Einsatzfähigkeit erforderlich; Annahme)")
    put(ws, "L22", "Funktion", bold=True, color="FFFFFF", bg=HEAD_DARK, border=THIN)
    for j in range(13):
        c = ws.cell(22, 13 + j, f"=INDEX(Unterrichte,{j + 1})")
        style(c, size=8, bold=True, color="FFFFFF", bg=HEAD_DARK, h="center", wrap=True, border=THIN)
    ws.row_dimensions[22].height = 36
    for i, f in enumerate(FUNKTIONEN):
        r = 23 + i
        put(ws, f"L{r}", f"=INDEX(Funktionen4,{i + 1})", bold=True, border=THIN)
        for j, u in enumerate(U_NAMES):
            put(ws, f"{L(13 + j)}{r}", "JA" if u in PFLICHT[f] else "NEIN", h="center",
                border=THIN, size=9)
    book.dv_list(ws, "M23:Y26", "=JaNein")
    book.cf(ws, "M23:Y26", 'M23="JA"', bg="C6EFCE", color="006100", bold=True)
    book.cf(ws, "M23:Y26", 'M23="NEIN"', color="A6A6A6")
    for i, n in enumerate(["Req_Meister", "Req_AV", "Req_TVT", "Req_HW"]):
        book.name_ws(n, ws, f"M{23 + i}:Y{23 + i}")
    book.name_ws("ReqMatrix", ws, "M23:Y26")

    # Kleine Listen
    put(ws, "L29", "Listen für Auswahlfelder", size=11, bold=True)
    put(ws, "L30", "JA", border=THIN)
    put(ws, "L31", "NEIN", border=THIN)
    book.name_ws("JaNein", ws, "L30:L31")
    for i in range(12):
        put(ws, f"N{30 + i}", i + 1, h="center", border=THIN)
        put(ws, f"O{30 + i}", MONATE[i], border=THIN)
    book.name_ws("MonNr", ws, "N30:N41")
    book.name_ws("MonName", ws, "O30:O41")
    for i, k in enumerate(["Überstunden", "JAZ", "Urlaub"]):
        put(ws, f"Q{30 + i}", k, border=THIN)
    book.name_ws("Konten", ws, "Q30:Q32")
    for i, k in enumerate(["Vortrag Vorjahr", "Auszahlung", "Korrektur", "Sonstiges"]):
        put(ws, f"S{30 + i}", k, border=THIN)
    book.name_ws("BuchArten", ws, "S30:S33")
    for i in range(4):
        put(ws, f"U{30 + i}", i + 1, h="center", border=THIN)
    book.name_ws("QuNr", ws, "U30:U33")
    put(ws, "N29", "Monate", bold=True)
    put(ws, "Q29", "Konten", bold=True)
    put(ws, "S29", "Buchungsarten", bold=True)
    put(ws, "U29", "Quartale", bold=True)
    ws.column_dimensions["O"].width = 12
    ws.freeze_panes = "A4"
    ws.sheet_properties.tabColor = "7F7F7F"


# ====================================================================== Mitarbeiterstamm
MA_COLS = ["Personalnummer", "Name", "Team", "Eintrittsdatum", "Telefon", "E-Mail", "Aktiv",
           "Jahresurlaub", "Resturlaub", "Überstunden", "JAZ", "Teamfunktion", "Sortierschlüssel",
           "Prüfung"]
MA_FIRST = 6


def build_stamm(book, ws, people, demo=False):
    title(ws, "Blatt 1 – Mitarbeiterstamm",
          "Gelbe Spaltenköpfe = Eingabe, türkis = automatisch. Neue Mitarbeiter in der nächsten freien "
          "Zeile erfassen (Platz für 60) und zusätzlich in der Team- und Funktionsmatrix anlegen.")
    put(ws, "A3", "Format: Personalnummer 00123456 · Eintrittsdatum TT.MM.JJJJ (leer = ganzjährig) · "
                  "Telefon 0151 1234567 · E-Mail vorname.nachname@firma.de · Aktiv JA/NEIN · "
                  "Jahresurlaub in Tagen (Annahme 30). Resturlaub, Überstunden und JAZ (Std.) werden berechnet.",
        size=8, color=MUTED, italic=True)
    hr = MA_FIRST - 1
    header_row(ws, hr, 1, MA_COLS, inputs=(0, 1, 2, 3, 4, 5, 6, 7))
    last = MA_FIRST + CAP_MA - 1
    for i in range(CAP_MA):
        r = MA_FIRST + i
        p = people[i] if i < len(people) else None
        if p:
            ws[f"B{r}"] = p["name"]
            ws[f"C{r}"] = p["team"]
            ws[f"G{r}"] = "JA"
            ws[f"H{r}"] = 30
            if demo:
                ws[f"A{r}"] = f"{10000100 + i * 7:08d}"
        ws[f"I{r}"] = (f'=IF(B{r}="","",H{r}+SUMIFS(ZB_Wert,ZB_MA,B{r},ZB_Konto,"Urlaub")'
                       f'-IFERROR(INDEX(ST_Urlaub,MATCH(B{r},C_Name,0)),0))')
        ws[f"J{r}"] = f'=IF(B{r}="","",IFERROR(INDEX(UE_SaldoUe,MATCH(B{r},C_Name,0)),""))'
        ws[f"K{r}"] = f'=IF(B{r}="","",IFERROR(INDEX(UE_SaldoJAZ,MATCH(B{r},C_Name,0)),""))'
        ws[f"L{r}"] = f'=IF(B{r}="","",IFERROR(INDEX(FM_TF,MATCH(B{r},FM_Name,0))&"",""))'
        ws[f"M{r}"] = (f'=IF(B{r}="","",IF(G{r}="NEIN",9,IFERROR(MATCH(C{r},Teams,0),8))*100000'
                       f'+IFERROR(MATCH(L{r},Teamfunktionen,0),6)*1000+ROW())')
        ws[f"N{r}"] = (f'=IF(B{r}="","",{X}TEXTJOIN("; ",TRUE,IF(COUNTIF($B${MA_FIRST}:$B${last},B{r})>1,'
                       f'"Name doppelt",""),IF(C{r}="","Team fehlt",""),IF(L{r}="","fehlt in Funktionsmatrix",""),'
                       f'IF(G{r}="","Aktiv leer","")))')
        for c in "ABCDEFGHIJKLMN":
            cell = ws[f"{c}{r}"]
            style(cell, size=10, h="center" if c not in "BFN" else None)
        ws[f"D{r}"].number_format = DATE_FMT
        for c in "IJK":
            ws[f"{c}{r}"].number_format = "0.0"
        style(ws[f"M{r}"], size=8, color=MUTED, h="center")
    book.table(ws, "tbl_MA", f"A{hr}:N{last}")
    rng = lambda c: f"{c}{MA_FIRST}:{c}{last}"
    for n, c in [("MA_PNr", "A"), ("MA_Name", "B"), ("MA_Team", "C"), ("MA_Eintritt", "D"),
                 ("MA_Aktiv", "G"), ("MA_Urlaub", "H"), ("MA_TF", "L"), ("MA_Sort", "M")]:
        book.name_ws(n, ws, rng(c))
    book.dv_list(ws, rng("C"), "=Teams")
    book.dv_list(ws, rng("G"), "=JaNein")
    book.dv_date(ws, rng("D"))
    book.dv_number(ws, rng("H"), 0, 60)
    team_cf(book, ws, rng("C"), f"C{MA_FIRST}")
    book.cf(ws, rng("N"), f'N{MA_FIRST}<>""', bg="FCE4D6", color="C00000")
    book.cf(ws, f"A{MA_FIRST}:N{last}", f'$G{MA_FIRST}="NEIN"', color="A6A6A6", italic=True)
    col_widths(ws, {"A": 15, "B": 20, "C": 10, "D": 13, "E": 16, "F": 28, "G": 8, "H": 11, "I": 11,
                    "J": 12, "K": 10, "L": 13, "M": 12, "N": 30})
    ws.freeze_panes = f"C{MA_FIRST}"
    ws.sheet_properties.tabColor = "FFC000"


# ====================================================================== Funktionsmatrix
FM_FIRST = 6


def build_matrix(book, ws, people):
    title(ws, "Blatt 2 – Team- und Funktionsmatrix",
          "Je Mitarbeiter: Teamfunktion (Stammposition im Team) und beliebig viele Funktionen/Qualifikationen "
          "(JA/NEIN). Farben der Unterrichte = Status aus Blatt Qualifikationen.")
    put(ws, "A3", "Teamfunktion = feste Position im Team (wie Bereitschaftskalender). Weitere JA-Funktionen = "
                  "Mehrfachfunktion → Mitarbeiter kann vertreten. Reserve = Springer im Team.",
        size=8, color=MUTED, italic=True)
    labels = ["Name", "Team", "Teamfunktion"] + FUNKTIONEN + U_NAMES + \
             ["Anzahl Funktionen", "Anzahl Qualifikationen", "Mehrfachfunktion", "Prüfung"]
    hr = FM_FIRST - 1
    inputs = tuple([0, 2] + list(range(3, 3 + 4 + 13)))
    header_row(ws, hr, 1, labels, inputs=inputs, height=48)
    # Gruppenkoepfe
    put(ws, "D4", "Funktionen", bold=True, h="center", bg=HEAD, border=THIN)
    ws.merge_cells("D4:G4")
    put(ws, "H4", "Unterrichte / Qualifikationen", bold=True, h="center", bg=HEAD, border=THIN)
    ws.merge_cells("H4:T4")
    last = FM_FIRST + CAP_FM - 1
    helper0 = 26  # Spalte Z: Status-Nr je Unterricht (Hilfsspalten fuer Farben)
    for i in range(CAP_FM):
        r = FM_FIRST + i
        p = people[i] if i < len(people) else None
        if p:
            ws[f"A{r}"] = p["name"]
            ws[f"C{r}"] = p["funktion"]
            extra = set(p.get("zusatz", []))
            fk = {p["funktion"]} | (extra & set(FUNKTIONEN))
            for j, f in enumerate(FUNKTIONEN):
                ws.cell(r, 4 + j, "JA" if f in fk else "NEIN")
            req = set()
            for f in fk:
                req |= set(PFLICHT.get(f, []))
            req |= extra & set(U_NAMES)
            for j, u in enumerate(U_NAMES):
                ws.cell(r, 8 + j, "JA" if u in req else "NEIN")
        ws[f"B{r}"] = f'=IF(A{r}="","",IFERROR(INDEX(MA_Team,MATCH(A{r},MA_Name,0))&"","–"))'
        ws[f"U{r}"] = f'=IF(A{r}="","",COUNTIF(D{r}:G{r},"JA"))'
        ws[f"V{r}"] = f'=IF(A{r}="","",COUNTIF(H{r}:T{r},"JA"))'
        ws[f"W{r}"] = f'=IF(A{r}="","",IF(U{r}>1,"JA","NEIN"))'
        ws[f"X{r}"] = (f'=IF(A{r}="","",{X}TEXTJOIN("; ",TRUE,IF(COUNTIF(MA_Name,A{r})=0,"nicht im Mitarbeiterstamm",""),'
                       f'IF(COUNTIF($A${FM_FIRST}:$A${last},A{r})>1,"doppelt",""),IF(C{r}="","Teamfunktion fehlt",""),'
                       f'IF(AND(C{r}<>"",C{r}<>"Reserve"),IF(IFERROR(INDEX(D{r}:G{r},MATCH(C{r},$D${hr}:$G${hr},0)),"")<>"JA",'
                       f'"Teamfunktion nicht als JA markiert",""),"")))')
        for j in range(13):
            qc = L(8 + j)
            hc = L(helper0 + j)
            ws[f"{hc}{r}"] = (f'=IF(OR($A{r}="",{qc}{r}<>"JA"),"",IF(COUNTIFS(Q_MA,$A{r},Q_Qual,{qc}${hr})=0,5,'
                              f'{X}MINIFS(Q_StatusNr,Q_MA,$A{r},Q_Qual,{qc}${hr})))')
            style(ws[f"{hc}{r}"], size=8, color=MUTED, h="center")
        for c in range(1, 25):
            style(ws.cell(r, c), size=9, h="center" if c not in (1, 24) else None)
    book.table(ws, "tbl_FM", f"A{hr}:X{last}")
    put(ws, f"{L(helper0)}{hr}", "Status-Nr je Unterricht (Hilfsspalten für die Farben)", size=8, color=MUTED)
    for j in range(13):
        ws.column_dimensions[L(helper0 + j)].hidden = True
    rng = lambda c: f"{c}{FM_FIRST}:{c}{last}"
    book.name_ws("FM_Name", ws, rng("A"))
    book.name_ws("FM_TF", ws, rng("C"))
    for j, n in enumerate(["FM_M", "FM_AV", "FM_TVT", "FM_HW"]):
        book.name_ws(n, ws, rng(L(4 + j)))
    book.dv_list(ws, rng("A"), "=MA_Liste", error=False)
    book.dv_list(ws, rng("C"), "=Teamfunktionen")
    book.dv_list(ws, f"D{FM_FIRST}:T{last}", "=JaNein")
    fr = FM_FIRST
    # Funktionen: Stammfunktion kraeftig, weitere JA hell
    book.cf(ws, f"D{fr}:G{last}", f'AND(D{fr}="JA",D${hr}=$C{fr})', bg="31869B", color="FFFFFF", bold=True)
    book.cf(ws, f"D{fr}:G{last}", f'D{fr}="JA"', bg="B7DEE8", bold=True)
    hz = L(helper0)
    for nr, bg, fc in [(1, "92D050", "000000"), (2, "FFFF00", "000000"), (3, "FF0000", "FFFFFF"),
                       (4, "C00000", "FFFFFF"), (5, "D9D9D9", "000000")]:
        book.cf(ws, f"H{fr}:T{last}", f'AND(H{fr}="JA",{hz}{fr}={nr})', bg=bg, color=fc, bold=True)
    book.cf(ws, f"D{fr}:T{last}", f'D{fr}="NEIN"', color="BFBFBF")
    book.cf(ws, rng("X"), f'X{fr}<>""', bg="FCE4D6", color="C00000")
    book.cf(ws, rng("W"), f'W{fr}="JA"', bg="DAEEF3", bold=True)
    col_widths(ws, {"A": 18, "B": 9, "C": 13, "U": 10, "V": 12, "W": 11, "X": 34})
    for c in range(4, 21):
        ws.column_dimensions[L(c)].width = 8.5
    put(ws, "Z3", "", size=8)
    # Legende
    put(ws, "J3", "Legende Unterrichte:", size=8, bold=True)
    for k, (t, bg, fc) in enumerate([("gültig > 180 T.", "92D050", "000000"), ("90–180 T.", "FFFF00", "000000"),
                                     ("< 90 T.", "FF0000", "FFFFFF"), ("abgelaufen", "C00000", "FFFFFF"),
                                     ("Nachweis fehlt", "D9D9D9", "000000")]):
        put(ws, f"{L(12 + 2 * k)}3", t, size=8, bg=bg, color=fc, h="center")
        ws.merge_cells(f"{L(12 + 2 * k)}3:{L(13 + 2 * k)}3")
    ws.freeze_panes = f"D{FM_FIRST}"
    ws.sheet_properties.tabColor = "FFC000"


# ====================================================================== Qualifikationen
Q_FIRST = 6


def build_quali(book, ws, people, demo_dates=None):
    title(ws, "Blatt 3 – Qualifikationen und Unterrichte",
          "Je Zeile ein Nachweis. Ablaufdatum = letzte Schulung + Gültigkeit (Einstellungen) oder manuelles "
          "Ablaufdatum. Unterrichte aus dem Blatt Abwesenheiten (Art Unterricht + Thema) zählen automatisch.")
    put(ws, "A3", "Farben: Grün > 180 Tage · Gelb 90–180 Tage · Rot < 90 Tage (Vorwarnung) · Dunkelrot abgelaufen · "
                  "Grau Nachweis fehlt. Restlaufzeit bezogen auf den Stichtag (Einstellungen).",
        size=8, color=MUTED, italic=True)
    labels = ["Mitarbeiter", "Qualifikation", "Letzte Schulung", "Ablaufdatum", "Restlaufzeit (Tage)",
              "Status", "Team", "Gültigkeit (Monate)", "Ablauf manuell", "Letzter Unterricht lt. Plan",
              "Nächster Unterricht geplant", "Status-Nr", "Bemerkung", "Prüfung", "Gültig für Planung bis"]
    hr = Q_FIRST - 1
    header_row(ws, hr, 1, labels, inputs=(0, 1, 2, 8, 12))
    last = Q_FIRST + CAP_Q - 1
    rows = []
    for p in people:
        fm_req = set()
        fk = {p["funktion"]} | (set(p.get("zusatz", [])) & set(FUNKTIONEN))
        for f in fk:
            fm_req |= set(PFLICHT.get(f, []))
        fm_req |= set(p.get("zusatz", [])) & set(U_NAMES)
        for u in U_NAMES:
            if u in fm_req:
                rows.append((p["name"], u))
    for i in range(CAP_Q):
        r = Q_FIRST + i
        if i < len(rows):
            ws[f"A{r}"], ws[f"B{r}"] = rows[i]
            if demo_dates and rows[i] in demo_dates:
                ws[f"C{r}"] = demo_dates[rows[i]]
            elif not demo_dates:
                ws[f"M{r}"] = "Datum der letzten Schulung eintragen"
        ws[f"J{r}"] = (f'=IF(OR(A{r}="",B{r}=""),"",IF({X}MAXIFS(Abw_BisW,Abw_MA,A{r},Abw_Art,"Unterricht",'
                       f'Abw_Thema,B{r},Abw_BisW,"<="&Stichtag)=0,"",{X}MAXIFS(Abw_BisW,Abw_MA,A{r},'
                       f'Abw_Art,"Unterricht",Abw_Thema,B{r},Abw_BisW,"<="&Stichtag)))')
        ws[f"H{r}"] = f'=IF(B{r}="","",IFERROR(INDEX(U_Monate,MATCH(B{r},Unterrichte,0)),12))'
        ws[f"D{r}"] = (f'=IF(OR(A{r}="",B{r}=""),"",IF(I{r}<>"",I{r},IF(MAX(N(C{r}),N(J{r}))=0,"",'
                       f'EDATE(MAX(N(C{r}),N(J{r})),H{r}))))')
        ws[f"E{r}"] = f'=IF(D{r}="","",D{r}-Stichtag)'
        ws[f"L{r}"] = (f'=IF(OR(A{r}="",B{r}=""),"",IF(D{r}="",5,IF(E{r}<0,4,IF(E{r}<QualWarnTage,3,'
                       f'IF(E{r}<=QualGelbTage,2,1)))))')
        ws[f"F{r}"] = (f'=IF(L{r}="","",CHOOSE(L{r},"gültig","gültig – Vorwarnung (≤ "&QualGelbTage&" Tage)",'
                       f'"⚠ läuft ab (< "&QualWarnTage&" Tage)","abgelaufen","Nachweis fehlt")&IF(AND(K{r}<>"",L{r}>=2),'
                       f'" – Unterricht geplant",""))')
        ws[f"O{r}"] = (f'=IF(D{r}="","",IF(K{r}="",D{r},IF(K{r}<=D{r},EDATE(K{r},H{r}),D{r})))')
        ws[f"G{r}"] = f'=IF(A{r}="","",IFERROR(INDEX(MA_Team,MATCH(A{r},MA_Name,0))&"","–"))'
        ws[f"K{r}"] = (f'=IF(OR(A{r}="",B{r}=""),"",IF({X}MINIFS(Abw_Von,Abw_MA,A{r},Abw_Art,"Unterricht",'
                       f'Abw_Thema,B{r},Abw_Von,">"&Stichtag)=0,"",{X}MINIFS(Abw_Von,Abw_MA,A{r},'
                       f'Abw_Art,"Unterricht",Abw_Thema,B{r},Abw_Von,">"&Stichtag)))')
        ws[f"N{r}"] = (f'=IF(A{r}="","",{X}TEXTJOIN("; ",TRUE,IF(COUNTIF(MA_Name,A{r})=0,"Name unbekannt",""),'
                       f'IF(B{r}="","Qualifikation fehlt",""),IF(COUNTIFS($A${Q_FIRST}:$A${last},A{r},'
                       f'$B${Q_FIRST}:$B${last},B{r})>1,"mehrere Einträge (neuester zählt)","")))')
        for c in "ABCDEFGHIJKLMNO":
            style(ws[f"{c}{r}"], size=9, h="center" if c not in "ABFMN" else None)
        for c in "CDIJKO":
            ws[f"{c}{r}"].number_format = DATE_FMT
        ws[f"E{r}"].number_format = "0"
    book.table(ws, "tbl_Q", f"A{hr}:O{last}")
    rng = lambda c: f"{c}{Q_FIRST}:{c}{last}"
    for n, c in [("Q_MA", "A"), ("Q_Qual", "B"), ("Q_Ablauf", "D"), ("Q_Rest", "E"), ("Q_StatusNr", "L"),
                 ("Q_Team", "G"), ("Q_PlanBis", "O")]:
        book.name_ws(n, ws, rng(c))
    book.dv_list(ws, rng("A"), "=MA_Liste", error=False)
    book.dv_list(ws, rng("B"), "=Unterrichte")
    book.dv_date(ws, f"C{Q_FIRST}:C{last} I{Q_FIRST}:I{last}")
    fr = Q_FIRST
    for nr, bg, fc in [(1, "92D050", "000000"), (2, "FFFF00", "000000"), (3, "FF0000", "FFFFFF"),
                       (4, "C00000", "FFFFFF"), (5, "D9D9D9", "000000")]:
        book.cf(ws, f"D{fr}:F{last}", f'$L{fr}={nr}', bg=bg, color=fc, bold=nr in (3, 4))
    book.cf(ws, rng("N"), f'N{fr}<>""', bg="FCE4D6", color="C00000")
    col_widths(ws, {"A": 18, "B": 17, "C": 13, "D": 13, "E": 11, "F": 38, "G": 9, "H": 10, "I": 13,
                    "J": 13, "K": 13, "L": 8, "M": 34, "N": 30, "O": 13})
    put(ws, "O4", "verlängert um rechtzeitig geplanten Unterricht", size=7, color=MUTED, italic=True)
    ws.freeze_panes = f"C{Q_FIRST}"
    ws.sheet_properties.tabColor = "FFC000"


# ====================================================================== Abwesenheiten
ABW_FIRST = 6


def build_abwesenheit(book, ws, demo_rows=None):
    title(ws, "Blatt 4 – Urlaub und Abwesenheiten",
          "Eine Zeile je Abwesenheit (ganze Kalendertage). Fließt automatisch in PEP 2027, Schichtplan, "
          "Einsatzfähigkeit, Zeitkonten, Statistiken und Dashboard ein.")
    put(ws, "A3", "Format: Von/Bis TT.MM.JJJJ (Bis leer = 1 Tag) · Art aus Liste · bei Unterricht das Thema wählen "
                  "(FDL-/Kabeleinweisung bekommen eigene Farben, Datum zählt als letzte Schulung). "
                  "Bei Überschneidung gilt die höhere Priorität (Krankheit vor Urlaub …).",
        size=8, color=MUTED, italic=True)
    labels = ["Mitarbeiter", "Art", "Von", "Bis", "Unterricht / Thema", "Bemerkung", "Team",
              "Bis (wirksam)", "Kalendertage", "Arbeitstage", "Code", "Prio", "Prüfung"]
    hr = ABW_FIRST - 1
    header_row(ws, hr, 1, labels, inputs=(0, 1, 2, 3, 4, 5))
    last = ABW_FIRST + CAP_ABW - 1
    demo_rows = demo_rows or []
    for i in range(CAP_ABW):
        r = ABW_FIRST + i
        if i < len(demo_rows):
            n, art, von, bis, thema, bem = demo_rows[i]
            ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = n, art, von
            ws[f"D{r}"], ws[f"E{r}"], ws[f"F{r}"] = bis, thema, bem
        ws[f"G{r}"] = f'=IF(A{r}="","",IFERROR(INDEX(MA_Team,MATCH(A{r},MA_Name,0))&"","–"))'
        ws[f"H{r}"] = f'=IF(OR(A{r}="",C{r}=""),"",IF(D{r}="",C{r},D{r}))'
        ws[f"I{r}"] = f'=IF(H{r}="","",H{r}-C{r}+1)'
        ws[f"J{r}"] = (f'=IF(H{r}="","",IF(MIN(H{r},Jahresende)<MAX(C{r},Jahresbeginn),0,'
                       f'NETWORKDAYS(MAX(C{r},Jahresbeginn),MIN(H{r},Jahresende),FT_Datum)))')
        ws[f"K{r}"] = (f'=IF(OR(A{r}="",B{r}="",C{r}=""),"",IF(B{r}="Unterricht",IF(E{r}="FDL-Einweisung","FDL",'
                       f'IF(E{r}="Kabeleinweisung","KAB","UR")),IFERROR(INDEX(AbwArtCode,MATCH(B{r},AbwArten,0)),"")))')
        ws[f"L{r}"] = f'=IF(K{r}="","",MATCH(K{r},AbwCodes,0))'
        ws[f"M{r}"] = (f'=IF(A{r}="","",{X}TEXTJOIN("; ",TRUE,IF(COUNTIF(MA_Name,A{r})=0,"Name unbekannt",""),'
                       f'IF(B{r}="","Art fehlt",""),IF(C{r}="","Von fehlt",""),'
                       f'IF(AND(D{r}<>"",C{r}<>""),IF(D{r}<C{r},"Bis liegt vor Von",""),""),'
                       f'IF(H{r}<>"",IF(COUNTIFS($A${ABW_FIRST}:$A${last},A{r},$C${ABW_FIRST}:$C${last},"<="&H{r},'
                       f'$H${ABW_FIRST}:$H${last},">="&C{r})>1,"überschneidet sich mit anderem Eintrag",""),""),'
                       f'IF(AND(B{r}="Unterricht",E{r}=""),"Thema fehlt (optional)","")))')
        for c in "ABCDEFGHIJKLM":
            style(ws[f"{c}{r}"], size=9, h="center" if c not in "AFM" else None)
        for c in "CDH":
            ws[f"{c}{r}"].number_format = DATE_FMT
    book.table(ws, "tbl_Abw", f"A{hr}:M{last}")
    rng = lambda c: f"{c}{ABW_FIRST}:{c}{last}"
    for n, c in [("Abw_MA", "A"), ("Abw_Art", "B"), ("Abw_Von", "C"), ("Abw_Thema", "E"),
                 ("Abw_BisW", "H"), ("Abw_AT", "J"), ("Abw_Code", "K"), ("Abw_Prio", "L")]:
        book.name_ws(n, ws, rng(c))
    book.dv_list(ws, rng("A"), "=MA_Liste", error=False)
    book.dv_list(ws, rng("B"), "=AbwArten")
    book.dv_list(ws, rng("E"), "=Unterrichte")
    book.dv_date(ws, f"C{ABW_FIRST}:D{last}")
    fr = ABW_FIRST
    for code, _t, bg, fc, b in ABW_CODES:
        book.cf(ws, rng("K"), f'K{fr}="{code}"', bg=bg, color=fc, bold=True)
    book.cf(ws, rng("M"), f'M{fr}<>""', bg="FCE4D6", color="C00000")
    col_widths(ws, {"A": 18, "B": 14, "C": 12, "D": 12, "E": 17, "F": 28, "G": 9, "H": 12, "I": 11,
                    "J": 11, "K": 7, "L": 6, "M": 40})
    ws.freeze_panes = f"B{ABW_FIRST}"
    ws.sheet_properties.tabColor = "FFC000"


# ====================================================================== Vertretungen (inkl. Ersatzsuche)
VERT_HR = 18


def build_vertretung(book, ws, demo_rows=None):
    title(ws, "Vertretungsmanagement – Ersatzsuche und bestätigte Vertretungen",
          "Oben: Ersatzvorschläge für Datum/Team/Funktion. Unten: bestätigte Vertretungen – sie zählen in "
          "PEP 2027 (Code V), Einsatzfähigkeit, Statistik und Zeitkonten.")
    put(ws, "A3", "Ersatzsuche", size=12, bold=True)
    put(ws, "A4", "Datum", bold=True, border=THIN)
    put(ws, "B4", "=Stichtag", bg=INPUT_BG, fmt=DATE_FMT, h="center", border=THIN, bold=True)
    put(ws, "A5", "Team (zu besetzen)", bold=True, border=THIN)
    put(ws, "B5", "=IFERROR(INDEX(Plan_Nacht,B4-Jahresbeginn+1),\"\")", bg=INPUT_BG, h="center", border=THIN, bold=True)
    put(ws, "A6", "Funktion", bold=True, border=THIN)
    put(ws, "B6", "Meister", bg=INPUT_BG, h="center", border=THIN, bold=True)
    book.dv_date(ws, "B4")
    book.dv_list(ws, "B5", "=Teams")
    book.dv_list(ws, "B6", "=Funktionen4")
    put(ws, "C4", "=IF(B4=\"\",\"\",\"Nachtteam: \"&IFERROR(INDEX(Plan_Nacht,B4-Jahresbeginn+1),\"–\")&"
                  "\"  ·  freigestellt: \"&IFERROR(INDEX(Plan_Frei,B4-Jahresbeginn+1)&\"\",\"–\"))", size=9)
    put(ws, "C5", "Priorität: 1 = gleiches Team (Reserve / Überdeckung) · 2 = anderes Team, gleiche Funktion · "
                  "3 = anderes Team mit Mehrfachfunktion. Innerhalb gleicher Priorität: wenigste Nachtdienste.",
        size=8, color=MUTED, italic=True)
    put(ws, "C6", "Vorausgesetzt: anwesend, Funktion JA, Pflichtqualifikationen am Datum gültig, nicht im "
                  "freigestellten Team.", size=8, color=MUTED, italic=True)
    # Hilfszellen
    put(ws, "H4", "=IF(B4=\"\",1,B4-Jahresbeginn+1)", size=8, color=MUTED)   # Tag-Nr
    put(ws, "H5", "=IFERROR(INDEX(Plan_Frei,H4)&\"\",\"\")", size=8, color=MUTED)
    put(ws, "H6", "=IFERROR(MATCH(B6,Funktionen4,0),1)", size=8, color=MUTED)
    put(ws, "G4", "Hilfswerte:", size=8, color=MUTED)
    header_row(ws, 8, 1, ["Rang", "Mitarbeiter", "Team", "Teamfunktion", "Priorität", "Nachtdienste (Jahr)"],
               height=18)
    q = "CHOOSE($H$6,C_QM,C_QA,C_QT,C_QH)"
    g = (f'((C_Name<>"")*(INDEX(C_Grid,0,$H$4)=0)*({q}>=$B$4)*('
         f'(C_Team=$B$5)*(C_Fkt="Reserve")'
         f'+(C_Team<>$B$5)*(C_Team<>$H$5)*((C_Fkt=$B$6)*2+(C_Fkt<>$B$6)*(C_Fkt<>"Reserve")*3+(C_Fkt="Reserve")*3)))')
    for k in range(1, 6):
        r = 8 + k
        put(ws, f"A{r}", k, h="center", border=THIN)
        put(ws, f"J{r}", f'=IF(OR($B$4="",$B$5=""),1E+9,SUMPRODUCT(SMALL({g}*1E+6+C_N*1000+C_Idx+({g}=0)*1E+9,{k})))',
            size=8, color=MUTED)
        put(ws, f"B{r}", f'=IF(J{r}>=1E+9,IF({k}=1,"kein Ersatz verfügbar",""),INDEX(C_Name,MOD(J{r},1000)))',
            border=THIN, bold=True)
        put(ws, f"C{r}", f'=IF(J{r}>=1E+9,"",INDEX(C_Team,MOD(J{r},1000)))', border=THIN, h="center")
        put(ws, f"D{r}", f'=IF(J{r}>=1E+9,"",INDEX(C_Fkt,MOD(J{r},1000)))', border=THIN, h="center")
        put(ws, f"E{r}", f'=IF(J{r}>=1E+9,"",CHOOSE(INT(J{r}/1E+6),"1 – gleiches Team","2 – anderes Team",'
                         f'"3 – Mehrfachfunktion"))', border=THIN)
        put(ws, f"F{r}", f'=IF(J{r}>=1E+9,"",INDEX(C_N,MOD(J{r},1000)))', border=THIN, h="center")
    team_cf(book, ws, "C9:C13", "C9")
    team_cf(book, ws, "B5", "B5")
    put(ws, "J8", "Hilfswert", size=8, color=MUTED)

    put(ws, f"A{VERT_HR - 2}", "Bestätigte Vertretungen", size=12, bold=True)
    put(ws, f"A{VERT_HR - 1}", "Format: Von/Bis TT.MM.JJJJ · Team = vertretenes Team (meist Nachtteam) · Funktion, die "
                               "übernommen wird · Vertreter aus der Liste.", size=8, color=MUTED, italic=True)
    labels = ["Von", "Bis", "Team (vertreten)", "Funktion", "Vertreter", "Vertretener Mitarbeiter", "Bemerkung",
              "Bis (wirksam)", "Tage", "Team Vertreter", "Prüfung"]
    header_row(ws, VERT_HR, 1, labels, inputs=(0, 1, 2, 3, 4, 5, 6))
    first = VERT_HR + 1
    last = VERT_HR + CAP_VERT
    demo_rows = demo_rows or []
    for i in range(CAP_VERT):
        r = first + i
        if i < len(demo_rows):
            for j, v in enumerate(demo_rows[i]):
                ws.cell(r, 1 + j, v)
        ws[f"H{r}"] = f'=IF(OR(A{r}="",E{r}=""),"",IF(B{r}="",A{r},B{r}))'
        ws[f"I{r}"] = f'=IF(H{r}="","",H{r}-A{r}+1)'
        ws[f"J{r}"] = f'=IF(E{r}="","",IFERROR(INDEX(MA_Team,MATCH(E{r},MA_Name,0))&"","–"))'
        fcol = f'INDEX(CHOOSE(MATCH(D{r},Funktionen4,0),FM_M,FM_AV,FM_TVT,FM_HW),MATCH(E{r},FM_Name,0))'
        ws[f"K{r}"] = (f'=IF(E{r}="","",{X}TEXTJOIN("; ",TRUE,IF(COUNTIF(MA_Name,E{r})=0,"Vertreter unbekannt",""),'
                       f'IF(OR(C{r}="",D{r}=""),"Team/Funktion fehlt",""),'
                       f'IF(H{r}="","",IF(COUNTIFS(Abw_MA,E{r},Abw_Von,"<="&H{r},Abw_BisW,">="&A{r})>0,'
                       f'"Vertreter im Zeitraum abwesend","")),'
                       f'IF(D{r}="","",IF(IFERROR({fcol},"")<>"JA","Funktion lt. Matrix nicht vorhanden","")),'
                       f'IF(AND(J{r}<>"",J{r}=C{r}),"Vertreter ist im selben Team","")))')
        for c in "ABCDEFGHIJK":
            style(ws[f"{c}{r}"], size=9, h="center" if c not in "EFGK" else None)
        for c in "ABH":
            ws[f"{c}{r}"].number_format = DATE_FMT
    book.table(ws, "tbl_Vert", f"A{VERT_HR}:K{last}")
    rng = lambda c: f"{c}{first}:{c}{last}"
    for n, c in [("Vert_Von", "A"), ("Vert_Team", "C"), ("Vert_Fkt", "D"), ("Vert_Vertreter", "E"),
                 ("Vert_BisW", "H")]:
        book.name_ws(n, ws, rng(c))
    book.dv_date(ws, f"A{first}:B{last}")
    book.dv_list(ws, rng("C"), "=Teams")
    book.dv_list(ws, rng("D"), "=Funktionen4")
    book.dv_list(ws, rng("E"), "=MA_Liste", error=False)
    book.dv_list(ws, rng("F"), "=MA_Liste", error=False)
    team_cf(book, ws, rng("C"), f"C{first}")
    book.cf(ws, rng("K"), f'K{first}<>""', bg="FCE4D6", color="C00000")
    col_widths(ws, {"A": 20, "B": 20, "C": 15, "D": 14, "E": 20, "F": 20, "G": 26, "H": 12, "I": 7,
                    "J": 12, "K": 40})
    ws.sheet_properties.tabColor = "FFC000"


# ====================================================================== Zeitbuchungen
ZB_FIRST = 6


def build_zeitbuchungen(book, ws, demo_rows=None):
    title(ws, "Zeitbuchungen – Vorträge, Auszahlungen, Korrekturen",
          "Manuelle Buchungen auf die Konten Überstunden (Std.), JAZ (Std.) und Urlaub (Tage). Positiv = Gutschrift, "
          "negativ = Abbau/Auszahlung. Die Konten rechnen im Blatt „Überstunden und JAZ“.")
    put(ws, "A3", "Beispiel: Mustermann · 01.01.2027 · Überstunden · Vortrag Vorjahr · 12,5   bzw.   … · Urlaub · "
                  "Vortrag Vorjahr · 3 (Resturlaub aus dem Vorjahr).", size=8, color=MUTED, italic=True)
    labels = ["Mitarbeiter", "Datum", "Konto", "Buchungsart", "Wert", "Bemerkung", "Monat", "Prüfung"]
    hr = ZB_FIRST - 1
    header_row(ws, hr, 1, labels, inputs=(0, 1, 2, 3, 4, 5))
    last = ZB_FIRST + CAP_ZB - 1
    demo_rows = demo_rows or []
    for i in range(CAP_ZB):
        r = ZB_FIRST + i
        if i < len(demo_rows):
            for j, v in enumerate(demo_rows[i]):
                ws.cell(r, 1 + j, v)
        ws[f"G{r}"] = f'=IF(B{r}="","",IF(YEAR(B{r})<Jahr,0,MONTH(B{r})))'
        ws[f"H{r}"] = (f'=IF(A{r}="","",{X}TEXTJOIN("; ",TRUE,IF(COUNTIF(MA_Name,A{r})=0,"Name unbekannt",""),'
                       f'IF(C{r}="","Konto fehlt",""),IF(E{r}="","Wert fehlt","")))')
        for c in "ABCDEFGH":
            style(ws[f"{c}{r}"], size=9, h="center" if c not in "AFH" else None)
        ws[f"B{r}"].number_format = DATE_FMT
        ws[f"E{r}"].number_format = "0.00"
    book.table(ws, "tbl_ZB", f"A{hr}:H{last}")
    rng = lambda c: f"{c}{ZB_FIRST}:{c}{last}"
    for n, c in [("ZB_MA", "A"), ("ZB_Datum", "B"), ("ZB_Konto", "C"), ("ZB_Art", "D"), ("ZB_Wert", "E"),
                 ("ZB_Monat", "G")]:
        book.name_ws(n, ws, rng(c))
    book.dv_list(ws, rng("A"), "=MA_Liste", error=False)
    book.dv_list(ws, rng("C"), "=Konten")
    book.dv_list(ws, rng("D"), "=BuchArten")
    book.dv_date(ws, rng("B"))
    book.cf(ws, rng("H"), f'H{ZB_FIRST}<>""', bg="FCE4D6", color="C00000")
    col_widths(ws, {"A": 18, "B": 12, "C": 13, "D": 16, "E": 10, "F": 34, "G": 8, "H": 28})
    ws.freeze_panes = f"B{ZB_FIRST}"
    ws.sheet_properties.tabColor = "FFC000"
