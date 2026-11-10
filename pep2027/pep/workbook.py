"""Baut die komplette Arbeitsmappe (ohne VBA) mit openpyxl."""
import datetime as dt

from openpyxl import Workbook
from openpyxl.styles import Font

from . import model as M
from .ctx import (CODE_NAMES, ORDER, S_START, Ctx)
from .source import apply_names, read_teams

D = dt.date
BSP = "BEISPIEL – bitte löschen"


def example_data():
    abw = [
        ("Ma7", D(2027, 1, 18), D(2027, 1, 22), "Urlaub", BSP),
        ("Ma13", D(2027, 1, 6), D(2027, 1, 8), "Krankheit", BSP),
        ("Ma20", D(2027, 1, 12), D(2027, 1, 13), "Unterrichte", BSP),
        ("Ma17", D(2027, 1, 15), D(2027, 1, 15), "JAZ + 99", BSP),
        ("Ma24", D(2027, 2, 1), D(2027, 2, 5), "Urlaub genehmigt eTime", BSP),
        ("Ma26", D(2027, 1, 26), D(2027, 1, 26), "FDL-Einweisungen", BSP),
        ("Ma10", D(2027, 1, 28), D(2027, 1, 28), "Kabeleinweisungen", BSP),
        ("Ma9", D(2027, 1, 27), D(2027, 1, 27), "Besonderheiten", "Arzt"),
    ]
    res = [
        ("Azubi 1", "AZUBIS", D(2027, 9, 1), D(2027, 12, 31), "Blau", "BEISPIEL – Ausbildungsbeginn"),
        ("Zusatz-MA 1", "Zusätzliche MA", D(2027, 3, 1), D(2027, 3, 31), "Rot", "BEISPIEL – Unterstützung"),
    ]
    fs = [("Rot", None, D(2027, 1, 5), D(2027, 1, 8), "BEISPIEL: Dienstag arbeiten, stattdessen Freitag frei")]
    ver = [(D(2027, 1, 18), D(2027, 1, 21), "Ma12", "AV", "Ma7", "BEISPIEL – Urlaubsvertretung")]
    return abw, res, fs, ver


def people_from(src, names_csv):
    people = apply_names(read_teams(src), names_csv)
    people.append({"Name": "Disponent", "Team": None, "Funktion": "Disponent",
                   "Bemerkung": "Platzhalter – Namen eintragen"})
    people.append({"Name": "Azubi 1", "Team": "Blau", "Funktion": "AZUBI", "Bemerkung": "BEISPIEL – anpassen"})
    people.append({"Name": "Zusatz-MA 1", "Team": None, "Funktion": "Zusätzliche MA",
                   "Bemerkung": "BEISPIEL – anpassen"})
    return people


def build_workbook(src, names_csv=None, parts="all"):
    from . import sh_core, sh_eval, sh_inputs, sh_master, sh_settings

    wb = Workbook()
    wb.remove(wb.active)
    normal = wb._named_styles["Normal"]
    normal.font = Font(name="Arial", size=10)
    for name in ORDER:
        ws = wb.create_sheet(name)
        ws.sheet_properties.codeName = CODE_NAMES[name]
    wb.code_name = "ThisWorkbook"
    ctx = Ctx(wb)
    people = people_from(src, names_csv)
    abw, res, fs, ver = example_data()
    ctx.people = people

    # Makrostatus (wird von VBA beim Oeffnen ueberschrieben)
    ws = wb[S_START]
    ws["B5"] = '=IFERROR(IF(PEPMakrosAktiv(),"Makros aktiv: Doppelklick (ein Tag) oder Rechtsklick (markierte Tage) im Jahreskalender öffnet die Eingabe – sie wird automatisch in die Tabellen geschrieben.",""),"⚠ Makros nicht aktiv: Eingaben über die Tabellen (z. B. Abwesenheiten) – der Kalender zeigt sie automatisch an.")'
    ctx.name("Makrostatus", f"{S_START}!$B$5")

    sh_settings.build(ctx)
    sh_master.build_mitarbeiter(ctx, people)
    sh_master.build_matrix(ctx, people)
    sh_inputs.build_abwesenheiten(ctx, abw)
    sh_inputs.build_ressourcen(ctx, res)
    sh_inputs.build_freistellungen(ctx, fs)
    sh_inputs.build_vertretungen(ctx, ver)
    sh_core.build_schicht(ctx)
    sh_core.build_berechnung(ctx)
    sh_core.build_kalender(ctx)
    sh_eval.build_einsatz(ctx)
    sh_eval.build_bedarf(ctx)
    if parts == "all":
        from . import sh_reports, sh_print, sh_start
        sh_reports.build_all(ctx)
        sh_print.build_all(ctx)
        sh_start.build_start(ctx)
        sh_start.build_dashboard(ctx)
        apply_protection(ctx)
    wb.active = 0
    return wb, ctx


def apply_protection(ctx):
    """Kalender und Berechnung mit Kennwort (Synchronisierung), Auswertungen ohne Kennwort schuetzen."""
    from openpyxl.utils import get_column_letter as gl
    from .ctx import (S_BED, S_BER, S_DASH, S_DB, S_DJ, S_DM, S_DT, S_EIN, S_JAHR, S_KAL, S_MON, S_SCH, S_STA,
                      S_TEAM, S_UK)
    from .style import protect
    wb = ctx.wb
    c0, cn = M.C_DAY0, M.C_DAYN
    protect(wb[S_KAL], "pep", ["C1", f"{c0}10:{cn}10", f"{c0}{M.ROWN + 4}:{cn}{M.ROWN + 7}"], autofilter=False)
    protect(wb[S_BER], "pep", autofilter=False)
    for name in (S_SCH, S_EIN, S_BED, S_STA, S_UK, S_JAHR, S_TEAM, S_DJ, S_DB):
        protect(wb[name])
    protect(wb[S_MON], unlocked=["C3"])
    protect(wb[S_DASH], unlocked=["C3", "F3"])
    protect(wb[S_DM], unlocked=["C2"])
    protect(wb[S_DT], unlocked=["C2", "E2"])
