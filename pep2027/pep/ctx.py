"""Gemeinsamer Kontext: Blattnamen, Code-Namen (VBA) und definierte Namen."""
from openpyxl.workbook.defined_name import DefinedName

from .style import q

S_START, S_DASH = "Start", "Dashboard"
S_MA, S_FM, S_ABW = "Mitarbeiter", "Funktionsmatrix", "Abwesenheiten"
S_KAL, S_SCH, S_EIN = "Jahreskalender", "Schichtplanung", "Einsatzfähigkeit"
S_RES, S_FS, S_VER = "Ressourcen", "Freistellungen", "Vertretungen"
S_BED, S_STA = "Personalbedarf", "Statistik"
S_MON, S_JAHR, S_UK, S_TEAM = "Monatsbericht", "Jahresbericht", "Urlaub-Krank", "Teamübersicht"
S_DJ, S_DM, S_DT, S_DB = "Druck Jahr", "Druck Monat", "Druck Team", "Druck Bereitschaft"
S_SET, S_BER = "Einstellungen", "Berechnung"

ORDER = [S_START, S_DASH, S_MA, S_FM, S_ABW, S_KAL, S_SCH, S_EIN, S_RES, S_FS, S_VER, S_BED, S_STA,
         S_MON, S_JAHR, S_UK, S_TEAM, S_DJ, S_DM, S_DT, S_DB, S_SET, S_BER]

CODE_NAMES = {S_START: "shStart", S_DASH: "shDashboard", S_MA: "shMitarbeiter", S_FM: "shMatrix",
              S_ABW: "shAbwesenheiten", S_KAL: "shKalender", S_SCH: "shSchicht", S_EIN: "shEinsatz",
              S_RES: "shRessourcen", S_FS: "shFreistellung", S_VER: "shVertretung", S_BED: "shBedarf",
              S_STA: "shStatistik", S_MON: "shMonat", S_JAHR: "shJahr", S_UK: "shUrlaubKrank",
              S_TEAM: "shTeam", S_DJ: "shDruckJahr", S_DM: "shDruckMonat", S_DT: "shDruckTeam",
              S_DB: "shDruckBereitschaft", S_SET: "shEinstellungen", S_BER: "shBerechnung"}

# Beschreibung fuer Startseite / README
DESCR = {
    S_DASH: "Management-Dashboard mit Kennzahlen, Ampeln und Diagrammen",
    S_MA: "Blatt 1 – Mitarbeiterstammdaten (Personalnummer, Name, Team, Funktion)",
    S_FM: "Blatt 2 – Team- und Funktionsmatrix (JA/NEIN je Qualifikation)",
    S_ABW: "Blatt 3 – Abwesenheiten (zentrale Datenquelle)",
    S_KAL: "Blatt 4 – Jahreskalender im PEP-Layout (nur Abweichungen, Bereitschaft grau)",
    S_SCH: "Blatt 5 – Schicht- und Bereitschaftsplanung (Rotation, Freistellung, Wochenende, Feiertage)",
    S_EIN: "Blatt 6 – Einsatzfähigkeitsprüfung mit automatischem Ersatzvorschlag",
    S_RES: "Ressourcen: Zusätzliche MA und AZUBIS (keine Abwesenheiten)",
    S_FS: "Manuelle Anpassungen der Freistellungen (haben Vorrang)",
    S_VER: "Vertretungen in der Bereitschaft",
    S_BED: "Personalbedarfsplanung Tag / Woche / Monat / Jahr",
    S_STA: "Statistik je Mitarbeiter",
    S_MON: "Monatsbericht (Monat wählbar)",
    S_JAHR: "Jahresbericht, Dienst- und Nachtdienstverteilung",
    S_UK: "Urlaubs- und Krankenstatistik je Mitarbeiter und Monat",
    S_TEAM: "Teamübersicht und Einsatzfähigkeitsbericht",
    S_DJ: "Druckansicht Jahresübersicht (DIN A3 quer)",
    S_DM: "Druckansicht Monatsübersicht (DIN A4 quer)",
    S_DT: "Druckansicht Teamübersicht (DIN A4 quer)",
    S_DB: "Druckansicht Bereitschaftsübersicht im Layout des Bereitschaftskalenders (DIN A3/A4 quer)",
    S_SET: "Einstellungen: Jahr, Rotation, Sollbesetzung, Kürzel, Feiertage Hessen",
}


class Ctx:
    def __init__(self, wb):
        self.wb = wb
        self.names = {}

    def ws(self, name):
        return self.wb[name]

    def name(self, nm, ref):
        """Definierten Namen anlegen. ref: 'Blatt!$A$1' oder 'tbl[Spalte]'."""
        if "!" in ref and not ref.startswith("'"):
            sheet, rng = ref.split("!", 1)
            ref = f"{q(sheet)}!{rng}"
        self.wb.defined_names[nm] = DefinedName(nm, attr_text=ref)
        self.names[nm] = ref
        return nm

    def abs_ref(self, sheet, rng):
        return f"{q(sheet)}!{rng}"
