"""Stellt das VBA-Projekt zusammen: Dokumentmodule je Blatt, modPEP (Logik), modKonfig (generiert)."""
from pathlib import Path

from . import model as M
from .ctx import CODE_NAMES, ORDER, S_BER, S_KAL
from .sh_core import kal_grid_formula_r1c1
from .vbaproject import build

VBA_DIR = Path(__file__).parent / "vba"


def vba_string(s):
    return '"' + s.replace('"', '""') + '"'


def konfig_source():
    formula = kal_grid_formula_r1c1()
    lines = [
        "Option Explicit",
        "' Automatisch erzeugt von build_pep.py – Positionen und Namen des Arbeitsblatt-Layouts",
        f"Public Const PEP_KENNWORT As String = {vba_string('pep')}",
        f"Public Const BLATT_KALENDER As String = {vba_string(S_KAL)}",
        f"Public Const BLATT_BERECHNUNG As String = {vba_string(S_BER)}",
        f"Public Const KAL_ZEILE1 As Long = {M.ROW0}",
        f"Public Const KAL_ZEILEN As Long = {M.CAP}",
        f"Public Const KAL_SPALTE1 As Long = {M.DAY0}",
        f"Public Const KAL_TAGE As Long = {M.NDAYS}",
        "Public Const KAL_DATUMZEILE As Long = 3",
        f"Public Const BER_ABW_VERSATZ As Long = {M.AROW0 - M.ROW0}",
        'Public Const TBL_ABW As String = "tblAbwesenheiten"',
        'Public Const TBL_FS As String = "tblFreistellungen"',
        'Public Const TBL_VER As String = "tblVertretungen"',
        'Public Const TBL_RES As String = "tblRessourcen"',
        'Public Const TBL_ARTEN As String = "tblArten"',
        'Public Const TBL_MA As String = "tblMitarbeiter"',
        'Public Const TBL_FM As String = "tblFunktionsmatrix"',
        f"Public Const ART_BESONDERHEIT As String = {vba_string(M.STATUS_BY_CODE[1][1])}",
        f"Public Const QUALIFIKATIONEN As String = {vba_string('|'.join(M.QUALIS))}",
        f"Public Const KAL_FORMEL_R1C1 As String = {vba_string(formula)}",
    ]
    for ln in lines:
        if len(ln) > 1000:
            raise ValueError("VBA-Zeile zu lang: " + ln[:60])
    return "\r\n".join(lines) + "\r\n"


def module_list():
    mods = [("ThisWorkbook", (VBA_DIR / "ThisWorkbook.cls").read_text(encoding="utf-8"), "workbook")]
    for sheet in ORDER:
        cn = CODE_NAMES[sheet]
        f = VBA_DIR / f"{cn}.cls"
        body = f.read_text(encoding="utf-8") if f.exists() else "Option Explicit\n"
        mods.append((cn, body, "sheet"))
    mods.append(("modPEP", (VBA_DIR / "modPEP.bas").read_text(encoding="utf-8"), "module"))
    mods.append(("modKonfig", konfig_source(), "module"))
    return mods


def build_vba_project(ctx=None):
    return build(module_list())


def export_sources(outdir):
    """Quelltexte (inkl. generiertem modKonfig) fuer den manuellen Import in Excel ablegen.

    Standardmodule als .bas (Datei > Datei importieren), Code der Arbeitsmappen-/Blattmodule als .txt
    (Inhalt in das gleichnamige Objekt im VBA-Editor einfuegen).
    """
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)
    targets = {"ThisWorkbook": "DieseArbeitsmappe", "shKalender": "Jahreskalender", "shMitarbeiter": "Mitarbeiter"}
    written = []
    for name, body, kind in module_list():
        body = body.replace("\r\n", "\n").replace("\n", "\r\n")
        if kind == "module":
            f = out / f"{name}.bas"
            f.write_bytes((f'Attribute VB_Name = "{name}"\r\n' + body).encode("cp1252"))
        elif name in targets:
            f = out / f"{name} ({targets[name]}).txt"
            f.write_bytes(body.encode("cp1252"))
        else:
            continue
        written.append(f.name)
    return written
