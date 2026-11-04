#!/usr/bin/env python3
"""Basisseminare: E-Mail-Entwürfe je Führungskraft aus Kontingenten und Kundenbedarf.

Liest drei Excel-Listen
  1. Kontingente   – Termine/Plätze (z. B. „Kontingente Netz Frankfurt“)
  2. Kundenbedarf  – Bedarfe je Mitarbeiter inkl. Vorgesetztem
  3. Feinplaner    – OE (LS) → Bezirksleiter und Feinplaner (FP)
und schreibt in den Ausgabeordner
  - Basisseminare_E-Mails.html    E-Mail-Entwürfe zum Kopieren nach Outlook,
                                   Klärfälle mit Vorschlägen, Hinweise
  - Basisseminare_Pruefliste.xlsx Zuordnung je Bedarf, Klärfälle, Termine, Plätze/Bedarf
  - ergebnis.json                 dieselben Ergebnisse maschinenlesbar

Regeln
  - Termine: nur Kategorie „BS Basisseminar“, Anfang ab Stichtag, Platz noch frei
  - Bedarfe: Status „Offen“ und PS-Nr. eines verwendeten Termins
  - Führungskraft eindeutig, wenn der Vorgesetzte aus dem Kundenbedarf in der
    Feinplaner-Liste als Bezirksleiter der OE des Mitarbeiters steht
    (OE ohne Präfix verglichen: V.IW-MI-N-FFM-IL 05 ≙ I.IA-MI-N-FFM-IL 05).
    Sonst Klärfall: keine E-Mail, nur Vorschläge, wen man anschreiben kann.
  - Cc: Feinplaner aus den Zeilen der Feinplaner-Liste mit dieser OE und diesem Bezirksleiter
  - je Führungskraft genau eine E-Mail, darin je Basisseminar ein Block
    (Termine + „Folgende deiner Mitarbeiter sind davon betroffen“)

Optional: --manuell <csv> ordnet Klärfälle nach Rückmeldung zu
  (Spalten Abt.;Name;Vorname;Führungskraft;Feinplaner, Trennzeichen „;“,
   mehrere Feinplaner mit „/“ trennen; ohne Name gilt die Zeile für alle
   Klärfälle der OE).
"""
import argparse
import csv
import datetime as dt
import html
import itertools
import json
import re
import sys
import unicodedata
import warnings
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore", module="openpyxl")

# ---------------------------------------------------------------- Einstellungen
FRIST_STANDARD = "31.01.2027"          # Satz „Anspruch verfällt ab …“ (aus der Vorlage)
ABSENDER = "Qualifizierung IH Frankfurt"
VORNAMEN_ALIAS = {"jenny": "jennifer"}  # Kurzformen = gleiche Person (Vorlage: Cc Jennifer Eberhardt)
EINSATZORT_BL = "bezirksleiter"         # Einsatzort-Wert der Zeilen für die Bezirksleitung selbst

KONT_SPALTEN = ["PS-Nr.", "Intern", "Terminnr.", "Kategorie", "Titel", "Anfang", "Ende",
                "Termin / Lieferadresse / Ort", "TNP", "Buchungscode", "Kontingentinhaber",
                "TNP verteilt an"]
BEDARF_SPALTEN = ["Name", "Vorname", "Firma", "Abt.", "KST Bedarf", "Status für Dringlichkeit",
                  "Status", "PS-Nr.", "Titel", "Kunde / Vorgesetzter / Vorname",
                  "Kunde / Vorgesetzter / Name"]
FP_SPALTEN = ["LS", "Einsatzort", "IH Bezirk", "Bezirksleiter", "FP"]

TEXT_STIL = "font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#000000"
TAB_STIL = "font-family:Arial,sans-serif;font-size:8pt;color:#000000"


# ---------------------------------------------------------------- Hilfen
def t(v):
    """Zelltext: getrimmt, Mehrfach-Leerzeichen zusammengefasst."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\s+", " ", str(v)).strip()


def datum(v):
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    s = t(v)
    for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def fmt_dt(d):
    return d.strftime("%d.%m.%Y %H:%M") if d else ""


def aufzaehlung(xs):
    xs = list(xs)
    return xs[0] if len(xs) == 1 else ", ".join(xs[:-1]) + " und " + xs[-1]


def esc(s):
    return html.escape(str(s), quote=True)


def eindeutig(xs):
    """Reihenfolge erhaltend Duplikate (und None) entfernen."""
    out = []
    for x in xs:
        if x is not None and x not in out:
            out.append(x)
    return out


def lies_tabelle(pfad, pflicht):
    wb = openpyxl.load_workbook(pfad, data_only=True)
    ws = wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    kopf = [t(v) for v in next(it)]
    fehlend = [s for s in pflicht if s not in kopf]
    if fehlend:
        sys.exit(f"{Path(pfad).name}: Spalte(n) fehlen: {', '.join(fehlend)}")
    zeilen = []
    for nr, werte in enumerate(it, start=2):
        if all(t(v) == "" for v in werte):
            continue
        z = {k: (werte[i] if i < len(werte) else None) for i, k in enumerate(kopf) if k}
        z["_zeile"] = nr
        zeilen.append(z)
    return zeilen


_OE_PRAEFIX = re.compile(r"^[A-Z]\.[A-Z]{2}(?:-[A-Z]+){3}(?:-(.+))?$")


def oe_kurz(s):
    """'V.IW-MI-N-FFM-IL 05' -> 'IL 05' (Präfix weg, zur Anzeige)."""
    s = t(s).upper()
    m = _OE_PRAEFIX.match(s)
    return (m.group(1) or "").strip() if m else s


def oe_schluessel(s):
    """Vergleichsschlüssel: 'V.IW-MI-N-FFM-IL 05' und 'I.IA-MI-N-FFM-IL 5' -> 'IL 5'."""
    return re.sub(r"(?<![0-9])0+(?=[0-9])", "", oe_kurz(s))


# ---------------------------------------------------------------- Personen
def norm(s):
    s = t(s).lower()
    for a, b in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss")):
        s = s.replace(a, b)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z\- ]+", " ", s)


def tokens(s):
    return [VORNAMEN_ALIAS.get(x, x) for x in norm(s).split()]


def schreibform(s):
    """Schreibweise ohne Komma/Reihenfolge/Groß-Klein – zum Erkennen echter Abweichungen."""
    return tuple(sorted(norm(s).split()))


def lev(a, b):
    if len(a) < len(b):
        a, b = b, a
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def gleiche_person(t1, t2):
    """Gleich viele Namensteile, Reihenfolge egal, je Teil höchstens ein Tippfehler (ab 4 Zeichen)."""
    if not t1 or len(t1) != len(t2):
        return False
    ok = lambda a, b: a == b or (min(len(a), len(b)) >= 4 and lev(a, b) <= 1)
    return any(all(ok(a, b) for a, b in zip(t1, p)) for p in itertools.permutations(t2))


@dataclass(eq=False)
class Person:
    pid: int
    tokens: list
    anzeige: str
    prio: int                       # 1 Kundenbedarf/manuell, 2 „Nachname, Vorname“, 3 sonst
    vorname: str = ""
    schreibweisen: dict = field(default_factory=dict)   # Originaltext -> {Fundstellen}

    @property
    def nachname(self):
        return self.anzeige.split()[-1]

    def varianten(self):
        """Schreibweisen, falls die Person unterschiedlich geschrieben vorkommt (sonst [])."""
        je_form = {}
        for s in sorted(self.schreibweisen):
            je_form.setdefault(schreibform(s), s)
        return sorted(je_form.values()) if len(je_form) > 1 else []


class Personen:
    def __init__(self):
        self.liste = []

    def finde(self, tok):
        for p in self.liste:                       # erst exakt (Reihenfolge egal) …
            if sorted(p.tokens) == sorted(tok):
                return p
        for p in self.liste:                       # … dann mit Tippfehler-Toleranz
            if gleiche_person(tok, p.tokens):
                return p
        return None

    def add(self, roh, anzeige, prio, quelle, vorname=""):
        tok = tokens(anzeige)
        if not tok:
            return None
        p = self.finde(tok)
        if p is None:
            p = Person(len(self.liste), tok, anzeige, prio, vorname)
            self.liste.append(p)
        elif prio < p.prio:
            p.anzeige, p.prio = anzeige, prio
        if vorname and not p.vorname:
            p.vorname = vorname
        p.schreibweisen.setdefault(t(roh), set()).add(quelle)
        return p

    @staticmethod
    def rufname(p):
        return p.vorname or p.anzeige.split()[0]


def personen_aus_zelle(zelle):
    """'Nachname, Vorname' | 'Vorname Nachname' | 'A / B' -> [(roh, anzeige, prio)]."""
    s = t(zelle)
    if not s or re.search(r"\bkeine?n?\b", s.lower()):      # „gibt keinen Feinplaner“
        return []
    out = []
    for teil in re.split(r"\s*/\s*", s):
        teil = teil.strip()
        if not teil:
            continue
        if "," in teil:
            nach, vor = [x.strip() for x in teil.split(",", 1)]
            out.append((teil, f"{vor} {nach}".strip(), 2))
        else:
            out.append((teil, teil, 3))
    return out


def namen(ps):
    return "; ".join(p.anzeige for p in ps)


def bl_uebersicht(zeilen):
    """'Charlotte Meier (FFM / FA; FFM / HBF); Maik Van Benthum (Zeile „Bezirksleiter“ / HBF)'."""
    je_person = {}
    for z in zeilen:
        for p in z.bl:
            je_person.setdefault(p.anzeige, []).append(z.ort)
    return "; ".join(f"{k} ({'; '.join(eindeutig(v))})" for k, v in je_person.items()) or "–"


def oe_uebersicht(zeilen):
    """'OE IS 01 (Hanau / FH/NAH; Hanau / HU)'."""
    je_oe = {}
    for z in zeilen:
        je_oe.setdefault(oe_kurz(z.ls), []).append(z.ort)
    return "; ".join(f"OE {k} ({'; '.join(eindeutig(v))})" for k, v in je_oe.items())


# ---------------------------------------------------------------- Datenklassen
@dataclass
class Platz:
    nr: int                  # laufende Nr. in der Kontingentliste (wie „Nr.“-Spalte der Vorlage)
    zeile: int               # Excel-Zeile
    ps: str
    intern: str
    terminnr: str
    kategorie: str
    titel: str
    anfang: object
    ende: object
    ort: str
    tnp: str
    buchungscode: str
    inhaber: str
    verteilt_an: str
    teilnehmer: str
    grund: str = ""          # leer = verwendet, sonst Ausschlussgrund

    def zellen(self):
        return [str(self.nr), self.ps, self.intern, self.terminnr, self.kategorie, self.titel,
                fmt_dt(self.anfang), fmt_dt(self.ende), self.ort, self.tnp, self.buchungscode,
                self.inhaber, self.verteilt_an]


@dataclass
class Bedarf:
    zeile: int
    name: str
    vorname: str
    firma: str
    abt: str
    kst: str
    dringlichkeit: str
    status: str
    ps: str
    produkt: str
    titel: str
    vg_vorname: str
    vg_name: str

    @property
    def vg(self):
        return f"{self.vg_vorname} {self.vg_name}".strip()

    @property
    def oe(self):
        return oe_schluessel(self.abt)

    @property
    def ma(self):
        return (self.name, self.vorname)


@dataclass
class FPZeile:
    zeile: int
    ls: str
    oe: str
    einsatzort: str
    bezirk: str
    bl_roh: str
    fp_roh: str
    bl: list
    fp: list

    @property
    def ist_bl_zeile(self):
        return self.einsatzort.lower() == EINSATZORT_BL

    @property
    def ort(self):
        teile = ["Zeile „Bezirksleiter“" if self.ist_bl_zeile else self.einsatzort, self.bezirk]
        return " / ".join(x for x in teile if x)


@dataclass
class Entscheidung:
    status: str                 # eindeutig | manuell | Klärfall
    grund: str
    an: object = None
    cc: list = field(default_factory=list)
    hinweise: list = field(default_factory=list)
    gruppe: str = ""            # Klärfall-Gruppe
    vorschlaege: list = field(default_factory=list)
    indiz: str = ""
    bl_text: str = ""


# ---------------------------------------------------------------- Einlesen
def lies_kontingente(pfad, stichtag):
    plaetze = []
    for i, z in enumerate(lies_tabelle(pfad, KONT_SPALTEN), start=1):
        p = Platz(nr=i, zeile=z["_zeile"], ps=t(z["PS-Nr."]), intern=t(z["Intern"]),
                  terminnr=t(z["Terminnr."]), kategorie=t(z["Kategorie"]), titel=t(z["Titel"]),
                  anfang=datum(z["Anfang"]), ende=datum(z["Ende"]),
                  ort=t(z["Termin / Lieferadresse / Ort"]), tnp=t(z["TNP"]),
                  buchungscode=t(z["Buchungscode"]), inhaber=t(z["Kontingentinhaber"]),
                  verteilt_an=t(z["TNP verteilt an"]), teilnehmer=t(z.get("Teilnehmer / Name")))
        if not p.kategorie.upper().startswith("BS"):
            p.grund = f"kein Basisseminar (Kategorie „{p.kategorie}“)"
        elif p.anfang is None:
            p.grund = "kein Anfangsdatum"
        elif p.anfang.date() < stichtag:
            p.grund = f"Termin liegt in der Vergangenheit (Anfang {fmt_dt(p.anfang)})"
        elif p.teilnehmer and p.teilnehmer != p.tnp:
            p.grund = f"Platz bereits vergeben an „{p.teilnehmer}“"
        plaetze.append(p)
    return plaetze


def lies_bedarf(pfad):
    return [Bedarf(zeile=z["_zeile"], name=t(z["Name"]), vorname=t(z["Vorname"]), firma=t(z["Firma"]),
                   abt=t(z["Abt."]), kst=t(z["KST Bedarf"]), dringlichkeit=t(z["Status für Dringlichkeit"]),
                   status=t(z["Status"]), ps=t(z["PS-Nr."]), produkt=t(z.get("Produkt-Nr.")), titel=t(z["Titel"]),
                   vg_vorname=t(z["Kunde / Vorgesetzter / Vorname"]), vg_name=t(z["Kunde / Vorgesetzter / Name"]))
            for z in lies_tabelle(pfad, BEDARF_SPALTEN)]


def lies_feinplaner(pfad, personen):
    zeilen = []
    for z in lies_tabelle(pfad, FP_SPALTEN):
        quelle = f"Feinplaner-Liste Z. {z['_zeile']}"
        bl = [personen.add(r, a, p, quelle + " (Bezirksleiter)") for r, a, p in personen_aus_zelle(z["Bezirksleiter"])]
        fp = [personen.add(r, a, p, quelle + " (FP)") for r, a, p in personen_aus_zelle(z["FP"])]
        zeilen.append(FPZeile(zeile=z["_zeile"], ls=t(z["LS"]), oe=oe_schluessel(z["LS"]),
                              einsatzort=t(z["Einsatzort"]), bezirk=t(z["IH Bezirk"]),
                              bl_roh=t(z["Bezirksleiter"]), fp_roh=t(z["FP"]),
                              bl=[p for p in bl if p], fp=[p for p in fp if p]))
    return zeilen


def lies_manuell(pfad):
    if not pfad:
        return []
    text = Path(pfad).read_text(encoding="utf-8-sig")
    return [{k.strip(): t(v) for k, v in r.items() if k} for r in csv.DictReader(text.splitlines(), delimiter=";")]


# ---------------------------------------------------------------- Zuordnung
def treffer_zeilen(person, zeilen):
    return [z for z in zeilen if person in z.bl]


def fp_aus(zeilen):
    return eindeutig(fp for z in zeilen for fp in z.fp)


def schreibweise_hinweis(person, zeilen, vg_text):
    """Hinweis, wenn der Bezirksleiter in der Feinplaner-Liste anders geschrieben ist als im Kundenbedarf."""
    abweichend = []
    for z in zeilen:
        for teil in re.split(r"\s*/\s*", z.bl_roh):
            if gleiche_person(tokens(teil), person.tokens) and schreibform(teil) != schreibform(vg_text):
                abweichend.append(f"„{teil}“ (Z. {z.zeile})")
    return ([f"Name in der Feinplaner-Liste anders geschrieben: {', '.join(eindeutig(abweichend))} – "
             f"als {person.anzeige} gewertet"] if abweichend else [])


def zuordnen(bedarfe, alle_bedarfe, fp_zeilen, personen, manuell):
    nach_oe = defaultdict(list)
    for z in fp_zeilen:
        nach_oe[z.oe].append(z)

    # Vorgesetzte je OE (alle Bedarfe, je Mitarbeiter einmal) – Indiz für Klärfälle
    vg_je_oe = defaultdict(Counter)
    for (oe, _, _), vg in {(b.oe, b.name, b.vorname): b.vg for b in alle_bedarfe}.items():
        vg_je_oe[oe][vg or "kein Eintrag"] += 1

    ergebnisse = []
    for b in bedarfe:
        zeilen = nach_oe.get(b.oe, [])
        oe = oe_kurz(b.abt)
        vg = personen.add(b.vg, b.vg, 1, f"Kundenbedarf Z. {b.zeile}", b.vg_vorname) if b.vg else None
        bl_text = bl_uebersicht(zeilen)

        if vg is None:
            e = Entscheidung("Klärfall", "Kein Vorgesetzter im Kundenbedarf eingetragen",
                             gruppe=f"{b.abt} – kein Vorgesetzter eingetragen")
        elif not zeilen:
            e = Entscheidung("Klärfall", f"OE {b.abt} steht nicht in der Feinplaner-Liste – "
                             f"Vorgesetzter {vg.anzeige} lässt sich nicht gegenprüfen",
                             gruppe=f"{b.abt} – Vorgesetzter {vg.anzeige}")
        else:
            treffer = treffer_zeilen(vg, zeilen)
            if treffer:
                cc = fp_aus(treffer)
                e = Entscheidung("eindeutig", f"Vorgesetzter {vg.anzeige} ist laut Feinplaner-Liste Bezirksleiter "
                                 f"der OE {oe} ({'; '.join(eindeutig(z.ort for z in treffer))})", an=vg, cc=cc)
                e.hinweise += schreibweise_hinweis(vg, treffer, b.vg)
                if len(cc) > 1:
                    e.hinweise.append(
                        f"Feinplaner für {vg.anzeige} / OE {oe} nicht eindeutig: "
                        + "; ".join(f"{fp.anzeige} ({', '.join(eindeutig(z.ort for z in treffer if fp in z.fp))})"
                                    for fp in cc) + " – alle in Cc, bitte vor dem Versand prüfen")
                if not cc:
                    e.hinweise.append(f"Kein Feinplaner in der Feinplaner-Liste für {vg.anzeige} / OE {oe} – Cc fehlt")
            else:
                e = Entscheidung("Klärfall", f"Vorgesetzter laut Kundenbedarf ({vg.anzeige}) steht in der "
                                 f"Feinplaner-Liste nicht als Bezirksleiter der OE {oe}; dort: {bl_text}",
                                 gruppe=f"{b.abt} – Vorgesetzter {vg.anzeige}")

        # manuelle Zuordnung (nach Rückmeldung)
        for m in manuell:
            per_name = bool(m.get("Name")) and norm(m["Name"]) == norm(b.name) and \
                (not m.get("Vorname") or norm(m["Vorname"]) == norm(b.vorname))
            per_oe = not m.get("Name") and bool(m.get("Abt.")) and oe_schluessel(m["Abt."]) == b.oe \
                and e.status == "Klärfall"
            if not (per_name or per_oe) or not m.get("Führungskraft"):
                continue
            fk = personen.add(m["Führungskraft"], m["Führungskraft"], 1, "manuelle Zuordnung",
                              m["Führungskraft"].split()[0])
            if m.get("Feinplaner"):
                cc = [personen.add(x, x, 1, "manuelle Zuordnung") for x in re.split(r"\s*/\s*", m["Feinplaner"]) if x]
            else:
                cc = fp_aus(treffer_zeilen(fk, zeilen)) or fp_aus(treffer_zeilen(fk, fp_zeilen))
            e = Entscheidung("manuell", f"manuell zugeordnet (vorher: {e.grund})", an=fk, cc=eindeutig(cc))
            if not e.cc:
                e.hinweise.append(f"Kein Feinplaner für {fk.anzeige} gefunden – Cc bitte ergänzen")

        if e.status == "Klärfall":
            e.vorschlaege = vorschlaege(b, vg, zeilen, fp_zeilen, vg_je_oe)
            andere = vg_je_oe[b.oe]
            e.indiz = (f"Vorgesetzte aller Mitarbeiter der OE {oe} im Kundenbedarf: "
                       + ", ".join(f"{k} ({v} MA)" for k, v in andere.most_common()))
        e.bl_text = bl_text
        ergebnisse.append(e)
    return ergebnisse


def vorschlaege(b, vg, zeilen, alle_zeilen, vg_je_oe):
    """Mögliche Empfänger für einen Klärfall – mit Herkunft, damit man selbst entscheiden kann."""
    oe = oe_kurz(b.abt)
    out = []
    if vg is not None:
        treffer = treffer_zeilen(vg, alle_zeilen)
        regulaer = [z for z in treffer if not z.ist_bl_zeile] or treffer
        if regulaer:
            cc = fp_aus(regulaer)
            info = "laut Feinplaner-Liste Bezirksleiter bei " + oe_uebersicht(regulaer)
        else:
            cc = fp_aus(zeilen)
            info = "steht nicht in der Feinplaner-Liste" + (
                f"; Feinplaner der OE {oe} laut Liste: {namen(cc)}" if cc else "")
        anzahl, gesamt = vg_je_oe[b.oe].get(b.vg, 0), sum(vg_je_oe[b.oe].values())
        out.append({"an": vg.anzeige, "cc": [p.anzeige for p in cc],
                    "quelle": f"Vorgesetzter laut Kundenbedarf (bei {anzahl} von {gesamt} Mitarbeitern der OE {oe} "
                              f"so eingetragen); {info}"})
    gesehen = {vg.pid} if vg is not None else set()
    for regulaer in (True, False):
        if not regulaer and out:      # „Bezirksleiter“-Zeilen nur, wenn es sonst keinen Vorschlag gibt
            break
        for z in zeilen:
            if z.ist_bl_zeile == regulaer:
                continue
            for p in z.bl:
                if p.pid in gesehen:
                    continue
                gesehen.add(p.pid)
                eigene = [y for y in zeilen if p in y.bl and y.ist_bl_zeile != regulaer]
                art = "Bezirksleiter" if regulaer else "Eintrag in Zeile „Bezirksleiter“ (vermutlich übergeordnet)"
                out.append({"an": p.anzeige, "cc": [x.anzeige for x in fp_aus(eigene)],
                            "quelle": f"{art} laut Feinplaner-Liste für OE {oe}: " + "; ".join(
                                f"{y.ort} → FP {namen(y.fp) or '–'}" for y in eigene)})
    return out


# ---------------------------------------------------------------- E-Mails
@dataclass
class Email:
    an: Person
    cc: list = field(default_factory=list)
    bedarfe: dict = field(default_factory=lambda: defaultdict(list))   # ps -> [Bedarf]
    hinweise: list = field(default_factory=list)
    seminare: list = field(default_factory=list)                        # [(ps, titel, [Platz])]
    betreff: str = ""

    @property
    def mitarbeiter(self):
        return eindeutig(b.ma for bs in self.bedarfe.values() for b in bs)


def baue_emails(bedarfe, ergebnisse, plaetze_je_ps, frist):
    emails = {}
    for b, e in zip(bedarfe, ergebnisse):
        if e.status not in ("eindeutig", "manuell"):
            continue
        em = emails.setdefault(e.an.pid, Email(an=e.an))
        em.cc = eindeutig(em.cc + e.cc)
        em.bedarfe[b.ps].append(b)
        em.hinweise = eindeutig(em.hinweise + e.hinweise)
    for em in emails.values():
        em.seminare = sorted(((ps, plaetze_je_ps[ps][0].titel, plaetze_je_ps[ps]) for ps in em.bedarfe),
                             key=lambda s: (min(p.anfang for p in s[2]), s[1]))
        for ps in em.bedarfe:
            em.bedarfe[ps].sort(key=lambda b: (norm(b.name), norm(b.vorname)))
        titel = [s[1] for s in em.seminare]
        em.betreff = ("Basisseminar " if len(titel) == 1 else "Basisseminare ") + aufzaehlung(titel)
        vor_frist = [p for s in em.seminare for p in s[2] if p.anfang.date() <= frist]
        if vor_frist:
            termine = eindeutig(f"{p.titel} {p.terminnr} (ab {fmt_dt(p.anfang)[:10]})" for p in vor_frist)
            em.hinweise.append(f"Beginnt vor/an der Frist {frist:%d.%m.%Y}: {'; '.join(termine)} – "
                               "Fristsatz für diesen Termin ggf. anpassen")
    return sorted(emails.values(), key=lambda em: (norm(em.an.nachname), norm(em.an.anzeige)))


def tabelle_html(zeilen):
    out = ['<table cellspacing="0" cellpadding="0" border="0" '
           f'style="border-collapse:collapse;background:#FFFFFF;{TAB_STIL}">']
    for i, z in enumerate(zeilen):
        bg = "#FFFFFF" if i % 2 == 0 else "#F2F2F2"
        zellen = "".join(f'<td style="padding:2px 6px;border:1px solid #FFFFFF;white-space:nowrap;{TAB_STIL}">'
                         f"{esc(c)}</td>" for c in z)
        out.append(f'<tr bgcolor="{bg}" style="background:{bg}">{zellen}</tr>')
    out.append("</table>")
    return "".join(out)


def ma_zellen(i, b):
    return [str(i), b.name, b.vorname, b.firma, b.abt, b.kst, b.dringlichkeit]


def email_teile(em, frist_text):
    """Inhalt der E-Mail als Folge von ('text'|'leer'|'titel'|'tabelle', wert) – für HTML und Text."""
    einzeln = len(em.seminare) == 1
    teile = [("text", f"Hallo {Personen.rufname(em.an)},"), ("leer", "")]
    if einzeln:
        teile.append(("text", f"anbei erhältst du die Termine für das anstehende Basisseminar {em.seminare[0][1]}:"))
    else:
        teile.append(("text", "anbei erhältst du die Termine für die anstehenden Basisseminare "
                              f"{aufzaehlung(s[1] for s in em.seminare)}:"))
    for ps, titel, plaetze in em.seminare:
        if not einzeln:
            teile += [("leer", ""), ("titel", f"Basisseminar {titel}")]
        teile += [("tabelle", [p.zellen() for p in plaetze]), ("leer", ""),
                  ("text", "Folgende deiner Mitarbeiter sind davon betroffen:"),
                  ("tabelle", [ma_zellen(i, b) for i, b in enumerate(em.bedarfe[ps], 1)])]
    teile += [("leer", ""),
              ("text", f"Bitte beachte, dass der Anspruch auf diese Plätze ab dem {frist_text} verfällt."),
              ("text", "Ich bitte dich daher um eine zeitnahe Rückmeldung/Buchung."), ("leer", ""),
              ("text", "Viele Grüße")]
    return teile


def email_html(em, frist_text):
    out = []
    for art, wert in email_teile(em, frist_text):
        if art == "tabelle":
            out.append(tabelle_html(wert))
        elif art == "leer":
            out.append(f'<div style="{TEXT_STIL}"><br></div>')
        elif art == "titel":
            out.append(f'<div style="{TEXT_STIL}"><b>{esc(wert)}</b></div>')
        else:
            out.append(f'<div style="{TEXT_STIL}">{esc(wert)}</div>')
    return "".join(out)


def email_text(em, frist_text):
    return "\n".join("\n".join("\t".join(z) for z in wert) if art == "tabelle" else wert
                     for art, wert in email_teile(em, frist_text))


# ---------------------------------------------------------------- Klärfälle gruppieren
def klaerfall_gruppen(bedarfe, ergebnisse):
    gruppen = {}
    for b, e in zip(bedarfe, ergebnisse):
        if e.status != "Klärfall":
            continue
        g = gruppen.setdefault(e.gruppe, {"titel": e.gruppe, "grund": e.grund, "bedarfe": [],
                                          "vorschlaege": e.vorschlaege, "indiz": e.indiz})
        g["bedarfe"].append(b)
    for g in gruppen.values():
        g["bedarfe"].sort(key=lambda b: (norm(b.name), norm(b.vorname), b.titel))
        g["mitarbeiter"] = eindeutig(b.ma for b in g["bedarfe"])
    return sorted(gruppen.values(), key=lambda g: (-len(g["bedarfe"]), g["titel"]))


# ---------------------------------------------------------------- HTML-Übersicht
SEITE_CSS = """
:root{--bg:#f6f5f2;--panel:#ffffff;--ink:#1d1c1a;--ink2:#55534e;--muted:#7d7a72;--line:#e3e1da;
--accent:#1f4e78;--warn-bg:#fff6e0;--warn-line:#e6b84d;--warn-ink:#5e4100;--bad-bg:#fdecea;--bad-line:#d9534f;
--bad-ink:#7a1f1c;--ok:#2e7d32;--chip:#eef2f7}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]){--bg:#17181a;--panel:#212226;--ink:#ecebe7;
--ink2:#b9b7b0;--muted:#9a978f;--line:#34353a;--accent:#8db8e8;--warn-bg:#3a3018;--warn-line:#a8842f;--warn-ink:#f1d48c;
--bad-bg:#3d2020;--bad-line:#d46a66;--bad-ink:#f3b7b4;--ok:#7bc47f;--chip:#2b3340}}
:root[data-theme="dark"]{--bg:#17181a;--panel:#212226;--ink:#ecebe7;--ink2:#b9b7b0;--muted:#9a978f;--line:#34353a;
--accent:#8db8e8;--warn-bg:#3a3018;--warn-line:#a8842f;--warn-ink:#f1d48c;--bad-bg:#3d2020;--bad-line:#d46a66;
--bad-ink:#f3b7b4;--ok:#7bc47f;--chip:#2b3340}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif}
main{max-width:1240px;margin:0 auto;padding:28px 16px 80px}
h1{font-size:26px;margin:0 0 4px;letter-spacing:-.01em}
h2{font-size:19px;margin:40px 0 12px;padding-top:8px;border-top:1px solid var(--line)}
h3{font-size:16px;margin:0}
.sub{color:var(--ink2);margin:0 0 18px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin:18px 0}
.kpi{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:12px 14px}
.kpi b{display:block;font-size:26px;line-height:1.1;font-variant-numeric:tabular-nums}
.kpi span{color:var(--ink2);font-size:13px}
.kpi.bad b{color:var(--bad-line)}
.card{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:16px;margin:0 0 16px}
.card.klaer{border-left:4px solid var(--bad-line)}
.meta{display:grid;grid-template-columns:64px minmax(0,1fr) auto;gap:6px 10px;align-items:center;margin:10px 0 12px}
.meta .k{color:var(--muted);font-size:13px}
.meta .v{font-weight:600;overflow-wrap:anywhere}
button{font:inherit;font-size:13px;border:1px solid var(--line);background:var(--chip);color:var(--ink);
border-radius:7px;padding:4px 10px;cursor:pointer;white-space:nowrap}
button:hover{border-color:var(--accent)}
button.primary{background:var(--accent);border-color:var(--accent);color:#fff;font-weight:600;padding:6px 14px}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]) button.primary{color:#10161f}}
:root[data-theme="dark"] button.primary{color:#10161f}
button.done{border-color:var(--ok);color:var(--ok)}
button.primary.done{background:var(--ok);border-color:var(--ok);color:#fff}
.warn{background:var(--warn-bg);border:1px solid var(--warn-line);color:var(--warn-ink);border-radius:8px;
padding:8px 12px;margin:0 0 12px;font-size:14px}
.warn ul{margin:4px 0 0;padding-left:20px}
.paper{background:#fff;color:#000;border:1px solid #d6d4cc;border-radius:8px;padding:18px 20px;overflow-x:auto}
.sig{color:#7d7a72;font:italic 13px Aptos,Calibri,Arial,sans-serif;margin-top:6px}
.bar{display:flex;gap:8px;justify-content:flex-end;margin-top:10px;flex-wrap:wrap}
table.list{border-collapse:collapse;width:100%;font-size:14px}
table.list th,table.list td{border-bottom:1px solid var(--line);padding:6px 8px;text-align:left;vertical-align:top}
table.list th{color:var(--ink2);font-weight:600;font-size:13px}
.wrap{overflow-x:auto}
table.nw td,table.nw th{white-space:nowrap}
.tag{display:inline-block;font-size:12px;font-weight:500;padding:1px 8px;border-radius:99px;background:var(--bad-bg);
color:var(--bad-ink);border:1px solid var(--bad-line);margin-left:8px;vertical-align:2px}
.vorschlag{border:1px dashed var(--line);border-radius:8px;padding:8px 12px;margin:8px 0}
.vorschlag .q{color:var(--ink2);font-size:13px}
.indiz{color:var(--ink2);font-size:13px;margin:8px 0 0}
.num{text-align:right;font-variant-numeric:tabular-nums}
.over{color:var(--bad-line);font-weight:700}
ol.toc{padding-left:22px;margin:0}
ol.toc li{margin:4px 0}
a{color:var(--accent)}
.muted{color:var(--muted)}
"""

SEITE_JS = """
function fertig(btn){const alt=btn.textContent;btn.textContent='Kopiert ✓';btn.classList.add('done');
setTimeout(()=>{btn.textContent=alt;btn.classList.remove('done')},1600)}
function kopiereText(text, btn){
  const f=()=>{const ta=document.createElement('textarea');ta.value=text;document.body.appendChild(ta);ta.select();
    document.execCommand('copy');ta.remove();fertig(btn)};
  if(navigator.clipboard&&window.isSecureContext){navigator.clipboard.writeText(text).then(()=>fertig(btn),f)}else{f()}
}
function kopiereMail(id, btn){
  const el=document.getElementById(id);
  const f=()=>{const r=document.createRange();r.selectNodeContents(el);const s=getSelection();s.removeAllRanges();
    s.addRange(r);document.execCommand('copy');s.removeAllRanges();fertig(btn)};
  try{
    if(!(navigator.clipboard&&window.ClipboardItem&&window.isSecureContext)) return f();
    const text=document.getElementById(id+'-txt').textContent;
    navigator.clipboard.write([new ClipboardItem({'text/html':new Blob([el.innerHTML],{type:'text/html'}),
      'text/plain':new Blob([text],{type:'text/plain'})})]).then(()=>fertig(btn),f);
  }catch(e){f()}
}
document.addEventListener('click',e=>{const b=e.target.closest('button[data-copy]');if(b)kopiereText(b.dataset.copy,b)});
"""


def seite_html(emails, gruppen, plaetze, stats, args):
    h = []
    a = h.append
    a("<!doctype html><html lang='de'><head><meta charset='utf-8'>"
      "<meta name='viewport' content='width=device-width,initial-scale=1'>"
      f"<title>Basisseminare E-Mails</title><style>{SEITE_CSS}</style></head><body><main>")
    a("<h1>Basisseminare – E-Mails an die Führungskräfte</h1>")
    a(f"<p class='sub'>Erzeugt am {dt.datetime.now():%d.%m.%Y %H:%M} · Stichtag {args.stichtag:%d.%m.%Y} · "
      f"Frist im Text: {esc(args.frist)} · Absender: {ABSENDER} (im Auftrag)<br>"
      f"Quellen: {esc(Path(args.kontingente).name)} · {esc(Path(args.bedarf).name)} · "
      f"{esc(Path(args.feinplaner).name)}</p>")

    a("<div class='kpis'>")
    a(f"<div class='kpi'><b>{len(emails)}</b><span>E-Mails, je Führungskraft eine</span></div>")
    a(f"<div class='kpi'><b>{stats['in_mails']}</b><span>Bedarfe in E-Mails ({stats['ma_in_mails']} Mitarbeiter)</span></div>")
    a(f"<div class='kpi bad'><b>{stats['klaer']}</b><span>Bedarfe zu klären ({stats['ma_klaer']} Mitarbeiter)</span></div>")
    a(f"<div class='kpi'><b>{stats['plaetze_ok']}</b><span>Plätze angeboten ({stats['plaetze_raus']} Zeile(n) "
      "ausgeschlossen)</span></div>")
    a("</div>")

    # Klärfälle
    a("<h2>Vor dem Versand klären</h2>")
    if not gruppen:
        a("<p>Keine Klärfälle – alle Bedarfe sind eindeutig zugeordnet.</p>")
    else:
        a("<p class='sub'>Diese Mitarbeiter stehen in <b>keiner</b> E-Mail, weil die Führungskraft nicht eindeutig ist. "
          "Bitte erst klären, wer zuständig ist – danach kommen sie in die passende E-Mail.</p>")
    for g in gruppen:
        n_b = len(g["bedarfe"])
        a(f"<div class='card klaer'><h3>{esc(g['titel'])}<span class='tag'>{len(g['mitarbeiter'])} MA · "
          f"{n_b} Bedarf{'e' if n_b != 1 else ''}</span></h3>")
        a(f"<p style='margin:6px 0 10px'>{esc(g['grund'])}</p>")
        a("<div class='wrap'><table class='list nw'><tr><th>Name</th><th>Vorname</th><th>Abt.</th><th>Basisseminar</th>"
          "<th>Dringlichkeit</th><th class='num' title='Zeile im Kundenbedarf'>Zeile</th></tr>")
        for b in g["bedarfe"]:
            a(f"<tr><td>{esc(b.name)}</td><td>{esc(b.vorname)}</td><td>{esc(b.abt)}</td><td>{esc(b.titel)} "
              f"<span class='muted'>({esc(b.ps)})</span></td><td>{esc(b.dringlichkeit) or '–'}</td>"
              f"<td class='num'>{b.zeile}</td></tr>")
        a("</table></div><p style='margin:12px 0 4px'><b>Vorschläge, wen man anschreiben kann:</b></p>")
        for i, v in enumerate(g["vorschlaege"], 1):
            a(f"<div class='vorschlag'><b>{i}. An {esc(v['an'])}</b> · Cc {esc('; '.join(v['cc']) or '–')}"
              f"<div class='q'>{esc(v['quelle'])}</div></div>")
        if g["indiz"]:
            a(f"<p class='indiz'>{esc(g['indiz'])}</p>")
        a("</div>")

    # Hinweise
    a("<h2>Hinweise</h2><div class='card'><ul style='margin:0;padding-left:20px'>")
    for p in plaetze:
        if p.grund:
            a(f"<li>Kontingent Nr. {p.nr} (Zeile {p.zeile}) – {esc(p.titel)} {esc(p.terminnr)}, TNP {esc(p.tnp)}: "
              f"<b>nicht angeboten</b>, {esc(p.grund)}.</li>")
    for x in stats["hinweise"]:
        a(f"<li>{esc(x)}</li>")
    a("</ul></div>")

    # Plätze vs. Bedarf
    a("<h2>Plätze und Bedarfe je Basisseminar</h2><div class='card wrap'><table class='list'><tr><th>Basisseminar</th>"
      "<th>PS-Nr.</th><th>Termine</th><th class='num'>Plätze</th><th class='num'>Bedarfe in E-Mails</th>"
      "<th class='num'>Bedarfe Klärfall</th><th class='num'>Bedarfe gesamt</th></tr>")
    for r in stats["plaetze_bedarf"]:
        over = " class='num over'" if r["gesamt"] > r["plaetze"] else " class='num'"
        a(f"<tr><td>{esc(r['titel'])}</td><td>{esc(r['ps'])}</td><td>{esc(r['termine'])}</td>"
          f"<td class='num'>{r['plaetze']}</td><td class='num'>{r['eindeutig']}</td><td class='num'>{r['klaer']}</td>"
          f"<td{over}>{r['gesamt']}</td></tr>")
    a("</table><p class='muted' style='font-size:13px;margin:8px 0 0'>Fett/rot: mehr Bedarfe als Plätze – "
      "die Plätze gehen an die zuerst eingehenden Buchungen.</p></div>")

    # E-Mails
    a("<h2>E-Mails</h2><div class='card'><ol class='toc'>")
    for i, em in enumerate(emails, 1):
        a(f"<li><a href='#mail{i}'>{esc(em.an.anzeige)}</a> – {esc(em.betreff)} "
          f"<span class='muted'>({len(em.mitarbeiter)} MA · Cc {esc(namen(em.cc)) or '–'})</span>"
          + (" <span class='tag'>prüfen</span>" if em.hinweise else "") + "</li>")
    a("</ol></div>")
    for i, em in enumerate(emails, 1):
        mid = f"mail{i}"
        a(f"<div class='card' id='{mid}'><h3>{i}. {esc(em.an.anzeige)}</h3><div class='meta'>")
        for k, v in (("An", em.an.anzeige), ("Cc", namen(em.cc)), ("Betreff", em.betreff)):
            knopf = f"<button data-copy='{esc(v)}'>kopieren</button>" if v else "<span></span>"
            a(f"<span class='k'>{k}</span><span class='v'>{esc(v) or '–'}</span>{knopf}")
        a("</div>")
        if em.hinweise:
            a("<div class='warn'><b>Vor dem Versand prüfen</b> (nicht Teil der E-Mail):<ul>"
              + "".join(f"<li>{esc(x)}</li>" for x in em.hinweise) + "</ul></div>")
        a(f"<div class='paper'><div id='{mid}-body'>{email_html(em, args.frist)}</div>"
          "<div class='sig'>(hier folgt deine Outlook-Signatur)</div></div>")
        a(f"<pre id='{mid}-body-txt' hidden>{esc(email_text(em, args.frist))}</pre>")
        a(f"<div class='bar'><button class='primary' onclick=\"kopiereMail('{mid}-body', this)\">"
          "E-Mail-Text kopieren</button></div></div>")
    a(f"<script>{SEITE_JS}</script></main></body></html>")
    return "".join(h)


# ---------------------------------------------------------------- Excel-Prüfliste
KOPF_FILL = PatternFill("solid", fgColor="1F4E78")
STATUS_FILL = {"eindeutig": PatternFill("solid", fgColor="E2F0D9"),
               "manuell": PatternFill("solid", fgColor="DDEBF7"),
               "Klärfall": PatternFill("solid", fgColor="FCE4D6")}


def blatt(wb, name, kopf, zeilen, breiten, wrap=""):
    ws = wb.create_sheet(name)
    ws.append(kopf)
    for z in zeilen:
        ws.append(z)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = KOPF_FILL
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    if zeilen:
        ws.auto_filter.ref = ws.dimensions
    for i, w in enumerate(breiten, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.alignment = Alignment(vertical="top", wrap_text=c.column_letter in wrap)
    return ws


def schreibe_excel(pfad, emails, gruppen, plaetze, bedarfe, ergebnisse, personen, stats):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    blatt(wb, "E-Mails", ["Nr", "An", "Cc", "Betreff", "Basisseminare", "Mitarbeiter", "Bedarfe", "Vor dem Versand prüfen"],
          [[i, em.an.anzeige, namen(em.cc), em.betreff, "\n".join(s[1] for s in em.seminare), len(em.mitarbeiter),
            sum(len(bs) for bs in em.bedarfe.values()), "\n".join(em.hinweise)] for i, em in enumerate(emails, 1)],
          [5, 22, 34, 60, 32, 12, 10, 80], wrap="CDEH")

    ws = blatt(wb, "Zuordnung", ["Zeile Kundenbedarf", "Name", "Vorname", "Abt.", "KST Bedarf", "Dringlichkeit", "PS-Nr.",
                                 "Titel", "Vorgesetzter (Kundenbedarf)", "Bezirksleiter laut Feinplaner-Liste (OE)",
                                 "Ergebnis", "Begründung", "An", "Cc", "Hinweise"],
               [[b.zeile, b.name, b.vorname, b.abt, b.kst, b.dringlichkeit, b.ps, b.titel, b.vg or "(leer)", e.bl_text,
                 e.status, e.grund, e.an.anzeige if e.an else "", namen(e.cc), "\n".join(e.hinweise)]
                for b, e in zip(bedarfe, ergebnisse)],
               [10, 16, 14, 22, 15, 30, 11, 30, 22, 50, 11, 60, 20, 30, 60], wrap="JLNO")
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 11).fill = STATUS_FILL.get(ws.cell(r, 11).value, STATUS_FILL["Klärfall"])

    rows = []
    for g in gruppen:
        vs = "\n".join(f"{i}. An {v['an']} · Cc {'; '.join(v['cc']) or '–'} – {v['quelle']}"
                       for i, v in enumerate(g["vorschlaege"], 1))
        for b in g["bedarfe"]:
            rows.append([g["titel"], b.zeile, b.name, b.vorname, b.abt, b.ps, b.titel, b.dringlichkeit,
                         b.vg or "(leer)", g["grund"], vs, g["indiz"]])
    blatt(wb, "Klärfälle", ["Gruppe", "Zeile Kundenbedarf", "Name", "Vorname", "Abt.", "PS-Nr.", "Titel", "Dringlichkeit",
                            "Vorgesetzter (Kundenbedarf)", "Grund", "Vorschläge", "Indiz"], rows,
          [34, 10, 16, 14, 22, 11, 30, 30, 22, 60, 90, 60], wrap="AJKL")

    blatt(wb, "Termine", ["Nr", "Zeile", "PS-Nr.", "Intern", "Terminnr.", "Kategorie", "Titel", "Anfang", "Ende", "Ort",
                          "TNP", "Buchungscode", "Angeboten", "Grund"],
          [[p.nr, p.zeile, p.ps, p.intern, p.terminnr, p.kategorie, p.titel, fmt_dt(p.anfang), fmt_dt(p.ende), p.ort,
            p.tnp, p.buchungscode, "nein" if p.grund else "ja", p.grund] for p in plaetze],
          [5, 7, 11, 15, 10, 20, 40, 17, 17, 12, 12, 15, 10, 55])

    blatt(wb, "Plätze vs. Bedarf", ["PS-Nr.", "Titel", "Termine", "Plätze", "Bedarfe in E-Mails", "Bedarfe Klärfall",
                                    "Bedarfe gesamt"],
          [[r["ps"], r["titel"], r["termine"], r["plaetze"], r["eindeutig"], r["klaer"], r["gesamt"]]
           for r in stats["plaetze_bedarf"]], [11, 32, 60, 8, 12, 12, 12], wrap="C")

    blatt(wb, "Namensabgleich", ["Person (so verwendet)", "Schreibweisen", "Fundstellen"],
          [[p.anzeige, "\n".join(p.varianten()),
            "\n".join(f"{k}: {', '.join(sorted(v))}" for k, v in sorted(p.schreibweisen.items()))]
           for p in sorted(personen.liste, key=lambda p: norm(p.nachname)) if p.varianten()],
          [26, 30, 110], wrap="BC")
    wb.save(pfad)


# ---------------------------------------------------------------- Hauptprogramm
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--kontingente", required=True)
    ap.add_argument("--bedarf", required=True)
    ap.add_argument("--feinplaner", required=True)
    ap.add_argument("--ausgabe", default="ausgabe")
    ap.add_argument("--stichtag", help="TT.MM.JJJJ – Termine mit Anfang davor werden nicht angeboten (Standard: heute)")
    ap.add_argument("--frist", default=FRIST_STANDARD, help="TT.MM.JJJJ im Satz „Anspruch verfällt ab …“")
    ap.add_argument("--manuell", help="CSV mit manuellen Zuordnungen (Abt.;Name;Vorname;Führungskraft;Feinplaner)")
    args = ap.parse_args()
    args.stichtag = dt.datetime.strptime(args.stichtag, "%d.%m.%Y").date() if args.stichtag else dt.date.today()
    frist = dt.datetime.strptime(args.frist, "%d.%m.%Y").date()

    personen = Personen()
    plaetze = lies_kontingente(args.kontingente, args.stichtag)
    alle_bedarfe = lies_bedarf(args.bedarf)
    for b in alle_bedarfe:      # Vorgesetzte zuerst anlegen: ihre Schreibweise wird angezeigt
        if b.vg:
            personen.add(b.vg, b.vg, 1, f"Kundenbedarf Z. {b.zeile}", b.vg_vorname)
    fp_zeilen = lies_feinplaner(args.feinplaner, personen)
    manuell = lies_manuell(args.manuell)

    plaetze_je_ps = defaultdict(list)
    for p in plaetze:
        if not p.grund:
            plaetze_je_ps[p.ps].append(p)
    passend = [b for b in alle_bedarfe if b.ps in plaetze_je_ps]
    nicht_offen = [b for b in passend if not b.status.upper().startswith("O")]
    bedarfe = [b for b in passend if b.status.upper().startswith("O")]
    ergebnisse = zuordnen(bedarfe, alle_bedarfe, fp_zeilen, personen, manuell)
    emails = baue_emails(bedarfe, ergebnisse, plaetze_je_ps, frist)
    gruppen = klaerfall_gruppen(bedarfe, ergebnisse)

    # Kennzahlen und Hinweise
    in_mails = [b for b, e in zip(bedarfe, ergebnisse) if e.status != "Klärfall"]
    klaer = [b for b, e in zip(bedarfe, ergebnisse) if e.status == "Klärfall"]
    plaetze_bedarf = []
    for ps, pl in sorted(plaetze_je_ps.items(), key=lambda kv: (min(p.anfang for p in kv[1]), kv[1][0].titel)):
        termine = "; ".join(f"{tn}: {fmt_dt(min(p.anfang for p in pl if p.terminnr == tn))[:10]} "
                            f"({sum(p.terminnr == tn for p in pl)} Pl.)" for tn in eindeutig(p.terminnr for p in pl))
        r = {"ps": ps, "titel": pl[0].titel, "plaetze": len(pl), "termine": termine,
             "eindeutig": sum(b.ps == ps for b in in_mails), "klaer": sum(b.ps == ps for b in klaer)}
        r["gesamt"] = r["eindeutig"] + r["klaer"]
        plaetze_bedarf.append(r)

    hinweise = [f"Frist im Fristsatz: {args.frist} (aus der Vorlage übernommen) – bitte bestätigen, dass sie auch "
                "für diese Kontingente gilt."]
    ohne = [r for r in plaetze_bedarf if r["gesamt"] == 0]
    if ohne:
        hinweise.append("Termine ohne passenden Bedarf: " + ", ".join(f"{r['titel']} ({r['ps']})" for r in ohne) + ".")
    if nicht_offen:
        hinweise.append(f"{len(nicht_offen)} Bedarf(e) mit Status ungleich „Offen“ nicht berücksichtigt (Zeilen "
                        + ", ".join(str(b.zeile) for b in nicht_offen) + ").")
    sonder = Counter(b.dringlichkeit or "leer" for b in bedarfe if not re.match(r"^[123]\b", b.dringlichkeit))
    if sonder:
        hinweise.append("Dringlichkeit ohne Stufe 1–3 wie in der Quelle übernommen: "
                        + ", ".join(f"„{k}“ {v}×" for k, v in sonder.items()) + ".")
    bq = Counter(f"{b.titel} ({b.ps})" for b in alle_bedarfe if "-BQ-" in b.produkt.upper()
                 or b.titel.lower().startswith("basisqualifikation"))
    if bq:
        hinweise.append("Bedarfe zu Basisqualifikationen (A-BQ-…) sind ein anderes Produkt als das Basisseminar und "
                        "wurden nicht zugeordnet: " + ", ".join(f"{k} {v}×" for k, v in sorted(bq.items())) + ".")
    gleich = [p for p in personen.liste if p.varianten()]
    if gleich:
        hinweise.append("Als gleiche Person gewertet (Schreibweisen): " + "; ".join(
            f"{p.anzeige} = " + " / ".join(f"„{s}“" for s in p.varianten()) for p in gleich) + ".")

    stats = {"in_mails": len(in_mails), "klaer": len(klaer),
             "ma_in_mails": len(eindeutig(b.ma for b in in_mails)), "ma_klaer": len(eindeutig(b.ma for b in klaer)),
             "plaetze_ok": sum(len(v) for v in plaetze_je_ps.values()),
             "plaetze_raus": sum(1 for p in plaetze if p.grund),
             "plaetze_bedarf": plaetze_bedarf, "hinweise": hinweise}

    aus = Path(args.ausgabe)
    aus.mkdir(parents=True, exist_ok=True)
    (aus / "Basisseminare_E-Mails.html").write_text(seite_html(emails, gruppen, plaetze, stats, args), encoding="utf-8")
    schreibe_excel(aus / "Basisseminare_Pruefliste.xlsx", emails, gruppen, plaetze, bedarfe, ergebnisse, personen, stats)
    json_out = {
        "stichtag": f"{args.stichtag:%d.%m.%Y}", "frist": args.frist,
        "emails": [{"an": em.an.anzeige, "cc": [p.anzeige for p in em.cc], "betreff": em.betreff,
                    "hinweise": em.hinweise, "text": email_text(em, args.frist),
                    "seminare": [{"ps": ps, "titel": titel, "plaetze": [p.nr for p in pl],
                                  "mitarbeiter": [{"zeile": b.zeile, "name": b.name, "vorname": b.vorname,
                                                   "abt": b.abt, "dringlichkeit": b.dringlichkeit}
                                                  for b in em.bedarfe[ps]]} for ps, titel, pl in em.seminare]}
                   for em in emails],
        "klaerfaelle": [{"gruppe": g["titel"], "grund": g["grund"], "vorschlaege": g["vorschlaege"],
                         "indiz": g["indiz"],
                         "bedarfe": [{"zeile": b.zeile, "name": b.name, "vorname": b.vorname, "abt": b.abt,
                                      "ps": b.ps, "titel": b.titel} for b in g["bedarfe"]]} for g in gruppen],
        "nicht_angebotene_plaetze": [{"nr": p.nr, "zeile": p.zeile, "ps": p.ps, "titel": p.titel, "grund": p.grund}
                                     for p in plaetze if p.grund],
        "hinweise": hinweise,
    }
    (aus / "ergebnis.json").write_text(json.dumps(json_out, ensure_ascii=False, indent=1), encoding="utf-8")

    print(f"{len(emails)} E-Mails, {len(in_mails)} Bedarfe zugeordnet, {len(klaer)} Bedarfe als Klärfall "
          f"({len(gruppen)} Gruppen), {stats['plaetze_raus']} Kontingentzeile(n) ausgeschlossen -> {aus}")


if __name__ == "__main__":
    main()
