# Personal- und Einsatzplanung 2027 (Excel, vollautomatisch)

Erzeugt aus der bestehenden PEP-Datei bzw. dem Bereitschaftskalender 2027 eine Excel-Arbeitsmappe für
Personal-, Schicht-, Bereitschafts-, Urlaubs-, Qualifikations- und Ressourcenplanung. Alle Ergebnisse sind
Formeln – Änderungen in den Eingabeblättern wirken sofort auf Plan, Prüfungen, Zeitkonten, Statistiken,
Dashboard und Druckansichten.

Übernommen aus den Quelldateien: Teamaufstellung (Meister, AV, TVT Fahrer, Handwerker 1–3 je Team),
Teamfarben (Blau 4F81BD, Rot FF4B4B, Gelb FFFF00, Grün 92D050, Orange F79646), Kalenderlayout mit KW,
Feiertags-/Wochenend-/Ferienfarben, PEP-Raster (Namen × Tage, Kopf B7DEE8, Wochenende orange) und die
PEP-Legende (Urlaub 00B050, JAZ C4D79B, Unterricht FF0000, FDL FFC000, Kabel 7030A0, Nacht FFFF00 …).

## Blätter

| Gruppe | Blatt | Inhalt |
|---|---|---|
| Start | Start | Navigation, Legende, Annahmen |
| Eingabe | Einstellungen | Jahr, Stichtag, Rotation (fest), Dienstzeiten, Sollbesetzung, Ampel, Feiertage Hessen (automatisch), Schulferien, Unterrichte, Pflichtqualifikationen je Funktion, Listen |
| | Mitarbeiterstamm (Blatt 1) | Stammdaten; Resturlaub, Überstunden, JAZ berechnet |
| | Team- und Funktionsmatrix (Blatt 2) | Teamfunktion + beliebig viele JA/NEIN-Funktionen und Unterrichte |
| | Qualifikationen (Blatt 3) | Nachweise, Ablaufdatum, Restlaufzeit, Ampel, geplanter Unterricht |
| | Urlaub und Abwesenheiten (Blatt 4) | Urlaub, Krankheit, JAZ, Unterricht, Fortbildung, Freistellung, Sonderurlaub |
| | Vertretungen | Ersatzsuche (Top 5) und bestätigte Vertretungen |
| | Zeitbuchungen | Vorträge, Auszahlungen, Korrekturen (Überstunden, JAZ, Urlaub) |
| Planung | Schichtplan 2027 (Blatt 5) | Datum, Wochentag, KW, Feiertag, Nachtteam, freigestelltes Team, Tagdienstteams, Bemerkung … |
| | PEP 2027 | Raster Mitarbeiter × Tag mit Codes T/N/V/F und Abwesenheiten |
| | Einsatzfähigkeitsprüfung (Blatt 6) | je Team und Tag Soll/Ist je Funktion, Status ✅/❌, fehlende Funktionen, Ersatzvorschläge |
| | Überstunden und JAZ (Blatt 7) | Soll/Ist, Über-/Minderstunden, Monats-/Jahreswert, Salden, JAZ-Konto |
| | Personalbedarfsplanung | Soll/Ist, Verfügbarkeit, Ausfallreserve, Unter-/Überbesetzung – Tag, Woche, Monat, Jahr, Struktur |
| Auswertung | Dashboard, Statistiken, Dienstverteilung, Urlaubs-/Krankenstatistik, Monats-/Jahresbericht, Einsatzfähigkeitsbericht | |
| Druck | Druck Jahr / Monat / Quartal / Team / Bereitschaft / Qualifikation | DIN A4/A3 quer, Monat/Quartal/Team wählbar |
| Technik | Hilfsberechnung (ausgeblendet) | sortierte Mitarbeiterliste, Qualifikations-Gültigkeit, Abwesenheits-Priorität je Tag, Monatsstunden |

## Fachlogik

- **Rotation** ab Rotationsbeginn (01.01.2027, Freitag) in fester Folge: Fr–So Rot · Mo–Do Gelb · Fr–So Grün ·
  Mo–Do Orange · Fr–So Blau · Mo–Do Rot · Fr–So Gelb · Mo–Do Grün · Fr–So Orange · Mo–Do Blau, fortlaufend
  (Block-Nr = Halbwochen seit Rotationsbeginn, Team = Folge[Block mod 10]).
- **Freistellung**: nach Mo–Do Freitag, nach Fr–So Montag + Dienstag (= Team des vorherigen Blocks).
- **Tagdienst**: Mo–Fr ohne Feiertag alle Teams außer Nacht- und freigestelltem Team. Ganze Tage, keine Stunden.
- **Abwesenheiten** überschreiben den Dienst; bei Überschneidung gewinnt die höhere Priorität
  (Krankheit > Kabel > FDL > Unterricht > Fortbildung > Urlaub > JAZ > Sonderurlaub > Freistellung).
- **Einsatzfähigkeit** je Team und Tag: Mitarbeiter mit passender Teamfunktion, anwesend und mit gültigen
  Pflichtqualifikationen (Einstellungen) + bestätigte Vertretungen ≥ Soll (1 M, 1 AV, 1 TVT, 3 HW).
  Meldungen: „Meister fehlt“, „AV fehlt“, „TVT fehlt“, „Zu wenig Handwerker (x/3)“, „Qualifikation abgelaufen (…)“.
- **Ersatz** für das Nachtteam: P1 gleiches Team (Reserve/Überdeckung) → P2 anderes Team gleiche Funktion →
  P3 anderes Team mit Mehrfachfunktion; nur anwesend, qualifiziert, nicht freigestellt, nicht doppelt.
  Innerhalb der Priorität rotieren die Vorschläge blockweise gleichmäßig über alle Kandidaten.
- **Qualifikationen**: Ablauf = letzte Schulung (oder durchgeführter Unterricht) + Gültigkeit; ein vor dem
  Ablauf geplanter Unterricht verlängert die Gültigkeit für die Planung. Ampel > 180 / 90–180 / < 90 / abgelaufen.
- **Zeitkonten**: Soll = Arbeitstage × 8 Std.; Ist = T × Std. Tagdienst + N/V × angerechnete Bereitschaft (60 %)
  + Abwesenheiten an Arbeitstagen × Soll; Freistellung nach Nacht 0 Std.; JAZ-Tage belasten das JAZ-Konto.
- **Ampel**: Grün Soll erreicht · Gelb kritisch (Reserve < Puffer bzw. Lücke mit Ersatz) · Rot Unterbesetzung.

Alle Annahmen (Pause, Anrechnung, Sollstunden, Mindestbesetzung, Pflichtqualifikationen, Gültigkeitsdauern)
stehen im Blatt Einstellungen und auf der Startseite.

## Neu erzeugen

```bash
pip install openpyxl lxml        # zusätzlich LibreOffice mit Calc (soffice)
python build_personalplanung.py --quelle "<PEP_2026_OL_FH-FG.xlsx>" --out Personalplanung_2027.xlsx \
    --overrides seed_overrides.json            # optional
python build_personalplanung.py ... --demo --out Personalplanung_2027_DEMO.xlsx   # mit Testdaten
python verify_personalplanung.py --datei <berechnete Datei> --quelle "<Bereitschaftskalender_2027.xlsx>"
```

- `pp_source.py` liest Teamaufstellung, Kalenderfarben und Feiertage aus der Quelldatei
- `pp_input.py`, `pp_plan.py`, `pp_reports.py`, `pp_print.py` bauen die Blätter, `pp_base.py` enthält
  Konstanten/Farben, `pp_demo.py` die Testdaten
- `build_personalplanung.py` rechnet eine Kopie mit LibreOffice durch, bricht bei jedem Formelfehler ab und
  speichert die Ergebnisse als Cache in der Datei (`inject_cache.py`)
- `verify_personalplanung.py` rechnet Rotation, Freistellung, Feiertage, PEP-Raster, Einsatzfähigkeit,
  Ersatzvorschläge, Statistiken, Personalbedarf und Zeitkonten unabhängig in Python nach und vergleicht
  außerdem die Rotation mit den Teamfarben des Original-Bereitschaftskalenders

`seed_overrides.json` (Format siehe `seed_overrides.example.json`) korrigiert Teamzuordnungen und ergänzt
Mehrfachfunktionen gegenüber der Quelldatei. Excel-Dateien und die Overrides enthalten Personaldaten und
werden nicht eingecheckt (`.gitignore`).
