# Ausbildungsstand-Dashboard LST (Bedarf_Bestand_LST + Zielzustand)

Erzeugt aus „Bedarf_Bestand_LST.xlsx“ (Mitarbeiter je Bezirk) und „LST_Zielzustand_gesamt.xlsx“ (Soll)
eine Kopie mit Excel-Dashboard. Die Original-Blätter bleiben unverändert; neu hinzu kommen:

| Blatt | Inhalt |
|---|---|
| Übersicht | Kennzahlen, „Alle Bezirke nach Status“, „Status je Bezirk“, Tabelle mit Links zu den Bezirksseiten |
| Soll-Ist | Fehlende Stellen je Bezirk, LST gesamt je Stufe (Verlauf + Zielzustand), Verlauf je Bezirk gegen Soll, Tabellen |
| Dashboard &lt;Bezirk&gt; (12×) | Kennzahlen, Status, Soll/heute/nach Plan je Stufe, Verlauf gestapelt + Zielzustand, kleine Verlaufsdiagramme je Stufe mit Soll-Linie, Zahlentabellen, Abgleich mit dem Soll, Mitarbeiterliste |
| Prüfliste | auffällige Einträge je Mitarbeiter (per Formel) |
| Zielzustand | Kopie von „LST_Zielzustand_gesamt.xlsx“, Tabelle1 – hier das Soll ändern |
| Einstellungen | Bezugsjahr, Jahr für „davon“, Status, Stufen, Zuordnung Blatt → Zielzustand, Zuordnung Qualifikation → Spalte |
| Daten (ganz hinten) | Auswertung je Mitarbeiter (per Formel) |

Alles rechnet mit Formeln (begrenzte Bereiche, keine ganzen Spalten). Ändert man Werte in den Bezirks-Blättern,
im Blatt Zielzustand oder das Bezugsjahr, passen sich Zahlen und Diagramme automatisch an. Druck: A4 quer.

## Spalten in den Bezirks-Blättern

B Name, C Vorname, G Ziel-Qualifikation, J Ist-Qualifikation, L Azubi, M Arb LST, N Weichenmechaniker,
O Signalmechaniker, P Signalmechaniker RBEG, S Teamleiter, U/V örtliche Verwendungsprüfung Wmech/SigMech.
„x“ = fertig, Zahl = geplant fertig im Jahr (27 = 2027). Alles andere (z. B. 0) zählt wie leer und steht in der Prüfliste.
Kopfzeile ist Zeile 1 oder 2; die Überschriften werden beim Erzeugen und im Blatt Einstellungen geprüft.

## Soll-Ist und Verlauf (nur Spalten L–P)

- Stufen von links nach rechts: Azubi (L) → Arb LST (M) → Wmech (N) → SigMech (O) → SigMech RBEG (P).
- Jeder Mitarbeiter zählt genau einmal, auf der rechtesten erreichten Stufe. „x“ in U = Wmech erreicht, „x“ in V = SigMech erreicht.
- Heute: rechteste Spalte mit „x“. Ende eines Jahres: rechteste Spalte mit „x“ oder Jahr ≤ diesem Jahr.
  Nach Plan: alle eingetragenen Jahre – auch über die Ziel-Qualifikation hinaus.
- Teamleiter zusätzlich über Spalte S (ein Teamleiter mit „x“ in P zählt auch bei SigMech RBEG – in Tabellen und Diagrammen).
- Ziel- (G) und Ist-Qualifikation (J) werden dafür nicht verwendet.
- Soll = Spalte „Zielzustand“ je Bezirk (Zeilen Arbeiter LST, Wmech, SigMech, SigMech RBEG, Teamleiter).
- Je Stufe und Zeitpunkt: „Soll erreicht“, „es fehlen X“ oder „X mehr als Soll“. Fehlende Stellen = Summe über die Stufen.
- Verlauf: 7 Jahre ab dem Bezugsjahr (wandert mit).

## Status je Mitarbeiter (gemessen an der Ziel-Qualifikation, Spalte G)

1. **Fertig**: „x“ in der Ziel-Spalte; bei Wmech/SigMech reicht auch „x“ bei der örtlichen Verwendungsprüfung.
2. **Abschluss &lt;Jahr&gt;**: Jahr in der Ziel-Spalte (Bezugsjahr … Bezugsjahr+3, danach „ab“).
3. **Überfällig**: das Jahr ist vorbei, aber es steht kein „x“.
4. **Fehlt (nichts geplant)**: weder „x“ noch Jahr.
5. **Ziel unklar**: Spalte G leer oder das Ziel hat keine eigene Spalte (z. B. Senior Expert LST).

## Neu erzeugen und prüfen

```bash
pip install openpyxl lxml                 # zusätzlich LibreOffice Calc: apt-get install -y libreoffice-calc
python build_dashboard.py "<Bedarf_Bestand_LST>.xlsx" "<LST_Zielzustand_gesamt>.xlsx" "<ausgabe>.xlsx"
KEEP_CALC=calc.xlsx python build_dashboard.py ...     # berechnete Kopie für die Prüfung aufheben
python verify.py calc.xlsx "<Bedarf_Bestand_LST>.xlsx" "<LST_Zielzustand_gesamt>.xlsx" [Bezugsjahr] [--liste]
python test_dynamic.py "<ausgabe>.xlsx"   # ändert eine Testkopie und prüft, dass alles nachrechnet
```

- `build_sheets.py`: baut die neuen Blätter mit openpyxl (Diagramme an Zellbereiche gebunden, Schrift Aptos Narrow)
- `merge_into_original.py`: setzt sie direkt in das .xlsx-Paket des Originals ein (Kommentare, externe Verknüpfungen,
  Diagramme und Metadaten bleiben erhalten; die Original-Datei wird nie mit openpyxl gespeichert)
- `inject_cache.py`: speichert die mit LibreOffice berechneten Werte in Zellen und Diagrammen
- `reference.py`: unabhängige Nachrechnung in Python; `verify.py` vergleicht sie Zelle für Zelle mit dem Dashboard

Die Excel-Dateien enthalten Personaldaten und werden nicht eingecheckt (`.gitignore`).
