# Ausbildungsstand-Dashboard (NFP – Gesamtübersicht aller Bezirke)

Erzeugt aus der Datei „NFP – Gesamtübersicht aller Bezirke LST-intern.xlsx“ eine Kopie
mit Excel-Dashboard. Die Original-Blätter bleiben unverändert; neu hinzu kommen:

| Blatt | Inhalt |
|---|---|
| Übersicht | Kennzahlen, „Alle Bezirke nach Status“, „Status je Bezirk“, Tabelle mit Links zu den Bezirksseiten |
| Grafik &lt;Bezirk&gt; (12×) | Kennzahlen, „Mitarbeiter nach Status“, „Status nach Ziel-Qualifikation“, Zahlentabelle, Mitarbeiterliste |
| Daten | Auswertung je Mitarbeiter (per Formel, z. B. als Quelle für Power BI) |
| Einstellungen | Bezugsjahr, Status-Kategorien, Zuordnung Qualifikation → Spalte J–Q |

Alles rechnet mit Formeln. Ändert man Werte in den Bezirks-Blättern, passen sich Zahlen
und Diagramme automatisch an.

## Status je Mitarbeiter

Ziel-Qualifikation (Spalte F) → zugehörige Spalte J–Q (Zuordnung im Blatt Einstellungen):

1. **Fertig**: „x“ in der Ziel-Spalte (oder Ist-Qualifikation = Ziel-Qualifikation)
2. **Abschluss &lt;Jahr&gt;**: Jahreszahl in der Ziel-Spalte (27 → 2027), einzeln je Jahr ab dem Bezugsjahr, das letzte Jahr als „ab …“
3. **Überfällig**: das Jahr liegt vor dem Bezugsjahr, aber es steht noch kein „x“
4. **Fehlt**: weder „x“ noch Jahr, also keine Ausbildung geplant
5. **Ziel nicht eingetragen**: Spalte F ist leer

Das Blatt „Bezirksleiter FBÜW“ wird nicht ausgewertet.

## Neu erzeugen

```bash
pip install openpyxl lxml        # zusätzlich LibreOffice mit Calc (soffice)
python build_dashboard.py "<original>.xlsx" "<ausgabe>.xlsx"
```

- `build_sheets.py`: baut die neuen Blätter mit openpyxl
- `merge_into_original.py`: setzt sie direkt in das .xlsx-Paket des Originals ein, damit Kommentare, Diagramme und SharePoint-Metadaten erhalten bleiben
- `inject_cache.py`: speichert die mit LibreOffice berechneten Werte in Zellen und Diagrammen, damit auch die geschützte Ansicht die Zahlen zeigt

Die Excel-Dateien enthalten Personaldaten und werden nicht eingecheckt (`.gitignore`).
