# Kontingente je Standort (EL-Produkte)

Erzeugt aus der PowerPoint „Kontingentierung EL Produkte“ eine Excel-Liste für einen
Standort (Standard: FFM). Jeder Kontingent-Platz bekommt eine eigene Zeile, in die der
Mitarbeiter eingetragen wird, der den Platz wahrnimmt. Beispiel: SpDr S 60 vom
23.08.–03.09.2027 u. 13.09.–24.09.2027 mit 4 Plätzen FFM → 4 Zeilen (Platz 1 / 4 … 4 / 4).

| Blatt | Inhalt |
|---|---|
| Kontingente &lt;Standort&gt; | Kennzahlen (Plätze gesamt, eingeplant, offen, Belegung) und je Platz eine Zeile: Seminar, Termin 1, Termin 2, Platz, **Mitarbeiter** (gelb, zum Ausfüllen), Status, Bemerkung |
| Übersicht | Je Termin: Plätze, eingeplant, offen, Status (offen / teilweise / komplett); darunter Termine ohne Kontingent für den Standort (z. B. „Kontingentierung Zentrale“, „Termin folgt“) |

- Status, Kennzahlen und Übersicht rechnen mit Formeln und passen sich beim Eintragen an.
- Derselbe Name zweimal im selben Termin wird rot markiert.
- Grundlagencoachings werden nicht berücksichtigt.

## Neu erzeugen

```bash
pip install python-pptx openpyxl lxml     # zusätzlich LibreOffice mit Calc (soffice)
python build_kontingente.py "<Kontingentierung>.pptx" "<ausgabe>.xlsx" [FFM|FMZ|KO|KSL|RIS|P3]
```

Die Formelergebnisse werden wie beim Dashboard mit LibreOffice berechnet und in der
Datei gespeichert (`../dashboard/inject_cache.py`), damit auch die geschützte Ansicht
die Zahlen zeigt. Die Excel-Dateien werden nicht eingecheckt (`.gitignore`).
