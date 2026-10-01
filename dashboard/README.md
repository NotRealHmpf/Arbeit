# Ausbildungsstand-Dashboard LST (Bedarf_Bestand_LST + Zielzustand)

Erzeugt aus „Bedarf_Bestand_LST.xlsx“ (Mitarbeiter je Bezirk) und „LST_Zielzustand_gesamt.xlsx“ (Soll)
eine Kopie mit Excel-Dashboard. Die Original-Blätter bleiben unverändert; neu hinzu kommen:

| Blatt | Inhalt |
|---|---|
| Übersicht | Kennzahlen, „Alle Bezirke nach Status“, „Status je Bezirk“, Tabelle mit Links zu den Bezirksseiten |
| Soll-Ist | Fachkräfte je Bezirk (Soll / heute / nach Plan), Verlauf LST gesamt, Verlauf je Bezirk (12 kleine Diagramme, gleiche Skala), Tabellen je Bezirk und Qualifikation |
| Grafik &lt;Bezirk&gt; (12×) | Kennzahlen, Status-Diagramme, Soll-Ist je Qualifikation, Verlauf gegenüber Soll (gesamt und je Qualifikation), Zahlen, Mitarbeiterliste |
| Daten | Auswertung je Mitarbeiter (per Formel) |
| Zielzustand | Kopie von „LST_Zielzustand_gesamt.xlsx“, Tabelle1 – hier das Soll ändern |
| Einstellungen | Bezugsjahr, Status-Kategorien, Zuordnung Bezirks-Blatt → Zielzustand, Zuordnung Qualifikation → Spalte |

Alles rechnet mit Formeln. Ändert man Werte in den Bezirks-Blättern oder im Blatt Zielzustand, passen sich
Zahlen und Diagramme automatisch an.

## Spalten in den Bezirks-Blättern

G = Ziel-Qualifikation, J = Ist-Qualifikation, L–S = Stand je Qualifikation (Azubi … Teamleiter),
U/V = örtliche Verwendungsprüfung Wmech/SigMech. „x“ = erledigt, Jahreszahl (z. B. 27) = geplant.

## Status je Mitarbeiter (bezogen auf die Ziel-Qualifikation)

1. **Fertig**: „x“ in der Ziel-Spalte, bei Wmech/SigMech auch „x“ bei der örtlichen Verwendungsprüfung.
   Ist-Qualifikation = Ziel-Qualifikation zählt nur als fertig, wenn in der Ziel-Spalte gar nichts steht.
2. **Abschluss &lt;Jahr&gt;**: geplantes Jahr (das frühere aus Ziel-Spalte und Verwendungsprüfung), einzeln je Jahr ab dem Bezugsjahr
3. **Überfällig**: das Jahr liegt vor dem Bezugsjahr, aber es steht noch kein „x“
4. **Fehlt (nichts geplant)**: weder „x“ noch Jahr
5. **Ziel nicht eingetragen**: Spalte G ist leer

## Soll-Ist und Verlauf

- Jeder Mitarbeiter zählt genau einmal – auf der Stufe, die er erreicht hat. Die Stufen laufen von links nach rechts:
  Azubi (L) → Arbeiter LST (M) → Wmech (N) → SigMech (O) → SigMech RBEG (P) → Teamleiter (S).
  Wer SigMech RBEG wird, zählt nicht mehr als SigMech.
- Heute = rechteste Spalte mit „x“ (Wmech/SigMech auch „x“ bei der Verwendungsprüfung U/V); ohne „x“ zählt die Ist-Qualifikation.
- Verlauf: Zum Jahresende zählt die rechteste Spalte mit „x“ oder einem geplanten Jahr bis dahin. Standard: höchstens bis zur
  Ziel-Qualifikation (Spalte G); mit Einstellungen!C5 = „Ja“ zählen auch Planjahre darüber hinaus.
  Überfällige Jahre zählen ab dem Bezugsjahr. „Nach Plan“ = alle eingetragenen Jahre erreicht.
- Soll = Spalte „Zielzustand“ je Bezirk.
- Passend besetzt = Ziel-Stellen, auf denen jemand mit genau dieser Qualifikation sitzt (je Stufe höchstens so viele wie im Soll).
  Fehlende Ziel-Stellen = Soll − passend besetzt. Mehr als Soll = Mitarbeiter auf Stufen, die schon voll sind.
- Diagramm „Mitarbeiter je Stufe“: gestapelt je Jahr, rechts daneben der Zielzustand.

## Neu erzeugen

```bash
pip install openpyxl lxml        # zusätzlich LibreOffice mit Calc (Paket libreoffice-calc)
python build_dashboard.py "<Bedarf_Bestand_LST>.xlsx" "<LST_Zielzustand_gesamt>.xlsx" "<ausgabe>.xlsx"
```

- `build_sheets.py`: baut die neuen Blätter mit openpyxl (Diagramme sind an Zellbereiche gebunden)
- `merge_into_original.py`: setzt sie direkt in das .xlsx-Paket des Originals ein, damit Kommentare, Diagramme,
  externe Verknüpfungen und Metadaten erhalten bleiben
- `inject_cache.py`: speichert die mit LibreOffice berechneten Werte in Zellen und Diagrammen, damit auch die
  geschützte Ansicht die Zahlen zeigt

Die Excel-Dateien enthalten Personaldaten und werden nicht eingecheckt (`.gitignore`).
