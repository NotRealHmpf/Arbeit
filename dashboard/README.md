# Ausbildungsstand-Dashboard LST (Bedarf_Bestand_LST + Zielzustand)

Erzeugt aus „Bedarf_Bestand_LST.xlsx“ (Mitarbeiter je Bezirk) und „LST_Zielzustand_gesamt.xlsx“ (Soll)
eine Kopie mit Excel-Dashboard. Die Original-Blätter bleiben unverändert; neu hinzu kommen:

| Blatt | Inhalt |
|---|---|
| Übersicht | Kennzahlen, „Alle Bezirke nach Status“, „Status je Bezirk“, Tabelle je Bezirk (Klick öffnet die Bezirksseite), Status-Legende |
| Soll-Ist | Kennzahlen, LST gesamt je Stufe mit Zielzustand-Säule, fehlende Stellen je Bezirk, Verlauf je Bezirk (12 kleine Diagramme, gleiche Skala), Tabellen |
| Prüfliste | auffällige Einträge (Jahr vorbei ohne „x“, Ziel ohne passende Planung, Ziel unklar, ungültige Einträge, Ist ≠ L–P) |
| Bezirk &lt;Name&gt; (12×) | Kennzahlen, Status-Diagramm, Soll/heute/nach Plan je Stufe, Verlauf als gestapelte Säulen mit Zielzustand, kleine Verlaufsdiagramme je Stufe mit gestrichelter Soll-Linie, Zahlentabellen, Mitarbeiterliste |
| Daten | Auswertung je Mitarbeiter (per Formel) |
| Zielzustand | Kopie von „LST_Zielzustand_gesamt.xlsx“, Tabelle1 – hier das Soll ändern |
| Einstellungen | Bezugsjahr (= aktuelles Jahr), Jahr für „In Ausbildung, davon …“, Status, Zuordnung Bezirks-Blatt → Zielzustand, Zuordnung Qualifikation → Spalte |

Alles rechnet mit Formeln (ohne Formeln über ganze Spalten). Ändert man Werte in den Bezirks-Blättern, im Blatt
Zielzustand oder das Bezugsjahr, passen sich Zahlen und Diagramme automatisch an. Alle Blätter drucken auf A4 quer.

## Spalten in den Bezirks-Blättern

Kopfzeile in Zeile 2 (Hanau: Zeile 1), gelesen werden die Zeilen 2–61; die Überschriften werden beim Erzeugen geprüft.
B Name, C Vorname, G Ziel-Qualifikation, J Ist-Qualifikation, L Azubi, M Arb LST, N Weichenmechaniker,
O Signalmechaniker, P Signalmechaniker RBEG, S Teamleiter, U/V örtl. Verwendungsprüfung Wmech/SigMech.
„x“ = fertig, Zahl (z. B. 27) = geplant fertig im Jahr 2027.

## Soll-Ist und Verlauf (nur Spalten L–P, nicht G/J)

- Stufen von links nach rechts: Azubi (L) → Arb LST (M) → Wmech (N) → SigMech (O) → SigMech RBEG (P).
- Jeder Mitarbeiter zählt genau einmal – auf der rechtesten erreichten Stufe. Wer RBEG ist, ist kein SigMech mehr.
- Heute: rechteste Spalte mit „x“; „x“ in U zählt als Wmech erreicht, „x“ in V als SigMech erreicht.
- Ende eines Jahres: rechteste Spalte mit „x“ oder einer Jahreszahl ≤ diesem Jahr (alle Planjahre, auch über das Ziel hinaus).
  „Nach Plan“ = alle eingetragenen Jahre. Der Verlauf zeigt das Bezugsjahr und die 6 folgenden Jahre.
- Teamleiter zählen zusätzlich über Spalte S – auch in den gestapelten Säulen (z. B. bei RBEG und als Teamleiter).
- Soll = Spalte „Zielzustand“ (Zeilen Arbeiter LST, Wmech, SigMech, SigMech RBEG, Teamleiter). Je Stufe:
  „Soll erreicht“, „es fehlen X“ oder „X mehr als Soll“. Fehlende Stellen = Summe der Lücken über die Stufen
  (LST gesamt: Summe über die Bezirke).

## Status je Mitarbeiter (gemessen an der Ziel-Qualifikation G)

1. **Fertig**: „x“ in der Spalte der Ziel-Qualifikation; bei Wmech/SigMech reicht auch „x“ bei der örtlichen Verwendungsprüfung.
2. **Abschluss &lt;Jahr&gt;**: Jahreszahl in dieser Spalte (Bezugsjahr … ab Bezugsjahr + 4).
3. **Überfällig**: Das Jahr ist vorbei, aber es steht kein „x“.
4. **Fehlt (nichts geplant)**: In der Spalte steht nichts.
5. **Ziel unklar**: Spalte G leer oder ohne Spalte in L–S (z. B. „Senior Expert LST“; Zuordnung in Einstellungen ergänzbar).

## Neu erzeugen

```bash
pip install openpyxl lxml                 # zusätzlich LibreOffice mit Calc: apt-get install -y libreoffice-calc
python build_dashboard.py "<Bedarf_Bestand_LST>.xlsx" "<LST_Zielzustand_gesamt>.xlsx" "<ausgabe>.xlsx"
```

- `build_sheets.py`: baut die neuen Blätter mit openpyxl (Diagramme an Zellbereiche gebunden, Schrift Aptos Narrow)
- `merge_into_original.py`: setzt sie direkt in das .xlsx-Paket des Originals ein – die Originaldatei wird nie mit
  openpyxl gespeichert, Kommentare, externe Verknüpfungen und Metadaten bleiben erhalten
- LibreOffice rechnet eine Kopie durch (ohne die externen SharePoint-Verknüpfungen, die nur in H/I benutzt werden)
- `reference.py` / `verify.py`: unabhängige Nachrechnung in Python; jeder Wert des Dashboards wird verglichen,
  bei Abweichungen bricht das Erzeugen ab
- `inject_cache.py`: speichert die berechneten Werte in Zellen und Diagrammen, damit auch die geschützte Ansicht die
  Zahlen zeigt (Excel rechnet beim Öffnen ohnehin neu)

Die Excel-Dateien enthalten Personaldaten und werden nicht eingecheckt (`.gitignore`).
