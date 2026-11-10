# PEP 2027 – Personal- und Einsatzplanung (Excel mit Makros)

Erzeugt aus der Datei mit den Blättern „PEP_2026“ (PEP-Layout) und „Bereitschaft 2027“ (Bereitschaftskalender,
Teamzuordnung Ma1 … Ma28) eine vollständig verknüpfte Arbeitsmappe `PEP_2027.xlsm` für Schicht-, Bereitschafts-,
Urlaubs-, Abwesenheits-, Ressourcen- und Einsatzplanung.

## Blätter

| Blatt | Inhalt |
|---|---|
| Start | Navigation, Status (Makros, Prüfhinweise), Kurzanleitung, Legende |
| Dashboard | Kennzahlen am Stichtag/Monat/Jahr, Ampeln 🟢🟡🔴, Teamstatus, 4 Diagramme, 14-Tage-Ausblick |
| Mitarbeiter (Blatt 1) | Personalnummer, Name, Team (Dropdown), Funktion, Bemerkung |
| Funktionsmatrix (Blatt 2) | JA/NEIN je Qualifikation (14 Spalten), Mehrfachfunktionen |
| Abwesenheiten (Blatt 3) | zentrale Datenquelle: Mitarbeiter, Von, Bis, Art, Bemerkung (+ Prüfspalten) |
| Jahreskalender (Blatt 4) | PEP-Layout: Mitarbeiter × 366 Tage, Kopfzeilen Datum/Wochentag/KW/Feiertag/Nachtteam/Freigestelltes Team/Tagdienstteams/Bemerkung |
| Schichtplanung (Blatt 5) | je Tag Nacht-/Bereitschaftsteam, Freistellung (Vorschlag/wirksam), Wochenend- und Feiertagsrotation |
| Einsatzfähigkeit (Blatt 6) | tägliche Prüfung (1 Meister, 1 AV, 1 TVT, 2 Handwerker), Fehlermeldungen, Ersatzvorschläge |
| Ressourcen | Zusätzliche MA und AZUBIS (keine Abwesenheiten, ersetzen keine Sollbesetzung) |
| Freistellungen | manuelle Anpassungen (haben immer Vorrang) |
| Vertretungen | übernommene Ersatzvorschläge |
| Personalbedarf | Soll/Ist/Verfügbarkeit/Ausfallreserve/Unter-/Überbesetzung je Tag, Woche, Monat, Jahr |
| Statistik, Monatsbericht, Jahresbericht, Urlaub-Krank, Teamübersicht | Auswertungen inkl. Dienst- und Nachtdienstverteilung, Einsatzfähigkeitsbericht |
| Druck Jahr / Monat / Team / Bereitschaft | Druckansichten (A3 bzw. A4 quer, auf eine Seite skaliert) |
| Einstellungen | Planungsjahr, Rotationsstart und -folge, Sollbesetzung, Kürzel/Legende, Feiertage Hessen |
| Berechnung (ausgeblendet) | Statuscode je Mitarbeiter und Tag – Grundlage aller Ansichten |

## Fachliche Logik

- **Rotation**: ab Rotationsstart (01.01.2027, Freitag) im Wechsel Fr–So / Mo–Do mit der festen Folge
  Rot, Gelb, Grün, Orange, Blau (10 Blöcke, Zyklus 5 Wochen). Geprüft gegen alle 365 eingefärbten Tage des
  Bereitschaftskalenders. Die Folge steht im Blatt Einstellungen und wird nie automatisch verändert.
- **Freistellung (Vorschlag)**: nach Mo–Do → Freitag, nach Fr–So → Montag + Dienstag (nur Arbeitstage).
  Anpassungen im Blatt Freistellungen („entfällt am“ / „stattdessen frei am“, je Team oder Mitarbeiter)
  haben Vorrang.
- **Priorität je Tag**: Krankheit > Unterricht > FDL > Kabel > Freistellung > Urlaub eTime > Urlaub > JAZ >
  Besonderheit > Bereitschaft > Ressource. Urlaub an einem Freistellungstag verbraucht damit keinen Urlaubstag.
- **Einsatzfähigkeit**: Verfügbare Mitglieder des Bereitschaftsteams je Stammfunktion + eingetragene Vertretungen.
  Fehlt eine Funktion, darf ein Mitglied mit passender Zusatzqualifikation einspringen, wenn seine eigene
  Funktion überbesetzt ist (🟡 kritisch). Sonst ❌ / 🔴 mit Ersatzvorschlag: gleiches Team → anderes Team
  (Stammfunktion) → Mehrfachfunktion; bei Gleichstand weniger Bereitschaftstage zuerst.
- **Personalbedarf**: Ist = Teammitglieder − Abwesende − Freigestellte (Arbeitstage); Soll und Mindestreserve
  sind Annahmen im Blatt Einstellungen (20 bzw. 3) und müssen an den echten Bedarf angepasst werden.
- **Feiertage Hessen**: per Osterformel für jedes Jahr; Heiligabend, Silvester, Oster-/Pfingstsonntag optional.

## Synchronisierung Kalender ↔ Tabellen

Das Kalenderraster enthält ausschließlich Formeln und ist immer geschützt (Kennwort `pep`). Mit aktivierten
Makros öffnet ein **Doppelklick** auf einen Tag bzw. ein **Rechtsklick** auf markierte Tage eine Eingabe
(U, UE, K, JAZ, S, FDL, Kabel, FS, -, V, Z, A, leer = löschen, Freitext = Besonderheit). Das Makro schreibt den
Eintrag in die zuständige Tabelle (Abwesenheiten, Freistellungen, Vertretungen, Ressourcen), verlängert
angrenzende Zeiträume bzw. teilt sie auf; der Kalender zeigt das Ergebnis über seine Formeln. Es gibt damit nur
eine Datenquelle und nie abweichende Stände – auch ohne Makros, in Excel Online und mit AutoSave (es wird beim
Speichern nichts umgeschaltet). Ohne Makros erfolgt die Eingabe in den Tabellen.

Das VBA-Projekt (`xl/vbaProject.bin`) wird ohne Excel aus den Quellen in `pep/vba/` erzeugt
(`pep/vbaproject.py`: MS-OVBA-Kompression, dir-/PROJECT-Streams, Compound-File-Writer; der P-Code-Cache bleibt
leer, Excel kompiliert beim Öffnen). Sollte Excel das Projekt nicht laden, lassen sich die Module mit
`--vba-export` ausgeben und von Hand importieren (`modPEP.bas`, `modKonfig.bas`; Inhalt der `.txt`-Dateien in
DieseArbeitsmappe bzw. die Blattmodule Jahreskalender/Mitarbeiter einfügen).

## Neu erzeugen

```bash
pip install openpyxl lxml        # zusätzlich LibreOffice (soffice) für Durchrechnung und Prüfung
python build_pep.py "<original>.xlsx" "<ausgabe>.xlsm" --namen namen.csv
python tests/pruefung.py "<original>.xlsx" --namen namen.csv
```

- `namen.csv` (optional, Semikolon): `Platzhalter;Name;Personalnummer`, z. B. `Ma1;Nachname;`. Die Datei enthält
  Personaldaten und wird nicht eingecheckt (`.gitignore`), ebenso alle Excel-Dateien.
- Ablauf: openpyxl baut alle Blätter → LibreOffice rechnet eine Kopie durch → Prüfung auf Formelfehler →
  Ergebnisse werden als gespeicherte Werte übernommen (Vorschau/geschützte Ansicht) → VBA-Projekt einbetten.
- `tests/pruefung.py` prüft Rotation, Feiertage, Arbeitstage, Freistellung inkl. Anpassung, Prioritäten,
  Einsatzfähigkeit mit interner Vertretung und den Jahreswechsel auf 2028 (Schaltjahr).
- Kapazität des Kalenders: 50 Zeilen (Teamköpfe + Mitarbeiter), änderbar über `CAP` in `pep/model.py`.
- Benötigt Excel 2019 oder Microsoft 365 (MAXIFS, TEXTJOIN).
