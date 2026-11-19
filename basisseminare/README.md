# Basisseminare – E-Mails an die Führungskräfte

Erzeugt aus Excel-Listen je Führungskraft (Bezirksleiter) einen E-Mail-Entwurf im Aufbau der
Vorlage („Basisseminar SIMIS LC“): Termine aus den Kontingenten und die betroffenen Mitarbeiter.
Fallen unter eine Führungskraft mehrere Basisseminare, gibt es trotzdem nur eine E-Mail, mit einem
Block je Seminar.

| Eingabe | Inhalt |
|---|---|
| Kontingente (z. B. „Kontingente Netz Frankfurt“) | Termine/Plätze, eine Zeile je Platz (TNP) |
| Kundenbedarf mit Vorgesetzten | Bedarfe je Mitarbeiter mit OE (Abt.) und Vorgesetztem |
| Feinplaner-/Bezirksleiter-Liste | OE (LS) → Bezirksleiter und Feinplaner (FP) |
| optional: Mitarbeiter zu Feinplaner (UVV-Übersicht), `--uvv` | Blatt „… Übersicht“ mit allen Mitarbeitern und verantwortlicher Führungskraft, dazu ein Blatt je Feinplaner (Blattname = Feinplaner, auch Kurzform wie „Tobi“) |

## Regeln

1. **Termine:** nur Kategorie „BS Basisseminar“, Anfang ab Stichtag (Standard: heute), Platz noch frei.
2. **Bedarfe:** Status „Offen“ und PS-Nr. eines angebotenen Termins. Basisqualifikationen (A-BQ-…) sind ein eigenes Produkt und zählen nicht.
3. **Führungskraft (An):**
   - **Ohne `--uvv`:** Der Vorgesetzte aus dem Kundenbedarf muss in der Feinplaner-Liste als Bezirksleiter der OE des
     Mitarbeiters stehen. Verglichen wird die OE-Endung (V.IW-MI-N-FFM-IL 05 ≙ I.IA-MI-N-FFM-IL 05).
   - **Mit `--uvv`:** Die UVV-Übersicht gilt als aktuellster Stand je Mitarbeiter. Die Führungskraft ist eindeutig,
     wenn zwei Quellen übereinstimmen: Kundenbedarf und Übersicht, oder Übersicht und Feinplaner-Liste.
   - **Wechsel:** Nennt ein Feinplaner-Blatt für eine ganze OE eine andere Führungskraft als die Übersicht
     (z. B. Bemerkung „übernimmt ab TT.MM.JJJJ“), gilt das Blatt für alle Mitarbeiter der OE.
   - **Schreibweisen:** Kleine Tippfehler in Namen (ein Buchstabe, z. B. Meier/Maier) werden toleriert und angezeigt.
     Mitarbeiter werden nur bei exakt gleicher Schreibweise zugeordnet (Niklas ≠ Niclas).
4. **Feinplaner (Cc):**
   - Steht der Mitarbeiter auf genau einem Feinplaner-Blatt, ist das sein Feinplaner.
   - Sonst gilt die Feinplaner-Liste für OE und Führungskraft (normale Zeilen vor Zeilen mit Einsatzort „Bezirksleiter“).
   - Hat ein Feinplaner ein eigenes Blatt, auf dem kein Mitarbeiter der OE steht, obwohl laut Übersicht welche im
     Zeitraum des Blatts fällig sind, betreut er die OE offenbar nicht und scheidet aus (mit Prüfhinweis).
   - Mehrere Kandidaten werden über die Übersicht eingegrenzt: Blätter anderer Mitarbeiter derselben OE oder
     Bemerkungen wie „kommt noch laut <Vorname des Feinplaners>“.
   - Was mehrdeutig bleibt, kommt komplett in Cc und bekommt einen Prüfhinweis.
   - Je E-Mail zeigt „Cc-Herkunft“, woher jeder Feinplaner stammt.
5. **Klärfall:** Die Quellen widersprechen sich, oder der Mitarbeiter ist in der Übersicht mit „raus“ markiert. Dann gibt
   es keine E-Mail, nur Vorschläge (jeweils mit Feinplaner). Bemerkungen wie „dauerkrank“ erscheinen als Prüfhinweis.

## Ausgabe

| Datei | Inhalt |
|---|---|
| `Basisseminare_E-Mails.html` | Entwürfe mit Kopier-Buttons (An, Cc, Betreff, E-Mail-Text samt Tabellen für Outlook), Cc-Herkunft, Klärfälle, Hinweise, Plätze vs. Bedarf |
| `Basisseminare_Pruefliste.xlsx` | Zuordnung je Bedarf mit Begründung und Herkunft des Cc, Klärfälle, Termine, Plätze vs. Bedarf, Namensabgleich |
| `ergebnis.json` | dieselben Ergebnisse maschinenlesbar |

Die Signatur fügt Outlook selbst ein; der kopierte Text endet mit „Viele Grüße“.
Prüfhinweise (gelbe Kästen) sind nicht Teil der E-Mail.

## Aufruf

```bash
pip install openpyxl
python erstelle_emails.py --kontingente "Kontingente Netz Frankfurt.xlsx" --bedarf "Kundenbedarf.xlsx" --feinplaner "Feinplaner.xlsx" --uvv "Mitarbeiter zu Feinplaner Übersicht.xlsx" --frist 31.01.2027 --ausgabe ausgabe
```

- `--frist`: Datum im Satz „Bitte beachte, dass der Anspruch auf diese Plätze ab dem … verfällt.“
- `--stichtag TT.MM.JJJJ`: Termine mit früherem Anfang werden nicht angeboten (Standard: heute)
- `--uvv`: optionale Übersicht „Mitarbeiter zu Feinplaner“ (siehe oben)
- `--manuell zuordnung.csv`: Klärfälle nach Rückmeldung zuordnen. Dann neu erzeugen.

`zuordnung.csv` (Trennzeichen „;“, UTF-8). Ohne Name gilt die Zeile für alle Klärfälle der OE; ohne Feinplaner
wird er wie oben ermittelt; mehrere Feinplaner mit „/“ trennen:

```
Abt.;Name;Vorname;Führungskraft;Feinplaner
V.IW-MI-N-FFM-XX 01;;;Vorname Nachname;Vorname Nachname
;NACHNAME;Vorname;Vorname Nachname;
```

Die Eingabe- und Ausgabedateien enthalten Personaldaten und werden nicht eingecheckt (`.gitignore`).
