# Basisseminare – E-Mails an die Führungskräfte

Erzeugt aus drei Excel-Listen je Führungskraft (Bezirksleiter) einen E-Mail-Entwurf im Aufbau der
Vorlage („Basisseminar SIMIS LC“): Termine aus den Kontingenten und die betroffenen Mitarbeiter.
Fallen unter eine Führungskraft mehrere Basisseminare, gibt es trotzdem nur eine E-Mail, mit einem
Block je Seminar.

| Eingabe | Inhalt |
|---|---|
| Kontingente (z. B. „Kontingente Netz Frankfurt“) | Termine/Plätze, eine Zeile je Platz (TNP) |
| Kundenbedarf mit Vorgesetzten | Bedarfe je Mitarbeiter mit OE (Abt.) und Vorgesetztem |
| Feinplaner-/Bezirksleiter-Liste | OE (LS) → Bezirksleiter und Feinplaner (FP) |

## Regeln

1. **Termine:** nur Kategorie „BS Basisseminar“, Anfang ab Stichtag (Standard: heute), Platz noch frei
2. **Bedarfe:** Status „Offen“ und PS-Nr. eines angebotenen Termins. Basisqualifikationen (A-BQ-…) sind ein eigenes Produkt und zählen nicht.
3. **Eindeutig:** Der Vorgesetzte aus dem Kundenbedarf steht in der Feinplaner-Liste als Bezirksleiter der OE des Mitarbeiters.
   Verglichen wird die OE-Endung (V.IW-MI-N-FFM-IL 05 ≙ I.IA-MI-N-FFM-IL 05). Kleine Tippfehler in Namen
   (ein Buchstabe, z. B. Meier/Maier) werden toleriert und in der Ausgabe angezeigt.
4. **Klärfall:** Vorgesetzter fehlt, passt nicht zur Feinplaner-Liste oder die OE fehlt dort. Dann gibt es keine
   E-Mail, nur Vorschläge (Vorgesetzter laut Kundenbedarf, Bezirksleiter laut Feinplaner-Liste, jeweils mit Feinplaner).
5. **An/Cc:** An = Bezirksleiter, Cc = Feinplaner der passenden Zeile(n). Gibt es mehrere Feinplaner, kommen alle in Cc
   und die E-Mail bekommt einen Prüfhinweis.

## Ausgabe

| Datei | Inhalt |
|---|---|
| `Basisseminare_E-Mails.html` | Entwürfe mit Kopier-Buttons (An, Cc, Betreff, E-Mail-Text samt Tabellen für Outlook), Klärfälle, Hinweise, Plätze vs. Bedarf |
| `Basisseminare_Pruefliste.xlsx` | Zuordnung je Bedarf mit Begründung, Klärfälle, Termine, Plätze vs. Bedarf, Namensabgleich |
| `ergebnis.json` | dieselben Ergebnisse maschinenlesbar |

Die Signatur fügt Outlook selbst ein; der kopierte Text endet mit „Viele Grüße“.
Prüfhinweise (gelbe Kästen) sind nicht Teil der E-Mail.

## Aufruf

```bash
pip install openpyxl
python erstelle_emails.py --kontingente "Kontingente Netz Frankfurt.xlsx" --bedarf "Kundenbedarf.xlsx" --feinplaner "Feinplaner.xlsx" --frist 31.01.2027 --ausgabe ausgabe
```

- `--frist`: Datum im Satz „Bitte beachte, dass der Anspruch auf diese Plätze ab dem … verfällt.“
- `--stichtag TT.MM.JJJJ`: Termine mit früherem Anfang werden nicht angeboten (Standard: heute)
- `--manuell zuordnung.csv`: Klärfälle nach Rückmeldung zuordnen. Dann neu erzeugen.

`zuordnung.csv` (Trennzeichen „;“, UTF-8). Ohne Name gilt die Zeile für alle Klärfälle der OE; ohne Feinplaner
wird er aus der Feinplaner-Liste gesucht; mehrere Feinplaner mit „/“ trennen:

```
Abt.;Name;Vorname;Führungskraft;Feinplaner
V.IW-MI-N-FFM-XX 01;;;Vorname Nachname;Vorname Nachname
;NACHNAME;Vorname;Vorname Nachname;
```

Die Eingabe- und Ausgabedateien enthalten Personaldaten und werden nicht eingecheckt (`.gitignore`).
