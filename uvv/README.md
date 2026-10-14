# UVV-Überwachung (Arbeiten im Gleisbereich)

Erzeugt aus „UVV Überwachung.xlsx“ eine neue, übersichtliche Arbeitsmappe für die Unterweisung
„1.3334 Unterweisung Arbeiten im Gleisbereich nach DGUV Vorschrift 78; DGUV Regel 101-024“.
Grundlage sind die Blätter „Tabelle1“ (Mitarbeiter) und „ZW-Daten (Stichtag)“ (Name in Spalte A, FK in Spalte BP).

| Blatt | Inhalt |
|---|---|
| Tabelle X | Übersicht je Mitarbeiter: Personennummer, Name, Abteilung, Führungskraft, Letzte UVV, Gültig bis, Resttage, Status. Bezeichnung, Personalakte und Bemerkung sind ausgeblendet (über „+“ einblendbar). |
| Tabelle Y | Namen und FK aus „ZW-Daten (Stichtag)“. Wird bei jedem Update neu eingefügt, Abgleich nur über den Namen. |
| Anleitung | Bedienung und Farben |

## Regeln in Tabelle X

| Farbe | Status | Bedingung |
|---|---|---|
| Lila | Nicht mehr in ZW-Daten | Name steht nicht in Tabelle Y (hat Vorrang) |
| Rot | Abgelaufen | Letzte UVV + 364 Tage liegt vor heute |
| Gelb | Bald fällig | läuft in ≤ 30 Tagen ab |
| Blau | Neu – Name festschreiben / UVV-Datum fehlt / Datum prüfen | neuer Name aus Tabelle Y bzw. kein gültiges Datum |
| Grau | UVV entfällt | in „Letzte UVV“ steht „entfällt“ |

- Gültigkeit (364) und Vorwarnung (30) stehen oben rechts in Tabelle X.
- Namensvergleich ohne Leerzeichen und Groß-/Kleinschreibung („Heß, Astrid“ = „Heß,Astrid“).
- Die Führungskraft kommt per Formel aus Tabelle Y (keine Verknüpfung mehr auf „Alle Personen .xlsx“).
- Zugänge: Unter der Tabelle steht in Spalte B eine Formel, die den nächsten Namen aus Tabelle Y anhängt,
  der weiter oben noch fehlt. Sie bezieht sich nur auf Zeilen darüber (kein Zirkelbezug) und hört nach der
  ersten leeren Formelzeile auf. Angehängte Namen werden als Werte festgeschrieben, danach trägt man den Rest
  von Hand ein. Formeln und Farben sind bis Zeile 2500 vorbereitet.
- Mitarbeiter aus den ZW-Daten, die in „Tabelle1“ fehlen, hängt das Skript beim Erzeugen gleich als Werte an
  (nur der Name).

## Neu erzeugen

```bash
pip install openpyxl lxml        # zusätzlich LibreOffice mit Calc (soffice)
python build_uvv.py "UVV Überwachung.xlsx" "UVV Übersicht.xlsx"
```

Das Skript rechnet die Datei zur Kontrolle mit LibreOffice durch, vergleicht den Status jedes Mitarbeiters mit
einer eigenen Berechnung und speichert die Ergebnisse als Werte in den Formelzellen (wie `dashboard/`, dessen
Hilfsfunktionen es mitbenutzt).

Die Excel-Dateien enthalten Personaldaten und werden nicht eingecheckt (`.gitignore`).
