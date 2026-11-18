# Sicherungsplanheft – wer fehlt?

`abgleich.py` gleicht die Mitarbeiterliste mit den Sicherungsplanheft-Exporten aus dem
Lernportal ab und schreibt eine einfache Excel-Liste: Personennummer, Name, Status
(„Fehlt“ oder „Erledigt“), die Fehlenden oben.

```bash
python abgleich.py Mitarbeiter.xlsx Sicherungsplanheft_Abgleich.xlsx Export1.xls Export2.xls Export3.xls
```

- Wer in mindestens einem Export mit Status „Abgeschlossen“ steht, gilt als erledigt.
- Namen werden ohne Rücksicht auf Groß-/Kleinschreibung, Akzente und Umlaut-Schreibweise
  verglichen („Nachname, Vorname“ in der Liste, „Vorname Nachname“ im Export). Ein Name in
  Klammern („Nachname, Vorname (Zweitname)“) gilt als zweiter Nachname.
- Eine Personennummer, die mehrfach in der Mitarbeiterliste steht, zählt einmal.
- Abweichende Schreibweisen und doppelte Personennummern stehen als Kommentar am Namen.
- Namen aus den Exporten, die keinem Mitarbeiter zugeordnet werden können, kommen ins
  Blatt „Nicht zugeordnet“.

Benötigt `openpyxl`. Die Excel-Dateien enthalten Personaldaten und werden nicht eingecheckt.
