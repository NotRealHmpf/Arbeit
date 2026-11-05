# UVV-Suche in Outlook

Das Skript `UVV-Suche.ps1` durchsucht das eigene Postfach und das Gruppenpostfach im Outlook auf dem
Firmenlaptop nach Unterlagen zur UVV-Unterweisung. Es speichert alle Dokument-Anhänge, liest ihren Text
(auch eingescannte PDFs und Fotos) und zeigt pro Mitarbeiter, in welchen Dokumenten und Mails der Name vorkommt.

- Läuft direkt auf dem Laptop, ohne Installation und ohne Admin-Rechte.
- Liest nur. In Outlook wird nichts verändert, verschoben, gelöscht oder gesendet.
- Alle Daten bleiben auf dem Laptop. Es wird nichts ins Internet geschickt, auch nicht an Claude.

## Ergebnis

Ordner `C:\Users\<Name>\UVV-Suche\<Datum_Uhrzeit>\`:

| Datei / Ordner | Inhalt |
|---|---|
| `Bericht.html` | öffnet sich am Ende automatisch: Übersicht, ein Abschnitt pro Mitarbeiter, Sammeldokumente, UVV-Unterlagen ohne erkannten Namen |
| `Zusammenfassung.csv` | eine Zeile pro Mitarbeiter (öffnet in Excel) |
| `Treffer.csv` | jede Fundstelle: eine Zeile pro Mitarbeiter und Dokument, mit der Zeile aus dem Dokument |
| `Alle_Dokumente.csv` | Verzeichnis aller gespeicherten Anhänge |
| `Treffer_nach_Mitarbeiter\` | Kopien der Treffer mit UVV-Bezug, ein Ordner pro Mitarbeiter; Sammeldokumente beginnen mit `Sammel_` |
| `UVV_ohne_erkannten_Namen\` | Dokumente mit UVV-Begriff, aber ohne erkannten Namen (z. B. handschriftlich ausgefüllte Listen) |
| `Anhaenge\` | alle gespeicherten Dokument-Anhänge (doppelte nur einmal) |
| `Protokoll.txt` | Ablauf und Fehler |

## Sammeldokumente: was ein Treffer bedeutet

Ein Treffer heißt nur, dass der Name in einem Dokument oder einer Mail steht. Ob die Person wirklich unterwiesen
wurde, kann das Skript nicht entscheiden. Deshalb zeigt es zu jedem Treffer, wie eindeutig er ist:

- **Einzel oder Sammel**: Stehen mehrere Mitarbeiter der Namensliste im selben Dokument, ist es ein
  *Sammeldokument* (im Bericht gelb). Typisch sind Teilnehmerlisten oder Fristen-/Statuslisten wie die Liste mit
  „Abgelaufen“. Auch bei einem Einzeltreffer kann das Dokument weitere Personen enthalten, die nicht in der
  Namensliste stehen.
- **Zeile im Dokument**: die Zeile, in der der Name steht. Bei Tabellen und Scans ist das die ganze Tabellenzeile,
  z. B. `Klatt, André | V.IW-… | Peter Kremer | 10.12.2024 | 09.12.2025 | Abgelaufen`. Datumsangaben aus dieser
  Zeile stehen zusätzlich unter „Datum in der Zeile“. Excel-Datumszellen werden als Datum ausgegeben.
- **Dokumentart (geschätzt)**: „Teilnahme-/Unterschriftenliste?“ bei Wörtern wie Unterschrift oder Teilnehmer,
  „Fristen-/Statusliste?“ bei Wörtern wie abgelaufen, fällig oder gültig bis. Das ist nur ein Hinweis.
- Der Abschnitt **Sammeldokumente** im Bericht führt jedes Sammeldokument nur einmal auf, mit allen Mitarbeitern
  und ihren Zeilen. So lässt sich jedes Dokument in einem Durchgang prüfen: Hat die Person teilgenommen und
  unterschrieben?
- Unterschriften und Handschrift erkennt das Skript nicht, nur gedruckten Text.

**Bewertung**

| Bewertung | Bedeutung |
|---|---|
| UVV-Bezug + Name | Name und ein UVV-Begriff (UVV, Unterweisung, Belehrung, Teilnehmerliste …) im Dokument oder in der zugehörigen Mail |
| nur Name | Name gefunden, aber kein UVV-Begriff (im Bericht eingeklappt) |
| UVV-Bezug, kein Name erkannt | UVV-Unterlage ohne erkannten Namen, bei Bedarf von Hand ansehen |

**Namenstreffer**: *Vor- und Nachname* (direkt nebeneinander, auch in Tabellen oder Mailadressen) ·
*Name* (Vor- und Nachname getrennt im Dokument, „Klatt, A.“ oder „Herr Klatt“) ·
*nur Nachname* (nur im Betreff oder Dateinamen einer UVV-Mail). Nachnamen, die auch normale Wörter sind
(z. B. Haus, Heller, Nickel), können falsche Treffer der Stufe *Name* ergeben. Die Zeile im Dokument zeigt das.

## Voraussetzungen

- Windows 10 oder 11 mit dem **klassischen Outlook**. Mit „Neues Outlook“ geht es nicht; dann oben rechts den
  Schalter „Neues Outlook“ ausschalten.
- Das Gruppenpostfach steht links in der Outlook-Ordnerliste. Falls nicht: Mailadresse oben im Skript bei
  `$ZusaetzlichePostfaecher` eintragen.
- Windows PowerShell 5.1 (auf jedem Windows 10/11 vorhanden), für die Texterkennung von PDFs und Bildern.

## Schritt für Schritt

1. **Skript auf den Laptop holen**: `UVV-Suche.ps1` auf GitHub öffnen und über „Download raw file“ speichern.
   Oder den Inhalt kopieren, in den Editor einfügen und mit „Speichern unter“ als `UVV-Suche.ps1` speichern
   (Dateityp „Alle Dateien“). Am besten in einen eigenen Ordner, z. B. `C:\Users\<Name>\UVV-Skript\`.
2. **Namensliste**: im selben Ordner `namen.txt` anlegen, eine Person pro Zeile im Format `Nachname, Vorname`.
   Zeilen können direkt aus der Excel- oder Webliste kopiert werden; es zählt nur die erste Spalte.
   Fehlt die Datei, legt das Skript sie an und öffnet den Editor.
3. **Outlook öffnen.**
4. **Starten**: Rechtsklick auf `UVV-Suche.ps1` → **Mit PowerShell ausführen**. Fragt PowerShell nach der
   „Ausführungsrichtlinie“, `J` und Enter eingeben (gilt nur für dieses Fenster).
5. **Fragen beantworten**: welche Postfächer (Enter = alle), ab welchem Datum, Modus (1 = gründlich mit
   Texterkennung, 2 = schnell).
6. **Warten**: Der Fortschritt wird angezeigt. Strg+C bricht ab; der Bericht wird trotzdem mit dem bisherigen Stand
   erstellt. Am Ende öffnet sich `Bericht.html`.

Empfehlung: erst ein kurzer Testlauf (z. B. Mails der letzten vier Wochen), dann der vollständige Lauf.
Das Startdatum sollte vor dem ältesten Unterweisungsdatum der Liste liegen, damit auch neuere Nachweise
für lange abgelaufene Unterweisungen gefunden werden.

## Wenn etwas nicht klappt

| Meldung / Verhalten | Ursache und Abhilfe |
|---|---|
| Outlook fragt, ob ein Programm zugreifen darf | Zugriff erlauben. Die Abfrage kann sich alle 10 Minuten wiederholen. |
| „Outlook konnte nicht geöffnet werden“ | Das Skript braucht das klassische Outlook, nicht „Neues Outlook“. |
| Fenster schließt sofort, Meldung zur Ausführungsrichtlinie oder „eingeschränkter Modus“ | Skripte sind von der IT gesperrt. Dann die Suche von Hand machen (unten) oder die IT fragen. |
| „Outlook zeigt nur Mails ab …“ | Outlook hat ältere Mails nicht auf dem Laptop (Einstellung „E-Mail herunterladen für: 12 Monate“). Datei → Kontoeinstellungen → Kontoeinstellungen → Konto doppelklicken → Regler auf „Alle“, Outlook neu starten, warten bis alles geladen ist. Beim Gruppenpostfach zusätzlich unter „Weitere Einstellungen → Erweitert“ den Haken „Freigegebene Ordner herunterladen“ entfernen. Oder die älteren Mails in Outlook im Web suchen. |
| „Texterkennung nicht verfügbar“ | PDFs und Bilder werden dann nur über Dateiname und Mailtext gefunden. Word- und Excel-Dateien werden trotzdem gelesen. |
| Dauert sehr lange | Die Texterkennung braucht etwa eine Sekunde pro Seite. Passende und neueste Dokumente werden zuerst gelesen. Bei großen Postfächern am besten abends laufen lassen; Outlook ist währenddessen langsamer. |
| Verschlüsselte Mails | werden nicht geöffnet (sonst fragt Outlook nach Karte/PIN); nur der Betreff wird geprüft. |

## Ohne Skript: Suche in Outlook im Web

1. `outlook.office.com` öffnen.
2. Gruppenpostfach: oben rechts auf das Profilbild → „Weiteres Postfach öffnen“ → Adresse eingeben.
   Es öffnet sich in einem neuen Tab; dort suchen.
3. Suchen, z. B. `UVV Klatt` oder `Unterweisung`, dann über die Filter „Hat Anlagen“ und den Zeitraum eingrenzen.
4. Outlook im Web durchsucht auch den Text von PDF- und Word-Anhängen, aber nicht von eingescannten Dokumenten.

## Einstellungen

Oben im Skript (ohne Umlaute bearbeiten): `$NamenDatei`, `$AusgabeHauptordner`, `$StandardSeit`,
`$ZusaetzlichePostfaecher`, `$Suchbegriffe`, `$AlleAnhaengeBehalten`, `$OcrMaxSeiten`, `$MaxAnhangMB`, `$MinBildKB`.

## Datenschutz

Namensliste, Bericht und gespeicherte Anhänge enthalten Personaldaten von Kolleginnen und Kollegen.
Sie bleiben auf dem Laptop und gehören nicht in öffentliche Ordner. Nach Abschluss den Ergebnisordner löschen.
Inhalte daraus nur dann in Claude oder andere KI-Dienste geben, wenn IT bzw. Datenschutz das freigegeben haben.
`namen.txt` wird nicht ins Repository eingecheckt (`.gitignore`).

## Stand

Die Syntax ist geprüft, die Namens- und Zeilenerkennung wurde mit Beispieltexten nachgestellt.
Auf Windows mit Outlook konnte das Skript noch nicht ausprobiert werden, deshalb zuerst den Testlauf machen.
