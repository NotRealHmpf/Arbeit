# Fahrtenbuch nach dem Musterfahrtenbuch (DIN A6-ähnliches Format)

Erzeugt ein Fahrtenbuch mit den 15 Spalten des Musterfahrtenbuchs – von „Lfd. Nr.“ bis
„Unterschrift Fahrer“, mit Übertrag-Zeile und den Abkürzungsvorschlägen – als Excel-Datei und
als Druckvorlage.

| Datei | Inhalt |
|---|---|
| `Fahrtenbuch.xlsx` | Deckblatt (Fahrzeugdaten, berechtigte Fahrer, Jahresübersicht), 12 Monatsblätter mit je 3 Druckseiten à 20 Fahrten (A4 quer), Hinweise, Abkürzungen, Beispiel |
| `Fahrtenbuch_Druckvorlage.pdf` | Zum handschriftlichen Führen: Deckblatt, Hinweise, Abkürzungen, Ausfüllbeispiel und 48 nummerierte Fahrtenbuchseiten (A4 quer) |

## Excel-Version

- **Eingaben:** auf dem Deckblatt die gelben Felder (Jahr, Fahrzeugdaten, km-Stand bei Beginn,
  Fahrer), in den Monatsblättern die weißen Felder. Grau hinterlegte Felder rechnen.
- **km-Stand bei Abfahrt** = letzte Ankunft, auch über Seiten- und Monatsgrenzen. Weicht der
  Tacho ab, trägt man den abgelesenen Wert ein; die Zelle wird orange und rechts neben der
  Tabelle steht ein Hinweis (Poolfahrzeugdisponent informieren).
- **Gefahrene km** = Ankunft − Abfahrt, verteilt nach der Fahrtart (Spalte P, wird nicht gedruckt):
  „WO/AO“ bzw. „AO/WO“ im Zweck (auch mit DR/Ber) → WO/AO · „PRIVAT“ in Fahrweg oder Zweck →
  privat · „zurück“ → wie die vorherige Fahrt · sonst → gesch. Die Fahrtart lässt sich je Zeile
  von Hand wählen.
- **Auswahllisten:** Zweck der Fahrt (Blatt „Abkürzungen“, erweiterbar) und Fahrer (Deckblatt);
  freie Eingaben bleiben möglich.
- **Auswertung:** Übertrag, Monatssummen und Jahresübersicht (km gesch./WO/AO/privat, Anteile,
  getankte Liter) rechnen sich automatisch.

Eine Excel-Datei lässt sich nachträglich unbemerkt ändern und erfüllt deshalb nicht die
steuerlichen Anforderungen an ein ordnungsgemäßes Fahrtenbuch. Dafür die Druckvorlage vor Gebrauch
fest binden und handschriftlich führen.

## Neu erzeugen

```bash
pip install openpyxl lxml        # zusätzlich LibreOffice mit Calc (soffice)
python build_fahrtenbuch.py --jahr 2027 --seiten 48
```

- `--jahr`: Jahr auf dem Deckblatt (Standard: aktuelles Jahr; lässt sich in der Datei ändern)
- `--seiten`: Anzahl Fahrtenbuchseiten in der Druckvorlage
- `--ausgabe`: Zielordner (Standard: `ausgabe/`)

LibreOffice rechnet die Datei beim Erzeugen durch. Das Skript bricht ab, wenn eine Formel einen
Fehler liefert oder das Blatt „Beispiel“ nicht die Werte des Musters ergibt (818 km gesch.,
178 km WO/AO, 23 km privat, 76,0 l getankt). Die berechneten Werte werden in der Datei
gespeichert, damit auch die geschützte Ansicht sie zeigt.

Der Ordner `ausgabe/` wird nicht eingecheckt; ausgefüllte Fahrtenbücher enthalten Personendaten.
