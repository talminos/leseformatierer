# Leseformatierer

Web-Tool, das Word-Dokumente (``.docx``) automatisch für das Vorlesen aufbereitet:

- **schwarz/fett** für betonte Wörter,
- <span style="color:#ff0000">**rot**</span> für zusätzliche Hilfs- bzw. Bindewörter,
- <span style="color:#0000ff">**blau + fett**</span> für Hauptanker:
  bestehende Triggerwörter (Sprecher-/Regie-Markierungen) bleiben
  **unverändert**, bei reinem Text werden zusätzlich wenige Hauptanker
  automatisch erzeugt.

Die Anwendung läuft komplett auf dem Server, **ohne externe KI-API**:
Die Formatierungslogik basiert auf [`python-docx`](https://python-docx.readthedocs.io)
und einer transparenten, deterministischen Heuristik (siehe
[`formatter.py`](formatter.py)).

---

## Inhalt

| Datei                  | Zweck                                                       |
| ---------------------- | ----------------------------------------------------------- |
| `app.py`               | Flask-Anwendung (Upload, Validierung, Job-ID, Cleanup)      |
| `formatter.py`         | Eigentliche Formatierungslogik                              |
| `templates/index.html` | Deutsche Oberfläche mit Upload-Formular                     |
| `requirements.txt`     | Python-Abhängigkeiten                                       |
| `Procfile`             | Startbefehl für Render / Heroku-kompatible Plattformen      |
| `create_test_docx.py`  | Erzeugt eine kleine Beispieldatei `test_input.docx`         |
| `.gitignore`           | Entwicklungs-/Laufzeitdateien außerhalb der Versionierung   |

---

## Lokale Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
flask --app app run
```

Standardmäßig läuft die App dann auf <http://127.0.0.1:5000>.

Alternativ ohne Flask-CLI:

```bash
python app.py
```

### Test-Dokument erzeugen und prüfen

```bash
python create_test_docx.py
python -c "from formatter import format_document; \
  print(format_document('test_input.docx', 'test_output.docx', mode='loose'))"
```

`test_output.docx` lässt sich anschließend in Word/LibreOffice öffnen.

---

## Deployment auf Render

1. Repository auf Render verknüpfen (`Web Service`, Runtime: **Python 3**).
2. **Build Command**:
   ```
   pip install -r requirements.txt
   ```
3. **Start Command**:
   ```
   gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 120
   ```
4. **Environment Variables**:
   - `SECRET_KEY` – beliebiger, langer Zufallswert (Pflicht für Produktion).
   - `PYTHON_VERSION` (optional) – z. B. `3.11.9`.

Da der Dienst zustandslos arbeitet (Dateien werden nach jedem Request entfernt),
genügt ein einziger Web-Service ohne Persistent Disk. Der mitgelieferte
`Procfile` ist für Render und Heroku gleichermaßen kompatibel.

> Hinweis: Reine Static-Hoster wie Netlify reichen nicht – die Anwendung
> braucht einen Python-fähigen Application-Server.

---

## Bedienung

1. `.docx`-Datei auswählen (max. 10 MB).
2. **Modus** wählen:
   - *Beispielnah*: Wörter ab **3** Buchstaben werden formatiert.
   - *Streng*: Nur Wörter ab **4** Buchstaben werden formatiert.
3. Optionen setzen:
   - **Vorhandene rote Wörter behalten** – bestehendes Rot bleibt unangetastet.
   - **Nur Absätze mit blau/fetten Triggerwörtern bearbeiten** – nützlich,
     wenn nur Kapitel mit Sprecher-Markierungen umformatiert werden sollen.
   - **Manuskriptformat anwenden** – setzt Papierformat 30 cm breit × 22 cm hoch,
     umlaufende Ränder 0,7 cm, 1,15 Zeilenabstand und 12 pt Abstand nach
     jedem Absatz, wie in der Wunschdatei.
   - **Redezeilen erzeugen** – längere Absätze werden an Gedankenstrichen
     und Satzenden in kurze Lesezeilen gegliedert. Je 1–2 Zeilen bilden einen
     Block: innerhalb des Blocks wird ein **weicher Zeilenumbruch** gesetzt,
     zwischen den Blöcken entsteht ein **harter Absatz** mit Zeilenabstand 1,15
     und Abstand nach Absatz 12 pt.
4. Auf *Formatieren & herunterladen* klicken. Die fertige Datei heißt
   `<originalname>_formatiert.docx`.

---

## Formatierungslogik in Kürze

Pro Absatz:

1. Absatz wird in **Sätze** zerlegt (Satzende: `.`, `!`, `?` oder Absatzende).
2. Jeder Satz wird in **Tokens** (Wörter, Whitespace, Satzzeichen) zerlegt.
   Bestehende Run-Eigenschaften (Farbe, Fett, Italic, …) werden mitgeführt.
3. **Gesperrte** Tokens werden niemals überschrieben:
   - blau + fett (Triggerwörter, z. B. „PAUSE", „D.")
   - bereits rote Wörter, falls die Option aktiv ist
   - grün + fett (Kommentare/offene Fragen)
4. Aus den verbleibenden Wörtern werden Kandidaten gewählt (Mindestlänge nach
   Modus, 1–2-Buchstaben-Wörter normalerweise nicht).
5. Pro Satz werden ungefähr **22–28 %** der geeigneten Kandidaten
   **blau/fett**, **36–44 %** **rot**, **36–44 %** **schwarz/fett**,
   der Rest bleibt normal. Bei sehr kurzen Sätzen wird deutlich sparsamer
   markiert.
6. Heuristik-Vorlieben:
   - Nomen, Eigennamen, Großschreibungen und lange Schlüsselbegriffe werden
     bevorzugt **blau/fett**.
   - Jahreszahlen, ausgeschriebene Datumsangaben wie `30. März 2026`,
     Altersangaben wie `13 Jahren` und bevorzugte Namen werden als starke
     blaue Trigger behandelt.
   - Blaue Hauptanker werden möglichst nicht direkt nebeneinander gesetzt.
   - Eine anpassbare Liste bevorzugter Hauptanker kann Kundennamen und wichtige
     Begriffe wie `Luise` oder `Flügel` gezielt stärker gewichten.
   - Satzanfang bevorzugt **schwarz/fett**.
   - Wörter direkt nach einem Gedankenstrich (`-`, `–`, `—`) bevorzugt
     rot oder schwarz/fett.
   - Direkt nach einem blau-fetten Trigger keine Ballung.
   - Zwei rote Wörter direkt hintereinander werden nach Möglichkeit
     vermieden.
   - Wörter unter der gewählten Mindestlänge bleiben in der Regel normal.
   - Sehr kurze Sätze (≤ 4 Kandidaten) erhalten nur 1–2 Zusatzformatierungen.

Konstanten und Funktionen (`is_blue_bold_trigger`, `is_red`,
`split_into_sentences`, `classify_candidates`, `rebuild_paragraph`) stehen
oben in `formatter.py` und sind kommentiert.

Erkanntes Blau (Hex): `0000FF`, `0070C0`, `2F5496`, `002060`.

Sternchen-Einschübe wie

```text
*****************************************************
Intro: Andreas Gabalier…
*****************************************************
```

werden von der normalen Wortmarkierung ausgenommen und mit einfachem
Zeilenabstand, fett und mittig formatiert. Die Rubrikzeile, z. B. `Intro:`,
wird auf 16 pt gesetzt.

---

## Datenschutz

- Hochgeladene Dateien liegen nur kurz in `uploads/` und werden direkt nach
  der Verarbeitung gelöscht.
- Die fertige Datei wird nach erfolgreichem Download serverseitig gelöscht
  (`@after_this_request`). Zusätzlich entfernt eine Cleanup-Funktion
  verwaiste Dateien älter als 1 Stunde.
- Es werden **keine externen Dienste** (insbesondere keine KI-APIs) angesprochen.

---

## Bekannte Einschränkungen

- Die optionale Redezeilen-Funktion erzeugt kurze Lesezeilen anhand von
  Wortanzahl, Gedankenstrichen und Satzenden. Je 1–2 Zeilen werden in einem
  harten Absatz gebündelt; innerhalb dieses Absatzes werden weiche
  Zeilenumbrüche gesetzt. Die tatsächliche Zeilenzahl kann je nach
  Word-Seitenbreite und Schrift minimal abweichen.
- **Kopf- und Fußzeilen** werden bewusst nicht angefasst. Wer sie ebenfalls
  formatieren möchte, müsste `app.py`/`formatter.py` um eine Section/Header-
  Iteration erweitern.
- **Felder, Kommentare, Fußnoten, Endnoten und Textfelder** bleiben
  unverändert. python-docx greift hier nur eingeschränkt durch.
- **Listenformatierungen** (Aufzählungszeichen, Nummerierung) und Absatz-
  Stile bleiben erhalten; einzelne Run-Eigenschaften können beim Neuaufbau
  nur dann übernommen werden, wenn sie auf Run-Ebene gesetzt waren.
- `.docm`-Dateien (Word mit Makros) werden nicht direkt unterstützt – die
  Datei muss als `.docx` ohne Makros gespeichert werden.
- Die Heuristik ist absichtlich **nicht KI-basiert**. Sie liefert konsistente,
  reproduzierbare Ergebnisse, kann aber nicht „verstehen", welches Wort
  inhaltlich am wichtigsten ist.

---

## Empfohlener Testablauf

1. `python create_test_docx.py` ausführen → erzeugt `test_input.docx`.
2. Lokalen Server starten: `flask --app app run`.
3. Im Browser <http://127.0.0.1:5000> öffnen, `test_input.docx` hochladen,
   Modus *Beispielnah* wählen, Optionen ausprobieren.
4. Heruntergeladene `test_input_formatiert.docx` in Word/LibreOffice prüfen:
   - Triggerwort `PAUSE` muss weiterhin **blau + fett** sein.
   - Vorhandene rote Wörter bleiben rot, falls die Option aktiv ist.
   - Bei reinem Text sollten zentrale Hauptwörter zusätzlich **blau/fett**
     werden; rote und schwarz/fette Wörter verteilen sich ähnlich wie im
     Beispielbild.
5. Größe-Limits prüfen: Eine > 10 MB große Datei muss eine deutsche
   Fehlermeldung auslösen.
