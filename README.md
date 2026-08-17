# Präsentations-Generator für Open WebUI

Ein Open-WebUI-Tool, das aus strukturierten Folien-Daten fertig gestaltete
PowerPoint-Dateien (16:9) erzeugt und einen Download-Link direkt im Chat ausgibt.

---

## 1. Installation

Getestet gegen Open WebUI 0.9.0 sowie ältere Stände (mit und ohne `path`-Spalte,
synchrone und asynchrone Modell-Schicht).

1. In Open WebUI: **Workspace → Tools → „+"** (neues Tool anlegen).
2. Den gesamten Inhalt von `praesentations_generator.py` einfügen.
3. Name (z. B. „Präsentations-Generator") und Beschreibung werden aus dem Kopf der
   Datei übernommen. **Speichern**.
4. Open WebUI installiert `python-pptx` automatisch (steht in der `requirements:`-Zeile).
   Falls der Container keinen Internetzugang hat, vorher manuell installieren:
   ```bash
   docker exec -it open-webui pip install python-pptx
   ```
5. Tool aktivieren – entweder pro Modell (**Workspace → Models → Modell bearbeiten →
   Tools**) oder pro Chat über das Werkzeug-Symbol im Eingabefeld.

### Wichtig: Native Function Calling einschalten

Das Tool bekommt ein größeres JSON-Argument. Damit das zuverlässig funktioniert:
**Workspace → Models → Modell bearbeiten → Advanced Params → Function Calling: `Native`**
(bzw. in den Admin-Einstellungen als Standard). Mit dem Default-Modus („Default" /
prompt-basiert) brechen manche Modelle lange Argumente ab.

---

## 2. Nutzung

Einfach im Chat formulieren, das Modell füllt die Folien selbst:

> Erstelle mir eine Präsentation über unsere Wärmepumpen-Strategie 2027 für den
> Aufsichtsrat, ca. 12 Folien, mit Kennzahlen und einem Zeitstrahl. Theme: midnight.

Das Tool antwortet mit einem Download-Link und hängt die Datei an die Nachricht an.

Zweite Funktion: **`list_themes`** – zeigt alle Designs, wenn du fragst „Welche
Präsentations-Designs gibt es?".

---

## 3. Designs

| Theme | Charakter |
|---|---|
| `midnight` | Marine, souverän, klassisch Business |
| `graphite` | Reduziert, hoher Kontrast, ein scharfer Akzent |
| `teal` | Frisch, technisch, vertrauenswürdig |
| `ocean` | Tiefes Blau, ruhig, seriös |
| `forest` | Natürlich, nachhaltig, geerdet |
| `terracotta` | Warm, handwerklich, menschlich |
| `coral` | Energiegeladen, jung, Marketing |
| `berry` | Edel, kulinarisch, elegant |
| `cherry` | Fast monochrom mit kräftigem Rot |

Dazu der Modus **`mode`**:

- `auto` (Standard): dunkle Titel-, Kapitel-, Zitat- und Abschlussfolien, helle
  Inhaltsfolien – die klassische „Sandwich"-Dramaturgie.
- `light`: durchgehend hell.
- `dark`: durchgehend dunkel (Premium-Look).

---

## 4. Eigene CI (Valves)

**Workspace → Tools → Tool → Zahnrad (Valves)**

| Valve | Wirkung |
|---|---|
| `default_theme` | Standard-Design, wenn das Modell keines wählt |
| `mode` | `auto` / `light` / `dark` |
| `brand_accent` | Eigene Akzentfarbe, z. B. `#C8102E` – überschreibt das Theme |
| `brand_accent2` | Zweite Akzentfarbe (Diagramme) |
| `brand_deep` | Eigene dunkle Bühnenfarbe für Titel- und Kapitelfolien |
| `font_heading` / `font_body` | Schriften, Standard: Cambria / Calibri |
| `footer_text` | Fester Fußzeilentext, z. B. Firmenname |
| `show_footer`, `show_slide_numbers` | Fußzeile / Foliennummern an oder aus |
| `logo_url` | Logo (PNG/JPG) für Titel- und Abschlussfolie |
| `base_url` | Externe Basis-URL für absolute Download-Links |
| `max_slides` | Obergrenze pro Datei (Standard 40) |

Die Akzentfarbe wird automatisch so weit aufgehellt oder abgedunkelt, dass sie auf
hellem **und** dunklem Grund lesbar bleibt – eine sehr dunkle CI-Farbe kippt also
nicht in Unlesbarkeit.

Jeder Nutzer kann zusätzlich unter **UserValves** ein eigenes Lieblings-Theme und
seinen Namen für die Titelfolie hinterlegen.

---

## 5. Folientypen

Das Modell übergibt ein JSON-Array. Jedes Objekt hat `type` plus passende Felder;
`notes` (Sprechernotizen) und `kicker` (kleines Label über dem Titel) gehen überall.

```json
[
  {"type": "title",    "title": "…", "subtitle": "…", "kicker": "…"},
  {"type": "agenda",   "title": "Agenda", "bullets": ["…", "…"]},
  {"type": "section",  "title": "Kapitel", "subtitle": "…", "index": 1},
  {"type": "bullets",  "title": "…", "bullets": [{"title": "Label", "text": "Erklärung"}]},
  {"type": "cards",    "title": "…", "cards": [{"title": "…", "text": "…"}]},
  {"type": "stats",    "title": "…", "stats": [{"value": "87 %", "label": "…", "note": "…"}]},
  {"type": "compare",  "title": "…", "columns": [{"title": "Heute", "bullets": []},
                                                  {"title": "Morgen", "bullets": [], "highlight": true}]},
  {"type": "timeline", "title": "…", "steps": [{"label": "Q1", "title": "…", "text": "…"}]},
  {"type": "table",    "title": "…", "columns": ["A", "B"], "rows": [["a1", "b1"]]},
  {"type": "chart",    "title": "…", "chart_type": "column|bar|line|pie|doughnut|stacked",
                       "categories": ["2024", "2025"],
                       "series": [{"name": "Umsatz", "values": [12, 18]}],
                       "takeaway": "Kernaussage in einem Satz"},
  {"type": "quote",    "quote": "…", "attribution": "Name, Rolle"},
  {"type": "image",    "title": "…", "image_url": "https://…", "placement": "right|left|full"},
  {"type": "closing",  "title": "Vielen Dank", "contact": ["mail@firma.at"]}
]
```

Deutsche Feldnamen (`titel`, `untertitel`, `punkte`, `kennzahlen`, `zeitstrahl`,
`diagramm`, `tabelle` …) werden ebenfalls verstanden, ebenso JSON in Code-Fences –
falls das Modell „kreativ" wird, fällt das Tool nicht um.

---

## 6. Was die Engine automatisch macht

- **Kein Textüberlauf**: Zeilenumbrüche werden vor dem Rendern simuliert, Schriftgrößen
  passen sich an, Blöcke werden vertikal ausbalanciert.
- **Konsistente Typografie**: Kacheln und Kennzahlen einer Folie teilen sich eine
  Schriftgröße, damit Zeilen über die Spalten fluchten.
- **Automatischer Umbruch auf Folgefolien** bei mehr als 8 Aufzählungspunkten
  („(Forts.)").
- **Diagramme bleiben PowerPoint-nativ** – im Theme eingefärbt, ohne Gitterrauschen,
  in PowerPoint weiterhin editierbar.
- **Fehlertoleranz**: Eine kaputte Folie beendet nicht die Datei, sondern erscheint als
  Hinweis in der Antwort.

---

## 7. Skill: damit sich das Modell immer gleich verhält

Das Tool liefert die Fähigkeit, die Skill liefert das Verhalten – Ablauf, Folienwahl,
Textlängen, Regeln. Datei: `skill-praesentationen-erstellen.md`.

**Installation:** Workspace → **Skills** → **Import** → die `.md`-Datei wählen. Alternativ
**Create** und den Inhalt einfügen; `name` und `description` kommen aus dem Frontmatter.

**Aktivieren** – drei Wege:

| Weg | Verhalten |
|---|---|
| Am Modell hinterlegen (**Workspace → Models → Skills**) | Empfohlen. Das Modell sieht zunächst nur Name und Beschreibung und lädt den vollen Text selbst nach, sobald es um Präsentationen geht. |
| Pro Chat über das **+**-Menü | Gilt für diese eine Unterhaltung. |
| `$` im Eingabefeld tippen | Injiziert den Text einmalig in die Nachricht. |

Zwei Hinweise: Das Nachladen am Modell braucht **natives Function Calling** – ohne das
sieht das Modell nur die Beschreibung und kommt nie an die Anleitung. Und Skills sind
standardmäßig privat: Wenn Kolleginnen und Kollegen dasselbe Modell nutzen sollen, muss
die Skill für sie freigegeben sein, sonst wird sie stillschweigend ignoriert.

Zum Anpassen einfach die Markdown-Datei bearbeiten – etwa Standardumfang, Theme-Zuordnung
oder die Rückfrage-Schwelle.

## 8. Wenn etwas nicht klappt

Das Tool bringt zwei Diagnosefunktionen mit:

- **`check_link`** – „Prüf bitte diesen Link: …". Untersucht einen konkreten
  Download-Link: Gibt es den Eintrag, gehört er dir, liegt die Datei da? Das ist die
  erste Anlaufstelle bei einem toten Link.
- **`check_setup`** – „Prüf bitte das Setup des Präsentations-Generators". Kontrolliert
  Bibliotheken, Datei-Schnittstelle, Upload-Verzeichnis und Schreibrechte, legt testweise
  eine Datei an, prüft den Abruf und räumt sie wieder weg.

| Symptom | Ursache / Lösung |
|---|---|
| Download-Link liefert `{"detail":"We could not find what you're looking for :/"}` | Betraf die Versionen 1.0.0 bis 1.2.0, behoben ab **1.3.0**. **Wichtig:** Links aus Decks, die mit einer älteren Version erzeugt wurden, bleiben auch nach dem Update tot – dafür wurde nie ein Datenbankeintrag geschrieben. Diese Präsentationen neu erstellen. Ob genau das vorliegt, sagt `check_link`. |
| `check_setup` meldete „Alles bereit", der Link ging trotzdem nicht | Das war der Fehler in 1.2.0 unter Open WebUI 0.9.0: Dort ist die Modell-Schicht asynchron; nicht abgewartete Aufrufe liefern ein Coroutine-Objekt, das wie ein Erfolg aussieht. Ab 1.3.0 wird jeder Aufruf abgewartet und die abgelegte Datei über Pfad **und Dateigröße** verifiziert. |
| Nach dem Update auf 1.3.0 weiterhin tot | `check_link` mit dem neuen Link aufrufen – er unterscheidet zwischen fehlendem Eintrag, fehlender Datei (kein Volume, mehrere Instanzen) und fremdem Konto. |
| Auch ein frisch erzeugter Link geht nicht, `check_link` sagt „müsste funktionieren" | Dann liegt es außerhalb des Tools: Link im angemeldeten Browser öffnen (kein privates Fenster), und beim Reverse Proxy prüfen, ob `/api/v1/files/…` durchgereicht wird. |
| „Open-WebUI-Dateimodell nicht verfügbar" | Sehr alte oder stark abweichende Open-WebUI-Version. Das Tool schreibt dann in `DATA_DIR/uploads` – Pfad steht in der Meldung. |
| Kein Download-Link, nur Text | Function Calling auf `Native` stellen (siehe oben). |
| `ModuleNotFoundError: pptx` | `pip install python-pptx` im Container ausführen und Open WebUI neu starten. |
| Link führt ins Leere | `base_url` in den Valves auf die externe Adresse setzen (z. B. `https://chat.firma.at`). |
| Schriften sehen anders aus | `font_heading` / `font_body` auf Schriften setzen, die auf den Zielrechnern installiert sind (Cambria, Calibri, Arial sind mit Office überall vorhanden). |
