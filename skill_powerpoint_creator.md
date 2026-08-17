---
name: praesentationen-erstellen
description: Anleitung für das Erstellen von PowerPoint-Präsentationen mit dem Tool create_presentation. Nutze sie, sobald jemand eine Präsentation, ein Deck, Folien, einen Foliensatz, ein Pitch-Deck oder eine PPTX möchte – beim Neuerstellen ebenso wie beim Überarbeiten, Kürzen oder Erweitern eines Decks. Nutze sie außerdem, wenn ein Download-Link einer Präsentation nicht funktioniert oder ins Leere führt. Enthält den Gesprächsablauf, die Folientypen, die JSON-Struktur, die inhaltlichen Regeln und die Fehlersuche.
---

# Präsentationen erstellen

Diese Anleitung gilt für jeden Auftrag, der auf eine Präsentationsdatei hinausläuft.
Die Datei entsteht ausschließlich über das Tool **`create_presentation`** – niemals als
Markdown-Folien im Chat, niemals als beschriebene Gliederung „zum Selbstbauen", außer
der Nutzer verlangt das ausdrücklich.

## 1. Ablauf

**Zuerst entscheiden: bauen oder einmal nachfragen.**

Sofort bauen, wenn mindestens eines zutrifft:

- Thema und Zielgruppe oder Anlass gehen aus dem Auftrag hervor
  („Investorenpitch für unsere Logistik-Software", „10 Folien für die Teamklausur")
- Der Nutzer hat Material mitgeliefert (Dokument, Zahlen, Text, vorheriger Chatverlauf)
- Der Nutzer signalisiert Eile oder sagt sinngemäß „mach einfach"

Einmal nachfragen, wenn der Auftrag nur ein Stichwort ist („Präsentation über KI").
Dann gilt:

- **Ein** Rückfrage-Block, höchstens vier Fragen, danach wird gebaut
- Gefragt wird nur, was das Deck verändert: Zielgruppe, Anlass, Umfang, gewünschte
  Kernbotschaft, vorhandene Zahlen
- Nicht gefragt wird nach Design, Farben oder Schriften – dafür gibt es Themes, die du
  selbst passend wählst
- Antwortet der Nutzer nicht oder ausweichend: mit plausiblen Annahmen bauen und die
  Annahmen unter dem Ergebnis in ein bis zwei Sätzen offenlegen

Nie zwei Rückfragerunden hintereinander. Nie eine Gliederung zur Freigabe vorlegen und
dann auf Bestätigung warten – lieber bauen, das Deck ist in Sekunden nachgeschärft.

## 2. Umfang und Aufbau

| Anlass | Folien |
|---|---|
| Kurzes Update, Statusbericht | 5–8 |
| Standardvortrag, Workshop-Input | 10–14 |
| Entscheidungsvorlage, Pitch | 12–18 |
| Schulung, Deep Dive | 18–25 |

Grundgerüst: **Titel → (Agenda ab ~10 Folien) → Inhalt in 2–4 Blöcken → Abschluss.**
Ab drei Themenblöcken jeweils eine `section`-Folie als Trenner. Die Abschlussfolie
trägt die Handlungsaufforderung oder die Entscheidung, um die es geht – nicht bloß
„Danke".

## 3. Folientypen

| Typ | Wofür |
|---|---|
| `title` | Deckblatt, immer die erste Folie |
| `agenda` | Überblick, bis zu 10 Punkte |
| `section` | Kapitel-Trenner mit Nummer |
| `bullets` | Argumente, Beobachtungen, Aufzählungen |
| `cards` | 2–6 gleichrangige Elemente: Bausteine, Angebote, Handlungsfelder |
| `stats` | 2–4 Kennzahlen groß herausgestellt |
| `compare` | Zwei Seiten: heute/morgen, Variante A/B, pro/contra |
| `timeline` | 3–5 Schritte, Phasen, Meilensteine |
| `table` | Strukturierter Vergleich, bis 9 Zeilen |
| `chart` | Zahlenverlauf oder Verteilung, mit `takeaway` als Kernaussage |
| `quote` | Belegtes Zitat, Kundenstimme, Fundstelle |
| `image` | Halbseitiges Bild mit Text, `image_url` erforderlich |
| `closing` | Abschluss, Kontakt, nächster Schritt |

**Layouts nach Inhalt wählen, nicht nach Abwechslung.** Wo Zahlen im Mittelpunkt
stehen, gehören `stats` oder `chart` hin; wo etwas zeitlich abläuft, `timeline`; wo
zwei Optionen gegeneinanderstehen, `compare`. Wenn ein Thema zehn Folien lang wirklich
nur Argumente kennt, sind zehn `bullets`-Folien richtig. Umgekehrt: reine
Stichwortlisten dort, wo eine Struktur im Inhalt steckt, verschenken das Deck.

## 4. Aufruf des Tools

```
create_presentation(
  title    = "Titel der Präsentation",
  subtitle = "Untertitel für die Titelfolie",
  slides   = "[ ... JSON-Array ... ]",
  theme    = "midnight",
  mode     = "auto",
  author   = "Name, falls bekannt"
)
```

`slides` ist ein JSON-Array. Jedes Objekt hat `type` plus die passenden Felder; optional
überall `notes` (Sprechernotizen) und `kicker` (kleines Label über dem Titel).

```json
[
  {"type":"title","title":"…","subtitle":"…","kicker":"…"},
  {"type":"agenda","title":"Agenda","bullets":["…","…"]},
  {"type":"section","title":"Kapitel","subtitle":"…","index":1},
  {"type":"bullets","title":"…","subtitle":"…",
   "bullets":[{"title":"Kurzlabel","text":"Erklärung dazu"},"einfacher Punkt"]},
  {"type":"cards","title":"…","cards":[{"title":"…","text":"…"}]},
  {"type":"stats","title":"…","stats":[{"value":"87 %","label":"…","note":"…"}]},
  {"type":"compare","title":"…","columns":[
     {"title":"Heute","subtitle":"…","bullets":["…"]},
     {"title":"Morgen","bullets":["…"],"highlight":true}]},
  {"type":"timeline","title":"…","steps":[{"label":"Q1","title":"…","text":"…"}]},
  {"type":"table","title":"…","columns":["A","B"],"rows":[["a1","b1"]],"note":"Quelle"},
  {"type":"chart","title":"…","chart_type":"column",
   "categories":["2024","2025"],"series":[{"name":"Umsatz","values":[12,18]}],
   "takeaway":"Was die Zahlen bedeuten, in einem Satz"},
  {"type":"quote","quote":"…","attribution":"Name, Rolle"},
  {"type":"closing","title":"…","subtitle":"…","contact":["mail@firma.at"]}
]
```

`chart_type`: `column`, `bar`, `line`, `pie`, `doughnut`, `stacked`.

## 5. Textmaße

Das Layout passt Schriftgrößen selbst an, aber gedrängte Folien bleiben gedrängt.
Deshalb:

- Folientitel unter 60 Zeichen
- Bullet-Text unter 140 Zeichen, höchstens 6 Punkte pro Folie
- Kartentext unter 180 Zeichen
- Kennzahl-Werte kurz halten: „1,3 Mio. €" statt „1.284.000,00 Euro"
- Timeline-Texte unter 120 Zeichen
- Tabellenzellen unter 40 Zeichen, höchstens 6 Spalten

Lieber eine Folie mehr als eine überfüllte.

## 6. Inhaltliche Regeln

**Keine erfundenen Zahlen.** Konkrete Werte, Marktanteile, Umsätze, Prozentsätze, Daten
und Zitate stammen ausschließlich aus dem, was der Nutzer mitgeliefert hat, aus
angehängten Dateien oder aus einer Recherche, die du tatsächlich durchgeführt hast.
Fehlen belastbare Zahlen, gibt es drei zulässige Wege:

1. Die Kennzahl- oder Diagrammfolie weglassen und qualitativ argumentieren
2. Platzhalter offen kennzeichnen, etwa als Kennzahl-Wert `[Wert]` mit dem Hinweis
   „vom Controlling zu ergänzen" im Feld `note`
3. Den Nutzer nach den Zahlen fragen, wenn sie den Kern des Decks ausmachen

Das gilt genauso für Zitate: kein `quote`-Slide mit einer erfundenen Kundenstimme oder
einem Satz, den eine reale Person nie gesagt hat. Erfundene Quellenangaben unter
Tabellen sind ebenfalls tabu.

**Sprechernotizen, wo sie helfen.** Bei Folien, die vorgetragen werden – Kennzahlen,
Diagramme, Entscheidungspunkte – gehört in `notes`, was gesagt, aber nicht auf die
Folie geschrieben wird: Einordnung, Hintergrund, Übergang zur nächsten Folie. Bei
selbsterklärenden Folien wie Agenda oder Kapiteltrenner kann `notes` entfallen.

## 7. Design

Wähle das Theme passend zum Thema, ohne zu fragen:

| Theme | Passt zu |
|---|---|
| `midnight` | Vorstand, Strategie, Finanzen, klassisch seriös |
| `graphite` | Technik, Produkt, reduziert und sachlich |
| `teal` | Software, Digitalisierung, Gesundheit |
| `ocean` | Banken, Versicherung, öffentliche Hand |
| `forest` | Nachhaltigkeit, Energie, Landwirtschaft |
| `terracotta` | Handwerk, Bau, Immobilien, Gastronomie |
| `coral` | Marketing, Kampagnen, junge Zielgruppen |
| `berry` | Lifestyle, Kulinarik, Premium-Marken |
| `cherry` | Bühnen-Keynote, starke These, hoher Kontrast |

`mode` bleibt auf `auto`, außer der Nutzer wünscht durchgehend hell (`light`) oder den
dunklen Premium-Look (`dark`). Nennt der Nutzer eigene CI-Farben, weise ihn darauf hin,
dass diese einmalig in den Tool-Einstellungen (Valves) hinterlegt werden können.

## 8. Nach dem Erzeugen

Das Tool stellt den Download-Link selbst in die Antwort. **Wiederhole ihn nicht.**
Antworte stattdessen kurz:

- ein Satz zum Aufbau des Decks („Titel, drei Kapitel, Kennzahlenblock, Zeitstrahl,
  Entscheidungsfolie")
- getroffene Annahmen, falls du welche getroffen hast
- offen gelassene Platzhalter, falls vorhanden
- ein konkretes Angebot zum Nachschärfen, keine Floskel: etwa „Soll ich die
  Wettbewerbsfolie durch einen Vergleich der drei Anbieter ersetzen?"

Bei Änderungswünschen wird die Präsentation komplett neu erzeugt – mit dem angepassten
JSON und demselben Theme. Einzelne Folien lassen sich nicht nachträglich patchen.
Übernimm dabei die unveränderten Folien wortgleich, damit nur das geändert wird, was
gemeint war.

## 9. Wenn ein Download-Link nicht funktioniert

Meldet der Nutzer einen toten Link oder eine Fehlermeldung wie
„We could not find what you're looking for", **rate nicht und erzeuge das Deck nicht
einfach neu**. Erst messen:

1. **`check_link`** mit dem gemeldeten Link oder der Datei-ID aufrufen. Die Funktion
   sagt, ob der Datenbank-Eintrag existiert, wem er gehört und ob die Datei auf der
   Platte liegt – und nennt die passende Empfehlung.
2. Nur wenn `check_link` das Deck als „vor dem Update erzeugt" einordnet, die
   Präsentation neu erstellen.
3. Bei einem Verdacht auf ein grundsätzliches Problem – etwa wenn auch ein frisch
   erzeugter Link tot ist – zusätzlich **`check_setup`** aufrufen und den Bericht
   ungekürzt weitergeben.

Gib dem Nutzer den Befund im Klartext weiter, statt ihn zu beruhigen. Ein „sollte
eigentlich funktionieren" hilft niemandem; die Diagnosefunktionen liefern eine echte
Antwort.
