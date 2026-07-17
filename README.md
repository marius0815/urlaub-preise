# cron-urlaubspreise – Marktbeobachtung

> **Namenskonvention:** Das `cron-`-Prefix kennzeichnet dieses Verzeichnis als
> Projekt eines geplanten (gescheduleten) Tasks. Alle weiteren Cron-Projekte folgen
> demselben Schema (`cron-<thema>`), damit geplante Aufgaben auf einen Blick erkennbar sind.

Regelmäßige, automatisierte Abfrage von Flugpreisen für grob definierte
Reisen. Läuft **lokal** über die Windows-Aufgabenplanung und einen
**headless Browser** (Playwright/Chromium) – unsichtbar im Hintergrund, **datumsgenau**.
Sucht die aktuell günstigsten Optionen und pflegt eine Historie sowie eine
Best-of-Übersicht. **Zustellung erfolgt passiv**: kein aktiver Benachrichtigungskanal –
die Ergebnisse liegen lokal (`best-of.xlsx` / `preise.csv`) und werden bei Bedarf
selbst eingesehen. **Läuft nur, wenn der PC an ist**; verpasste Läufe werden **nicht**
nachgeholt.

> **Warum lokal + headless statt Cloud?** Datumsgenaue Flugpreise stehen nur hinter
> der interaktiven Suchmaske der Portale – ein statischer Cloud-Abruf liefert nur grobe
> „ab"-Lockpreise (im Test um Faktor ~3 daneben). Ein echter Browser liefert exakte
> Preise; headless läuft er unsichtbar im Hintergrund, ohne bei der Arbeit zu stören.
> Preis dafür: der PC muss laufen.

## Architektur (festgelegt)

| Aspekt | Entscheidung |
|---|---|
| Ausführung | **lokal** – Windows-Aufgabenplanung + headless Browser (Playwright/Chromium) |
| Takt | **konfigurierbar**: tägliches Zeitfenster (von/bis) + Frequenz in Stunden (s. `job.zeitplan`) |
| Verpasste Läufe | werden **nicht** nachgeholt (`verpasste_nachholen: false`); PC muss an sein |
| Persistenz | **lokal** im Projektordner; GitHub optional zur Versionierung |
| Zustellung | **passiv** – kein E-Mail/Push; Ergebnisse in `best-of.xlsx` / `preise.csv` selbst einsehen |

## Dateien

| Datei | Zweck | Verhalten |
|---|---|---|
| `config-pfad.json` | **Zentrale Pfad-Konfiguration**: `zielordner` = Ordner für alle Dateien (absolut oder relativ zum Projektordner), `config` und `best_of_datei` = Dateinamen **relativ zum Zielordner**. Aktueller Zielordner: `C:/Users/marius.neumann/OneDrive/Documents/Urlaub/Best-Of-Urlaub`. | – |
| `config.xlsx` | Einzige Quelle der Vorgaben. **Hier pflegst du alles.** Reiter **„Reisen"** = Tabelle mit einer Zeile je Reise; Reiter **„Einstellungen"** = Schlüssel/Wert (Abflughäfen, Zeitplan, Pfade, n_best). Liegt im Zielordner (s. `config-pfad.json`). Eine `.json` wird als Altformat weiterhin akzeptiert. | – |
| `preise.csv`  | Vollständige Preis-**Historie** für den Preisverlauf. | nur **anhängen** |
| `best-of.xlsx`| **Best-of-Übersicht** mit mehreren Reitern (s. unten): „Best Of" global + ein Reiter je Reise. Immer nur aktuell beste Angebote, **keine** Historie. Liegt im Zielordner, Dateiname pflegbar in `config-pfad.json` → Schlüssel `best_of_datei`. | jeder Lauf **überschrieben** |
| `best-of.html`| **Statische Lese-Ansicht** derselben Inhalte (ohne Charts, ohne JavaScript) – funktioniert auch in der iOS-Dateivorschau/OneDrive-Vorschau. Liegt im Zielordner, Dateiname pflegbar in `config-pfad.json` → `best_of_html`. | jeder Lauf **überschrieben** |
| `best-of-app.html`| **Interaktive App-Ansicht** (Tabs „Best Of"/„Reisen"/„Historie", Filter; braucht JavaScript). Wird nach jedem Lauf automatisch als GitHub-Pages-Seite veröffentlicht: **https://marius0815.github.io/urlaub-preise/** (Repo in `config-pfad.json` → `github_pages_repo`; leer lassen = keine Veröffentlichung). Auf dem iPhone: Adresse in Safari öffnen → Teilen → „Zum Home-Bildschirm". | jeder Lauf **überschrieben + veröffentlicht** |
| `lauf.log`    | **Lauf-Protokoll**: pro Ausführung ein Eintrag **je Reise** (Status, Treffer, günstigster Preis, Dauer). | **rollierend**: nur die letzten 7 Tage (wie die Historie) |
| `README.md`   | Diese Anleitung. | – |

## Ablauf eines Laufs

Ein Lauf wird von der Windows-Aufgabenplanung gestartet und:

1. liest `config.xlsx` ein (Abflughäfen, Reisen, Zeitplan).
2. fragt für jede `aktiv: true`-Reise die Routen (Abflughäfen × Ziele) per
   **headless Browser** auf Google Flights ab – datumsgenau.
3. hängt die Ergebnisse mit Zeitstempel an `preise.csv` an (Historie).
4. baut `best-of.xlsx` neu: je Reise die **Top N** günstigsten Angebote
   (`n_best` aus `config.xlsx`, Standard 3). Angebote unter `Schwelle_EUR` werden markiert.
5. baut `best-of.html` (statische Lese-Ansicht) und `best-of-app.html` (App-Ansicht) neu
   und veröffentlicht die App-Ansicht als GitHub-Pages-Seite (s. Dateien-Tabelle).
6. schreibt **je Reise einen Eintrag** in `lauf.log` (s. unten).

## Logging (`lauf.log`)

Da zur Laufzeit **keine KI** mitläuft, ist das Log die einzige Kontrolle, ob ein Lauf
sauber war. Pro Ausführung wird **je Reise genau eine Zeile** angehängt:

```
<zeitstempel> | <reise> | Status: <OK|KEINE_TREFFER|BLOCKIERT|FEHLER> | Routen: <n> | gueltige Treffer: <n> | guenstigster: <preis> EUR (<route>) | Dauer: <sek>s | <hinweis>
```

Beispiel:

```
2026-06-14T14:00:07 | kenia | Status: OK | Routen: 10 | gueltige Treffer: 7 | guenstigster: 1125 EUR (FRA->MBA) | Dauer: 138s |
2026-06-14T14:02:31 | kenia | Status: BLOCKIERT | Routen: 10 | gueltige Treffer: 0 | guenstigster: - | Dauer: 22s | Captcha auf Google Flights
```

`Status`-Werte: **OK** (Treffer geschrieben), **KEINE_TREFFER** (Abfrage lief, aber nichts
erfüllte die Filter), **BLOCKIERT** (Captcha/Bot-Wall), **FEHLER** (Skript-/Ladefehler).
So erkennst du auf einen Blick, wenn das Scraping gewartet werden muss.

**Wichtig:** Reisen, die in einem Lauf **gar nicht abgefragt** wurden (z. B. wegen
`--limit` übersprungen), erzeugen **keine** Log-Zeile. Geloggt wird nur, was tatsächlich
abgefragt wurde.

## Zeitplan (konfigurierbar, `job.zeitplan`)

Der Job läuft **nur innerhalb eines täglichen Zeitfensters** und darin in fester Frequenz:

| Feld | Bedeutung | Beispiel |
|---|---|---|
| `zeitfenster.von` | frühester Start am Tag (HH:MM) | `"08:00"` |
| `zeitfenster.bis` | spätester Lauf am Tag (HH:MM) | `"22:00"` |
| `frequenz_stunden` | Abstand zwischen Läufen im Fenster | `6` |
| `verpasste_nachholen` | verpasste Läufe (PC war aus) nachholen? | `false` |

Beispiel `08:00`–`22:00`, alle `6` h → Läufe um **08:00, 14:00, 20:00**. War der PC
zu einem Slot aus, wird dieser **nicht** nachgeholt – der nächste reguläre Slot greift.
Diese Werte werden beim Einrichten in die Trigger der Windows-Aufgabenplanung übersetzt
(„Task bei verpasstem Start sobald möglich ausführen" bleibt **deaktiviert**).

## Best-of-Excel: Reiter-Struktur

Die Datei wird bei jedem Lauf frisch erzeugt und enthält **nur aktuell gültige** Angebote
(Filter: Preis, max. Flugdauer/Richtung). Zwei Reiter-Typen:

| Reiter | Inhalt |
|---|---|
| **`Best Of`** | Die **Top N** besten gültigen Angebote **über alle Reisen** hinweg – die globale Bestenliste. `N` aus `ausgaben.best_of_excel.n` (Standard 3). |
| **je Reise** (Reitername = Reisename, z. B. `kenia`) | Für **diese** Reise der **beste Preis je Route/Ziel** (Abflughafen × Ziel), nach Preis sortiert. |

So zeigt „Best Of" die absoluten Top-Treffer, und jeder Reise-Reiter den Detail-Überblick,
welche Route innerhalb der Reise gerade am günstigsten ist.

## Vorgaben pflegen

Vorgaben ändern = einfach `config.xlsx` in Excel bearbeiten und speichern. Der nächste
Lauf nutzt automatisch die neue Version – die Routine selbst muss **nicht** neu angelegt
werden.

### Reiter „Einstellungen" (zentral, gilt für alle Reisen)

| Schlüssel | Bedeutung | Beispiel |
|---|---|---|
| `abflug` | mögliche Abflughäfen (IATA), Komma-getrennt | `FRA, AMS, BER` |
| `frequenz_stunden` | Job-Intervall in Stunden | `1` |
| `verpasste_nachholen` | verpasste Läufe nachholen (`ja`/`nein`) | `ja` |
| `n_best` | Top N im Reiter „Best Of" | `30` |
| `hopping_min_aufenthalt_h` | Default-Mindestaufenthalt je Hopping-Station in Stunden (greift ohne Klammer-Angabe am Ziel) | `48` |
| `log_datei` | Lauf-Protokoll, relativ zum Projektordner | `lauf.log` |

### Reiter „Reisen" (Tabelle, eine Zeile je Reise)

| Spalte | Bedeutung | Beispiel |
|---|---|---|
| `Name` | Anzeigename (wird Reitername in best-of.xlsx) | `Mittelmeer Sommer` |
| `Aktiv` | ob diese Reise abgefragt wird | `ja` / `nein` |
| `Typ` | Reisetyp: `return` (Standard, auch bei leerem Feld) oder `hopping` (s. unten) | `return` |
| `Ziele` | bei `return`: **alternative** Zielflughäfen (IATA), Komma-getrennt; bei `hopping`: **Stationen in Abflug-Reihenfolge**, Alternativen je Station per `/`, optional Mindestaufenthalt in Stunden in Klammern | `PMI, HER, FAO` bzw. `SFO/LAX (72), HNL (48)` |
| `Zeitraum` | Reisefenster als Text | `08.08.2026-28.08.2026` |
| `Naechte` | bei `return`: Aufenthaltsdauern in Nächten, Komma-getrennt; bei `hopping`: **minimale und maximale Gesamtdauer** der Reise in Nächten | `7, 14` |
| `Schwelle_EUR` | Wunschpreis – darunter wird in Excel markiert | `400` |
| `Personen` | Anzahl Reisende (optional, Standard 1) | `2` |
| `Max_Flugdauer_h` | max. **Bruttozeit** (Abflug → finale Ankunft) **je Richtung**, in Stunden (optional) | `16` |

### Reisetypen

- **`return`** (Standard): Hin- und Rückflug zwischen einem Abflughafen und **einem** Ziel;
  die Ziele-Liste sind Alternativen, aus denen das günstigste Angebot gesucht wird.
- **`hopping`**: eine **Kette von Zielen**, die ab dem Abflughafen in der angegebenen
  Reihenfolge abgeflogen wird (und am Ende zurück zum Abflughafen). Hinter jedem Ziel
  kann in Klammern der **Mindestaufenthalt in Stunden** stehen, z. B.
  `SFO (72), HNL (48), MNL (96), NBO (72)` = San Francisco (min. 72 h) → Hawaii (48 h) →
  Manila (96 h) → Kenia (72 h). Ohne Klammer gilt der Default aus Einstellungen →
  `hopping_min_aufenthalt_h` (aktuell 48 h). **Alternativen je Station** werden per `/`
  angegeben, z. B. `SFO/LAX (72), HNL (48), MNL/CRK (96), NBO (72)` – jede Kombination
  wird als eigene Ketten-Variante bepreist (Etappen-Abfragen werden dabei wiederverwendet;
  ab 12 Varianten wird gekappt).

  **So wird bepreist:** Jede Etappe wird als **One-Way-Suche** abgefragt; der Kettenpreis
  ist die Summe der günstigsten gültigen Etappen. Die Folgetermine ergeben sich aus
  Ankunftstag + Mindestaufenthalt (tagesgenau geplant); die komplette Kette muss ins
  `Zeitraum`-Fenster passen. `Naechte` begrenzt die **Gesamtdauer**: Ist die geplante
  Kette kürzer als das Minimum, wird der Aufenthalt an der **letzten Station** verlängert;
  liegt sie über dem Maximum, gilt die Datums-Probe als ungültig. Je (Abflughafen × Startdatum-Probe) entsteht ein Angebot;
  `Schwelle_EUR` gilt auf den **Gesamtpreis**. Der „Suche öffnen"-Link führt zur
  vorbefüllten **Multi-City-Suche** bei Google Flights (dort steht der buchbare Preis –
  die Etappen-Summe kann leicht darüber liegen, weil Durchgangstarife fehlen).
  In Excel zeigen die Hinflug-/Rückflug-Spalten die **erste bzw. letzte Etappe**;
  App und HTML listen alle Etappen einzeln.

**Filter `Max_Flugdauer_h`:** Gilt für **Hin- und Rückflug getrennt**. Die
Bruttozeit ist die gesamte Reisedauer einer Richtung von Abflug bis zur finalen Ankunft
**inkl. aller Zwischenstopps/Umstiege**. Überschreitet **eine** der beiden Richtungen den
Wert, ist die Verbindung **kein gültiger Treffer** und fällt aus Historie und Best-of raus.

Neue Reise = einfach eine weitere Zeile in die Tabelle einfügen.

## Einrichten & Starten

**Voraussetzung (einmalig):** Python + Playwright inkl. Chromium
(`pip install playwright openpyxl`, dann `python -m playwright install chromium`).

| Aktion | Befehl |
|---|---|
| **Zeitplan registrieren** (aus `job.zeitplan`) | `powershell -ExecutionPolicy Bypass -File register_task.ps1` |
| **Manuell jetzt starten** | Doppelklick auf `run_jetzt.bat` (oder `python flugpreise.py`) |
| nur eine Reise | `python flugpreise.py --reise Kenia` |
| schneller Test (Browser sichtbar) | `python flugpreise.py --sichtbar --limit 4` |
| Zeitplan ändern | `frequenz_stunden` / `verpasste_nachholen` in `config.xlsx` (Reiter „Einstellungen") anpassen, dann `register_task.ps1` erneut ausführen |

## Repository & Einrichtung auf neuem Rechner

Das GitHub-Repo [marius0815/urlaub-preise](https://github.com/marius0815/urlaub-preise)
enthält den Code und die Doku; die `index.html` im Repo-Root ist die **automatisch
veröffentlichte App-Ansicht** – sie wird von jedem Lauf per GitHub-API neu committet
(GitHub Pages: https://marius0815.github.io/urlaub-preise/). Laufzeitdaten
(`lauf.log`, `state.json`, `historie.json`, `best-of-data.json`) und die
maschinenspezifische `config-pfad.json` sind bewusst **nicht** im Repo (`.gitignore`).

Einrichtung auf einem neuen Rechner:

1. Repo klonen, dann `pip install playwright openpyxl` und `python -m playwright install chromium`.
2. `config-pfad.beispiel.json` als `config-pfad.json` kopieren und die Pfade anpassen
   (Zielordner = wo config.xlsx, best-of.xlsx und die HTML-Ansichten liegen sollen).
3. Eine `config.xlsx` im Zielordner anlegen (Reiter „Reisen" + „Einstellungen", s. oben) –
   oder eine bestehende dorthin kopieren.
4. Optional Veröffentlichung: GitHub CLI (`gh`) installieren und anmelden,
   `github_pages_repo` in der `config-pfad.json` setzen.
5. `register_task.ps1` ausführen → richtet den stündlichen Task ein.

Hinweis für Code-Änderungen: Vor einem `git push` immer erst `git pull` – der stündliche
Lauf committet die `index.html` direkt auf GitHub, der lokale Stand hängt also hinterher.

`register_task.ps1` legt den Task **cron-urlaubspreise** an: täglich ab `von`, alle
`frequenz_stunden` h bis `bis`, **ohne** Nachholen verpasster Läufe, lautlos via `pythonw`.

## Status

**Lauffähig & geplant.** Skript generalisiert (alle Routen aus `config.xlsx`),
datumsgenaue Preise + Flugdauern **beider Richtungen**, Filter (Dauer beide Richtungen),
Schreiben in `preise.csv` / `best-of.xlsx` (Best Of + Reiter je Reise) / `lauf.log`.
Windows-Aufgabenplanung-Task **cron-urlaubspreise** registriert (08/12/16/20 Uhr, kein Nachholen).

**Bekannte Näherung:** Der Dauer-Filter prüft, dass in **beiden** Richtungen Flüge
≤ `max_flugdauer_h_pro_richtung` existieren; `dauer_rueck` ist die kürzeste gültige
Rück-Richtungs-Dauer (keine itinerar-genaue Kopplung an den jeweiligen Preis).

**Offen / Tuning:** Laufzeit. Volle Config (mehrere Abflughäfen × Ziele × Datums-Proben
× 2 Richtungen) kann pro Lauf länger dauern – bei Bedarf Abflughäfen reduzieren oder
Datums-Schritt vergröbern (`--schritt-tage`).
