# Kandidatenprofil-Generator

Erstellt Kandidatenprofile im Vertico-Design auf Knopfdruck aus einem
Interviewprotokoll (Textdatei aus ECHO KI) und optional einem Lebenslauf --
läuft im Browser. Die KI-Unterstützung läuft über die Anthropic API (Claude).

## Was das Programm macht

1. Du lädst per Drag & Drop das Interviewprotokoll (.txt) und optional den
   Lebenslauf (.docx/.pdf) hoch. Ein Kandidatenfoto für das Deckblatt wird
   dabei automatisch aus dem Lebenslauf erkannt (siehe eigener Abschnitt
   unten) -- ein separater Upload ist nur nötig, falls kein oder das
   falsche Foto erkannt wird.
2. Claude (Anthropic API) liest beide Dokumente, extrahiert die Personendaten
   (Name, Geburtsdatum, Gehalt usw.) und entwirft die Fließtexte für alle
   festen Abschnitte (Hintergrund, Motivation, Eignung, ... Gesamteinschätzung).
3. Du prüfst und bearbeitest alles in der Oberfläche -- nichts wird ungeprüft
   in ein fertiges Dokument geschrieben.
4. Auf Knopfdruck entsteht eine .docx-Datei exakt im Design deiner Vorlage
   (Logo, Farben, Schrift, Kopf-/Fußzeile) und wird lokal gespeichert.

## Automatische Fotoerkennung aus dem Lebenslauf

Beim Hochladen eines Lebenslaufs (.docx oder .pdf) sucht die App automatisch
nach einem darin eingebetteten Foto und schlägt es für das Deckblatt vor. Die
Erkennung wählt das größte enthaltene Bild, das ein foto-typisches
Seitenverhältnis hat -- kleine Icons, Symbole und breite Trennlinien/Banner
werden dabei automatisch ausgeschlossen.

Das erkannte Foto wird **immer als Vorschau angezeigt**, bevor es verwendet
wird -- nie automatisch "blind" übernommen. Falls kein Foto gefunden wird
oder das falsche Bild erkannt wurde, kannst du direkt darunter manuell ein
anderes Foto hochladen; ein manuell hochgeladenes Foto hat immer Vorrang vor
der automatischen Erkennung.

## API-Key besorgen (einmalig)

Auf **https://platform.claude.com/settings/keys** einen Account anlegen
bzw. einloggen und unter "Create Key" einen neuen API-Key erstellen. Der
Key wird nur einmal angezeigt -- direkt sicher speichern (z.B. in einem
Passwort-Manager). Für die Nutzung ist eine hinterlegte Zahlungsmethode
nötig (Abrechnung erfolgt nach tatsächlicher Nutzung/Token-Verbrauch, keine
feste monatliche Gebühr).

## API-Key dauerhaft speichern (empfohlen)

**Die einfachste Methode:** Im Projektordner liegt die Datei
**„API-Key speichern.command"**. Doppelklick darauf (einmalig ggf.
Rechtsklick → „Öffnen", wie bei den anderen `.command`/`.app`-Dateien auch
-- siehe unten). Es öffnet sich ein Terminal-Fenster und fragt nach deinem
API-Key. Key einfügen (`Cmd + V`) und Enter drücken -- fertig, die Datei
wird automatisch richtig angelegt. Die Eingabe erscheint dabei bewusst
unsichtbar (wie bei einem Passwortfeld); das ist normal.

Ab dem nächsten Start ist das API-Key-Feld links in der App automatisch
ausgefüllt -- du kannst es dort bei Bedarf trotzdem jederzeit überschreiben.

**Von Hand (alternativ, falls das Skript nicht funktioniert):**
1. Im Projektordner den (versteckten) Ordner `.streamlit` öffnen
   (`Cmd + Umschalt + .` im Finder, falls nicht sichtbar).
2. Die Datei `secrets.toml.example` kopieren und die Kopie in
   `secrets.toml` umbenennen (Endung `.example` weglassen).
3. In `secrets.toml` den Platzhalter durch deinen echten Key ersetzen:
   ```
   ANTHROPIC_API_KEY = "sk-ant-dein-echter-key"
   ```
4. Speichern.

**Hinweis:** `secrets.toml` bleibt lokal auf deinem Rechner und wird von der
App nirgends automatisch hochgeladen. Trotzdem gilt: Diese Datei enthält ein
Geheimnis -- nicht per E-Mail verschicken und beim Weitergeben des
Projektordners an Kollegen vorher entfernen.

## Programm starten -- die einfachste Methode: die App

Im Projektordner liegt **„Kandidatenprofil-Generator.app"** -- eine echte,
doppelklickbare Mac-Anwendung mit eigenem Symbol (kein Terminal-Skript zum
manuellen Ausführen mehr).

**Einmalige Vorbereitung:**
1. Rechtsklick (bzw. zwei Finger auf dem Trackpad) auf
   „Kandidatenprofil-Generator.app" → **„Öffnen"** wählen. (Nur beim
   allerersten Mal nötig -- macOS warnt bei Programmen aus dem Internet
   standardmäßig; über "Öffnen" bestätigst du einmalig, dass du dem
   Programm vertraust. Ein normaler Doppelklick reicht danach.)
2. Falls ein Sicherheitsdialog erscheint ("kann nicht geöffnet werden, da
   der Entwickler nicht verifiziert werden kann"): Systemeinstellungen →
   Datenschutz & Sicherheit → dort erscheint unten ein Button
   "Dennoch öffnen" -- anklicken.
3. Ein Terminal-Fenster öffnet sich automatisch und richtet beim ersten
   Start alles ein (venv anlegen, Pakete installieren -- dauert ein bis
   zwei Minuten). Danach öffnet sich automatisch ein Browser-Fenster mit
   der App. Das Terminal-Fenster bleibt bewusst sichtbar im Hintergrund
   (zeigt Fortschritt/Fehlermeldungen); du kannst es minimieren, aber
   nicht schließen, solange du die App nutzt.

**Ab dann:** Einfach Doppelklick auf „Kandidatenprofil-Generator.app" --
die komplette App startet direkt im Browser.

**Ins Dock legen (empfohlen):** „Kandidatenprofil-Generator.app" in der
Menüleiste festhalten und ins Dock ziehen -- dann reicht künftig ein Klick
auf das Symbol im Dock, wie bei jeder anderen Mac-App auch. Alternativ in
den Programme-Ordner oder auf den Schreibtisch ziehen/kopieren (die App
funktioniert unabhängig davon, wo sie liegt, solange der restliche
Projektordner mit `app.py` usw. nicht verschoben wird -- am einfachsten:
die App im Projektordner liegen lassen und nur eine Verknüpfung/Alias
davon an anderer Stelle ablegen).

Zum Beenden das Terminal-Fenster schließen oder darin `Strg+C` drücken.

## Alternative: Startskript ohne eigenes App-Symbol

Falls die `.app` aus irgendeinem Grund nicht funktioniert, liegt im
Projektordner zusätzlich die Datei
**„Kandidatenprofil-Generator starten.command"** -- macht dasselbe, aber
ohne eigenes App-Symbol (öffnet direkt ein Terminal-Fenster statt über den
Umweg der App). Bedienung genauso: einmalig Rechtsklick → „Öffnen", danach
reicht ein normaler Doppelklick.

## Einrichtung von Hand (alternativ, über das Terminal)

1. **Python prüfen** (Terminal öffnen, z.B. über Spotlight-Suche "Terminal"):
   ```
   python3 --version
   ```
   Falls das einen Fehler gibt, Python über https://www.python.org/downloads/
   installieren (oder `brew install python` falls Homebrew vorhanden ist).

2. **Projektordner öffnen** (im Terminal in diesen Ordner wechseln, z.B.):
   ```
   cd ~/Downloads/kandidatenprofil-generator
   ```

3. **Virtuelle Umgebung anlegen und Pakete installieren:**
   ```
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```

4. **API-Key besorgen und ggf. dauerhaft speichern** -- siehe Abschnitte
   oben.

## Programm starten (manuell über das Terminal)

Im Terminal, im Projektordner (bei jedem Neustart des Terminals muss Schritt
"venv aktivieren" wiederholt werden):

```
source venv/bin/activate
streamlit run app.py
```

Es öffnet sich automatisch ein Browser-Fenster mit der Bedienoberfläche.
Zum Beenden im Terminal `Strg+C` drücken.

## Aufbau des Projekts

- `app.py` -- die Bedienoberfläche (Streamlit, läuft im Browser)
- `ai_extraction.py` -- Anbindung an die Anthropic API (Claude)
- `docx_builder.py` -- erzeugt das fertige Word-Dokument im Vertico-Design
- `config.py` -- alle Design-Werte (Farben, Schriftgrößen, Ränder, Logo-
  Position) und die feste Abschnittsstruktur des Kandidatenprofils. Jeder
  Wert ist im Kommentar mit seiner Quelle aus der Originalvorlage belegt.
- `file_reading.py` -- liest Text aus .txt/.docx/.pdf-Uploads
- `image_extraction.py` -- erkennt automatisch ein Kandidatenfoto in
  hochgeladenen .docx/.pdf-Lebensläufen
- `assets/` -- die aus deiner Vorlage extrahierten Logo-Bilder (Deckblatt
  und Kopfzeile)
- `Kandidatenprofil-Generator.app` -- doppelklickbare Mac-App zum Starten
- `Kandidatenprofil-Generator starten.command` -- Terminal-Startskript,
  wird von der App im Hintergrund genutzt (siehe oben)
- `API-Key speichern.command` -- fragt den API-Key interaktiv ab und legt
  `secrets.toml` automatisch korrekt an

## Wichtige Hinweise

- **Kosten:** Die Anthropic API wird nach tatsächlicher Nutzung abgerechnet
  (Token-Verbrauch), keine feste Abo-Gebühr. Die genauen Kosten pro
  Kandidatenprofil hängen von der Länge des Interviewprotokolls und des
  gewählten Seitenumfangs ab.
- **Seitenumfang:** Der Regler in der App liefert einen *Richtwert*, keine
  exakte Vorgabe. Wortzahl-pro-Seite wurde aus deiner Beispieldatei
  nachgerechnet (siehe Kommentar in `config.py`), der tatsächliche
  Seitenumbruch in Word hängt aber zusätzlich von Silbentrennung,
  Bildgrößen etc. ab.
- **Faktentreue:** Das Modell ist angewiesen, keine Angaben zu erfinden.
  Fehlt eine Information im Protokoll/Lebenslauf, bleibt das Feld leer bzw.
  wird ein Hinweistext eingefügt -- keine geratenen Werte.
- **Datenschutz:** Aktuell nicht weiter berücksichtigt (auf deinen Wunsch).
  Für den produktiven Einsatz mit echten Kandidatendaten empfiehlt sich vor
  Rollout eine Prüfung der Anthropic-API-Nutzung (Auftragsverarbeitung).
- **Nächste Schritte / Erweiterungsmöglichkeiten:** weitere Dokumenttypen
  (Berichte) nach demselben Muster ergänzen, Batch-Verarbeitung mehrerer
  Kandidaten, automatische Ablage nach Kandidatenname.
