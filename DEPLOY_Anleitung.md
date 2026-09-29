# Generator für den Geschäftspartner bereitstellen (Weblink)

Ziel: Dein Partner öffnet nur einen **Link**, gibt ein **Passwort** ein und nutzt
die App im Browser — **keine Installation, kein API-Key, egal ob Mac oder Windows.**
Dein Anthropic-Key liegt sicher auf dem Server; die KI-Kosten laufen über dein Konto.

Die App ist dafür bereits vorbereitet: Sobald auf dem Server die beiden Secrets
`APP_PASSWORD` und `ANTHROPIC_API_KEY` gesetzt sind, schaltet sie automatisch in
den geschützten Modus (Passwortabfrage + zentraler KI-Zugang, kein Key-Feld).

Empfohlener Hoster: **Streamlit Community Cloud** (kostenlos, für genau solche Apps).

---

## Einmalige Einrichtung (ca. 20–30 Min, nur du)

### Schritt 1 – Code zu GitHub bringen (ohne Terminal, per GitHub Desktop)
1. **GitHub-Konto** anlegen: https://github.com/signup (kostenlos).
2. **GitHub Desktop** installieren: https://desktop.github.com
3. In GitHub Desktop: **File → Add local repository** → diesen Projektordner
   (`Kandidatenprofilgenerator`) auswählen. Falls es fragt „create a repository",
   bestätigen.
4. Unten „Summary" ausfüllen (z. B. „Erste Version") → **Commit**.
5. Oben rechts **Publish repository** → Name vergeben, Häkchen bei
   **„Keep this code private"** setzen → **Publish**.
   > Deine `.streamlit/secrets.toml` mit dem echten Key wird dabei automatisch
   > NICHT hochgeladen (sie steht in `.gitignore`). Bitte trotzdem kurz prüfen,
   > dass sie in GitHub Desktop unter den Dateien nicht auftaucht.

### Schritt 2 – App bei Streamlit Community Cloud deployen
1. Auf https://share.streamlit.io mit dem **GitHub-Konto anmelden**.
2. **Create app → Deploy a public/private app from GitHub repo**.
3. Auswählen:
   - Repository: dein soeben veröffentlichtes Repo
   - Branch: `main`
   - Main file path: `app.py`
4. **Advanced settings → Secrets**: dort exakt Folgendes eintragen
   (deine echten Werte einsetzen):
   ```
   ANTHROPIC_API_KEY = "sk-ant-DEIN-ECHTER-KEY"
   APP_PASSWORD = "ein-gutes-passwort"
   ```
5. **Deploy** klicken. Nach 1–3 Minuten ist die App unter einer festen URL
   erreichbar (z. B. `https://kandidatenprofil-generator.streamlit.app`).

### Schritt 3 – An den Partner schicken
Dem Partner per Mail schicken:
- den **Link** zur App
- das **Passwort** (`APP_PASSWORD`)

Fertig. Er klickt den Link, gibt das Passwort ein und legt los — auf Mac,
Windows, Tablet, überall gleich.

---

## Gut zu wissen
- **Kosten:** Die KI-Nutzung deines Partners läuft über deinen Anthropic-Key
  (dein Konto). In der Anthropic Console unter *Settings → Limits* kannst du ein
  monatliches Ausgabenlimit setzen.
- **„App schläft":** Bei Nichtnutzung legt Streamlit die App schlafen; beim
  nächsten Aufruf braucht sie ~30 Sekunden zum Aufwachen. Das ist normal.
- **Passwort ändern / Key rotieren:** in Streamlit Cloud unter
  *App → Settings → Secrets* die Werte anpassen (kein neues Deployment nötig).
- **Nur ein Passwort für alle:** Es gibt ein gemeinsames Passwort, keine
  Einzel-Logins. Für einen kleinen Partnerkreis reicht das; zum Sperren später
  einfach das Passwort ändern.
- **Datenschutz:** Beim Hosten werden hochgeladene Interviewprotokolle/Lebensläufe
  auf dem Streamlit-Server verarbeitet (zusätzlich zur Anthropic-API). Für echte
  Kandidatendaten ggf. vorab die DSGVO-Lage prüfen (Auftragsverarbeitung).

## Alternative ohne GitHub
Wenn du GitHub vermeiden willst: **Hugging Face Spaces** (https://huggingface.co/spaces)
erlaubt das Hochladen der Dateien direkt im Browser (Space-Typ „Streamlit",
Secrets unter *Settings*). Funktioniert genauso, ist aber beim Datei-Upload etwas
fummeliger. Sag Bescheid, dann schreibe ich dir dafür eine eigene Anleitung.
