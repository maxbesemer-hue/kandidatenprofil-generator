"""
Nutzt die Anthropic API (Claude), um aus Interviewprotokoll + Lebenslauf
1) die strukturierten Personendaten zu extrahieren und
2) die Fließtexte für jeden festen Abschnitt des Kandidatenprofils zu entwerfen.

Wichtig zur Zuverlässigkeit (Inhalt): Das Modell wird angewiesen, ausschließlich
Angaben zu verwenden, die im Protokoll oder Lebenslauf tatsächlich enthalten
sind, und fehlende Felder als leeren String zurückzugeben statt zu raten.

Wichtig zur Zuverlässigkeit (Format): Es wird Anthropics "Structured
Outputs"-Funktion genutzt (output_config.format mit JSON-Schema). Das
erzwingt auf API-Seite mittels eingeschränkter Token-Erzeugung (constrained
decoding), dass die Antwort garantiert gültiges JSON gemäß Schema ist.
Quelle: https://platform.claude.com/docs/en/build-with-claude/structured-outputs
"""

import json
import time
import anthropic

import config

try:
    # Rettet leicht beschädigtes/abgeschnittenes JSON als letzte Möglichkeit.
    import json_repair
except Exception:
    json_repair = None

# Transiente Fehler, bei denen ein erneuter Versuch sinnvoll ist:
# 429 = Rate-Limit, 5xx/529 = serverseitige Überlastung/Ausfall.
_RETRYABLE_STATUS = {408, 409, 429, 500, 502, 503, 504, 529}
_MAX_ATTEMPTS = 5

# Hinweis: Modellversion zentral hier pflegen. "claude-sonnet-5" ist zum
# Zeitpunkt der Erstellung dieses Skripts (Juli 2026) das aktuelle
# Sonnet-Modell der Anthropic API und unterstützt Structured Outputs. Bei
# Bedarf in docs.claude.com/en/build-with-claude/structured-outputs prüfen,
# welche Modelle aktuell unterstützt werden.
MODEL = "claude-sonnet-5"

EXTRACTION_CONTENT_RULES = """Du bist Assistent einer Personalberatung (Executive Search) \
und hilfst, Kandidatenprofile aus Interviewprotokollen und Lebensläufen zu erstellen.

Du bekommst den Text eines Interviewprotokolls und optional einen Lebenslauf. \
Extrahiere daraus ausschließlich Informationen, die tatsächlich im Text stehen. \
Erfinde NIEMALS Fakten, Zahlen oder Daten. Wenn eine Information fehlt, gib für \
das jeweilige Feld einen leeren String "" zurück -- rate nichts.

"rolle" ist die volle Zielposition (z.B. "Vertriebsleitung Modernisierung und \
Servicegeschäft Region Südwest"), wie sie im Protokoll erwähnt wird. \
"rolle_kurz" ist eine kurze Version für das Deckblatt (2-4 Worte, z.B. \
"Vertrieb Neuanlage")."""

BLIND_EXTRACTION_RULES = """Du bist Assistent einer Personalberatung (Executive \
Search) und erstellst ein ANONYMISIERTES Kandidatenprofil (Blindprofil). Aus den \
extrahierten Angaben darf sich KEIN Rückschluss auf die konkrete Person ziehen lassen.

Du bekommst ein Interviewprotokoll und optional einen Lebenslauf. Extrahiere \
ausschließlich Informationen, die tatsächlich im Text stehen, und ANONYMISIERE sie:
- "jahrgang": nur das Geburtsjahr (z.B. "1985") oder eine Altersangabe, kein volles Datum.
- "region": nur eine grobe Region statt des genauen Wohnorts (z.B. "Raum Stuttgart", \
"Großraum München", "Südwestdeutschland") -- NIE die genaue Stadt oder Adresse.
- "branche": die Branche des aktuellen Arbeitgebers (z.B. "Maschinenbau", \
"Aufzugsbranche"), NIE den Firmennamen.
- "unternehmensgroesse": grobe Größe des aktuellen Arbeitgebers (z.B. "mittelständisch, \
ca. 500 Mitarbeitende", "Großkonzern"), ohne Firmennamen.
- "familienstand", "kuendigungsfrist", "ist_gehalt", "gehaltswunsch", "verfuegbarkeit": \
wie im Material angegeben.
- "interviewpartner": Name des Beraters/Interviewers (das ist NICHT der Kandidat).
- "rolle" ist die volle Zielposition, "rolle_kurz" eine kurze Version fürs Deckblatt.
Nenne NIEMALS den Namen des Kandidaten, keinen Arbeitgebernamen, keine genaue Adresse \
und keine Kontaktdaten. Erfinde nichts; fehlt eine Angabe, gib "" zurück."""

DRAFTING_BASE_RULES = """Du bist Assistent einer Personalberatung (Executive Search) \
und entwirfst die Fließtext-Abschnitte eines Kandidatenprofils auf Basis eines \
Interviewprotokolls und optional eines Lebenslaufs.

WICHTIGE REGELN:
- Nutze ausschließlich Informationen, die im Protokoll oder Lebenslauf enthalten \
sind. Erfinde keine Fakten, Zahlen, Ereignisse oder Zitate.
- Wenn zu einem Abschnitt im Ausgangsmaterial keine Information vorliegt, \
schreibe einen kurzen, ehrlichen Platzhaltersatz wie "Hierzu liegen aus dem \
Interview keine Angaben vor." statt etwas zu erfinden.
- Schreibe im Stil eines professionellen Personalberatungsberichts: sachlich, \
wertschätzend, in ganzen Absätzen, dritte Person ("Herr/Frau ..." oder Nachname).
- Trenne Absätze mit einer Leerzeile.
- WIEDERHOLUNGEN VERMEIDEN: Das gesamte Profil muss frei von Wiederholungen sein -- \
nicht nur wörtlich, sondern auch inhaltlich und sinngemäß. Jeder Sachverhalt, jedes \
Argument und jedes Beispiel erscheint nur EINMAL im gesamten Profil, und zwar in dem \
Abschnitt, der laut Abgrenzung dafür zuständig ist. Greife Fakten, Aussagen oder \
Formulierungen aus einem Abschnitt in keinem anderen erneut auf, auch nicht \
paraphrasiert oder aus anderer Perspektive. Da du alle Abschnitte in einem Durchgang \
erzeugst, stimme ihre Inhalte bewusst so aufeinander ab, dass sie sich nicht \
überschneiden.
- Schreibe in natürlichem, professionellem Deutsch ohne werbliche Übertreibungen, \
Floskeln oder aneinandergereihte Gedankenstriche."""


def _rating_block(rating: int) -> str:
    """Anweisung, wie die 1-10-Bewertung Tonfall/Gewichtung steuert -- OHNE
    Fakten zu erfinden und OHNE den Kandidaten schlechtzuschreiben."""
    return f"""BEWERTUNG DES KANDIDATEN: {rating} von 10 (1 = schwacher Kandidat, \
10 = perfekter Kandidat; ein solider, guter Durchschnittskandidat liegt bei 5-6).

Diese Bewertung steuert AUSSCHLIESSLICH Tonfall, Enthusiasmus und Gewichtung im \
Bericht -- niemals die Fakten. Beachte dabei strikt:
- Erfinde niemals Schwächen, Kritik oder negative Aussagen, die nicht aus dem \
Material hervorgehen. Faktentreue hat immer Vorrang.
- Schreibe den Kandidaten NIE schlecht oder herabwürdigend, unabhängig von der \
Bewertung. Der Bericht bleibt immer professionell und wertschätzend, ohne \
abwertende Formulierungen.
- Hohe Bewertung (7-10): Stärken und Eignung selbstbewusst und deutlich \
hervorheben; in der Gesamteinschätzung klar empfehlender Tonfall.
- Mittlere Bewertung (5-6): ausgewogenes, positives, aber realistisches Bild \
eines soliden, geeigneten Kandidaten.
- Niedrige Bewertung (1-4): zurückhaltenderer, nüchternerer Tonfall; \
Entwicklungsfelder und Einschränkungen (nur sofern im Material belegt) erhalten \
etwas mehr Gewicht; die Empfehlung fällt vorsichtiger und neutraler aus -- \
weiterhin sachlich, respektvoll und ohne Abwertung.
- Erwähne die Bewertungszahl selbst NICHT im Text."""


def _position_block(zielposition: str) -> str:
    """Anweisung, die Art der Zielposition aus der Bezeichnung (und ggf. der
    Beschreibung) zu erkennen und das Profil inhaltlich darauf auszurichten."""
    return f"""ZU BESETZENDE POSITION: {zielposition}

Erkenne aus dieser Positionsbezeichnung -- und, falls am Ende der Quellen \
vorhanden, aus der Beschreibung der Zielposition -- um welche ART von Position es \
sich handelt (z.B. Vertrieb, Vertriebsleitung, Montageleitung, Montage, \
Serviceleitung, Servicetechnik, technisches Backoffice, Projektleitung o.Ä.).

Richte das Profil INHALTLICH auf genau diese Art von Position aus: wähle \
Schwerpunkte, hervorgehobene Kompetenzen, Fachbegriffe und Bewertungskriterien \
passend zur Position. Beispiele für die unterschiedliche Ausrichtung:
- Vertrieb/Vertriebsleitung: Kundengewinnung und -bindung, Abschluss- und \
Verhandlungsstärke, Umsatz-/Ergebnisverantwortung, Markt- und Wettbewerbskenntnis, \
Pipeline- und Gebietsentwicklung.
- Montageleitung/Montage: Steuerung von Baustellen und Montageprojekten, Führung und \
Koordination von Montageteams, Termin-, Kosten- und Qualitätskontrolle, \
Arbeitssicherheit, Schnittstelle zu Technik und Kunde.
- Service/Servicetechnik: technische Problemlösung, Wartung/Instandhaltung, \
Reaktionszeiten und Servicequalität, Kundenorientierung im Einsatz.
Übertrage dieses Prinzip sinngemäß auf die tatsächlich erkannte Position.

Bleibe dabei faktentreu: leite nur ab, was aus Interview und Lebenslauf hervorgeht. \
Übernimm die Anforderungen der Position NICHT als angebliche Eigenschaften des \
Kandidaten, wenn sie im Material nicht belegt sind -- prüfe die Passung ehrlich."""


def _ziel_context_block() -> str:
    """Anweisung, das Profil inhaltlich auf die beschriebene Zielposition zu
    beziehen. Wird nur eingebunden, wenn eine Beschreibung vorliegt (die
    Beschreibung selbst steht am Ende der Quellen im User-Content)."""
    return ("ZIELPOSITION: Am Ende der Quellen steht eine Beschreibung der "
            "Zielposition (Tätigkeit, Firma, Aufgaben, Anforderungen). Beziehe das "
            "Profil inhaltlich konkret auf diese Zielposition -- insbesondere "
            "\"Verständnis der Zielrolle\" und die Gesamteinschätzung sollen die "
            "Passung des Kandidaten zu genau dieser Position begründen. Nutze die "
            "Beschreibung ausschließlich als Zielkontext; erfinde keine Fakten über "
            "den Kandidaten und übernimm keine Anforderungen als angebliche "
            "Eigenschaften des Kandidaten, wenn sie nicht im Material belegt sind.")


def _build_drafting_base(rating, zielposition, blind=False, zielbeschreibung=""):
    """Gemeinsamer System-Prompt-Kopf für die Abschnitts-Erzeugung: Grundregeln +
    (optional) Anonymisierung + Bewertung + (optional) Zielposition + Zielkontext.
    Die abschnittsspezifische Anweisung wird pro Aufruf ergänzt (siehe draft_sections)."""
    parts = [DRAFTING_BASE_RULES]
    if blind:
        parts.append(_blind_drafting_block())
    parts.append(_rating_block(rating))
    if zielposition:
        parts.append(_position_block(zielposition))
    if zielbeschreibung and zielbeschreibung.strip():
        parts.append(_ziel_context_block())
    return "\n\n".join(parts)

PERSONAL_DATA_KEYS = [
    "nachname", "vorname", "geburtsdatum", "familienstand", "wohnort", "mobil",
    "email", "kuendigungsfrist", "ist_gehalt", "gehaltswunsch", "interviewpartner",
    "rolle", "rolle_kurz",
]
SECTION_KEYS = [e["key"] for e in config.PROFILE_SECTIONS]

PERSONAL_DATA_SCHEMA = {
    "type": "object",
    "properties": {k: {"type": "string"} for k in PERSONAL_DATA_KEYS},
    "required": PERSONAL_DATA_KEYS,
    "additionalProperties": False,
}

SECTIONS_SCHEMA = {
    "type": "object",
    "properties": {k: {"type": "string"} for k in SECTION_KEYS},
    "required": SECTION_KEYS,
    "additionalProperties": False,
}

# Anonymisierte Personendaten-Felder für das Blindprofil (siehe
# config.BLIND_PERSONAL_DATA_FIELDS) plus die auch dort benötigten
# Kontext-Felder interviewpartner/rolle/rolle_kurz.
BLIND_PERSONAL_DATA_KEYS = [
    "jahrgang", "region", "branche", "unternehmensgroesse", "familienstand",
    "kuendigungsfrist", "ist_gehalt", "gehaltswunsch", "verfuegbarkeit",
    "interviewpartner", "rolle", "rolle_kurz",
]
BLIND_PERSONAL_DATA_SCHEMA = {
    "type": "object",
    "properties": {k: {"type": "string"} for k in BLIND_PERSONAL_DATA_KEYS},
    "required": BLIND_PERSONAL_DATA_KEYS,
    "additionalProperties": False,
}


def _blind_drafting_block() -> str:
    """Anonymisierungs-Anweisung für die Fließtexte im Blindprofil."""
    return """ANONYMISIERUNG (BLINDPROFIL): Dieses Profil muss vollständig anonym \
sein -- aus dem Text darf sich kein Rückschluss auf die konkrete Person ziehen \
lassen, obwohl der Kandidat fachlich VOLLUMFÄNGLICH beschrieben wird.
- Nenne NIEMALS den Namen des Kandidaten. Bezeichne die Person neutral als "die \
Kandidatin" bzw. "der Kandidat" (nutze das aus dem Material erkennbare Geschlecht; \
ist es unklar, schreibe "die Kandidatin/der Kandidat").
- Nenne KEINE Firmennamen, Markennamen, konkreten Orte oder Städte, Projektnamen \
oder sonstige eindeutig identifizierenden Details. Ersetze sie durch allgemeine \
Beschreibungen (z.B. "ein mittelständischer Maschinenbauer", "im süddeutschen \
Raum", "ein namhafter Wettbewerber").
- Beschreibe Aufgaben, Kompetenzen, Erfolge und Persönlichkeit dennoch vollständig \
und aussagekräftig -- die Anonymisierung darf die fachliche Tiefe nicht schmälern."""


def _call_claude_structured(client, system_prompt, user_content, schema, max_tokens):
    """Ruft Claude mit erzwungenem JSON-Schema auf (Structured Outputs).
    Die Antwort ist garantiert gültiges JSON gemäß 'schema' -- kein
    Freitext-Parsing, kein Codeblock-Handling nötig.

    Bei transienten Fehlern (Überlastung 529, Rate-Limit 429, 5xx,
    Verbindungsabbruch) wird automatisch mit wachsender Wartezeit erneut
    versucht (2s, 4s, 8s, 16s). Nicht behebbare Fehler (z.B. ungültiger
    Key 401, fehlendes Guthaben) werden sofort weitergereicht."""
    last_err = None
    for attempt in range(_MAX_ATTEMPTS):
        try:
            message = client.messages.create(
                model=MODEL,
                max_tokens=max_tokens,
                system=system_prompt,
                messages=[{"role": "user", "content": user_content}],
                output_config={"format": {"type": "json_schema", "schema": schema}},
            )
            text_block = next(b for b in message.content if b.type == "text")
            try:
                return json.loads(text_block.text)
            except json.JSONDecodeError as je:
                # Antwort kam unvollständig/abgeschnitten an (kommt bei
                # Serverüberlastung vor). Bis zum vorletzten Versuch: komplett
                # neu anfordern. Beim letzten Versuch: so viel wie möglich
                # aus der Teil-Antwort retten (json_repair).
                last_err = je
                if attempt == _MAX_ATTEMPTS - 1 and json_repair is not None:
                    repaired = json_repair.loads(text_block.text)
                    if isinstance(repaired, dict) and repaired:
                        return repaired
                if attempt == _MAX_ATTEMPTS - 1:
                    raise
        except anthropic.APIConnectionError as e:
            last_err = e  # Netzwerk-/Verbindungsproblem -> erneut versuchen
        except anthropic.APIStatusError as e:
            status = getattr(e, "status_code", None)
            if status not in _RETRYABLE_STATUS:
                raise  # z.B. 401/400/402 -> sofort melden, kein Retry
            last_err = e
        if attempt < _MAX_ATTEMPTS - 1:
            time.sleep(2 * (2 ** attempt))  # 2, 4, 8, 16 Sekunden
    raise last_err


def _call_claude_text(client, system_prompt, user_content, max_tokens):
    """Ruft Claude für einen einzelnen Fließtext auf (kein JSON, kein Schema).
    Gleiche automatische Wiederholung bei transienten Fehlern wie
    _call_claude_structured; nicht behebbare Fehler (401/400/402) sofort weiter."""
    last_err = None
    for attempt in range(_MAX_ATTEMPTS):
        try:
            message = client.messages.create(
                model=MODEL,
                max_tokens=max_tokens,
                system=system_prompt,
                messages=[{"role": "user", "content": user_content}],
            )
            return "".join(b.text for b in message.content if b.type == "text")
        except anthropic.APIConnectionError as e:
            last_err = e
        except anthropic.APIStatusError as e:
            if getattr(e, "status_code", None) not in _RETRYABLE_STATUS:
                raise
            last_err = e
        if attempt < _MAX_ATTEMPTS - 1:
            time.sleep(2 * (2 ** attempt))
    raise last_err


def extract_personal_data(api_key: str, transcript_text: str, cv_text: str = "",
                          blind: bool = False) -> dict:
    """Extrahiert die Personendaten. blind=True liefert stattdessen die
    anonymisierten Blindprofil-Felder (siehe BLIND_EXTRACTION_RULES)."""
    client = anthropic.Anthropic(api_key=api_key)
    user_content = f"INTERVIEWPROTOKOLL:\n{transcript_text}\n\nLEBENSLAUF:\n{cv_text}"
    rules = BLIND_EXTRACTION_RULES if blind else EXTRACTION_CONTENT_RULES
    schema = BLIND_PERSONAL_DATA_SCHEMA if blind else PERSONAL_DATA_SCHEMA
    return _call_claude_structured(
        client, rules, user_content, schema, max_tokens=1024,
    )


def draft_sections(api_key: str, transcript_text: str, cv_text: str,
                    target_pages: float, rating: int = 5,
                    zielposition: str = "", section_keys=None,
                    blind: bool = False, zielbeschreibung: str = "") -> dict:
    """Entwirft die Fließtexte.

    rating: 1-10 (steuert Tonfall/Gewichtung, nie die Fakten -- siehe _rating_block)
    zielposition: gewählte Zielposition-Kategorie (dezenter Stil-Einfluss) oder ""
    section_keys: Liste der zu erzeugenden Abschnitts-Keys; None = alle
    blind: True erzeugt anonymisierte Texte (keine Namen/Firmen/Orte)
    zielbeschreibung: freie Beschreibung der Zielposition (Tätigkeit/Firma/Aufgaben);
        wird als Zielkontext an die Quellen angehängt, das Profil bezieht sich darauf
    """
    client = anthropic.Anthropic(api_key=api_key)
    if not section_keys:
        section_keys = list(SECTION_KEYS)
    target_words = round(target_pages * config.WORDS_PER_PAGE_ESTIMATE)
    per_section = max(70, round(target_words / max(1, len(section_keys))))
    base = _build_drafting_base(rating, zielposition, blind=blind,
                                zielbeschreibung=zielbeschreibung)

    sources = f"INTERVIEWPROTOKOLL:\n{transcript_text}\n\nLEBENSLAUF:\n{cv_text}"
    if zielbeschreibung and zielbeschreibung.strip():
        sources += ("\n\n===\nBESCHREIBUNG DER ZIELPOSITION "
                    "(Tätigkeit/Firma/Aufgaben/Anforderungen):\n" + zielbeschreibung.strip())

    # Ein eigener Aufruf pro Abschnitt: Das Modell kann so nicht mehr alle Inhalte in
    # ein Feld packen oder Überschriften in den Text schreiben, und lange Antworten
    # werden nicht abgeschnitten. Für die Wiederholungsfreiheit bekommt jeder Aufruf
    # die bereits verfassten Abschnitte als "nicht wiederholen"-Kontext mit.
    result = {}
    written = []  # (Überschrift, Text) der schon erzeugten Abschnitte
    for key in section_keys:
        entry = config.ALL_SECTIONS_BY_KEY.get(key, {"heading": key, "desc": ""})
        heading = entry.get("heading", key)
        desc = entry.get("desc", "")
        sec_rules = (
            f'Schreibe AUSSCHLIESSLICH den Fließtext für den Abschnitt „{heading}".\n'
            f'Inhalt und Abgrenzung dieses Abschnitts: {desc}\n'
            'Gib NUR den Inhalt dieses einen Abschnitts zurück -- KEINE Überschrift, '
            'keinen Rubriknamen und keinen Vorspann wie "In diesem Abschnitt". Trenne '
            'Absätze durch eine Leerzeile. '
        )
        if entry.get("bullets"):
            # In passenden Rubriken eine kurze Aufzählung erlauben -- lockert auf.
            sec_rules += (
                'Setze GENAU EINE kurze Aufzählung ein, wo es sich anbietet, konkret für '
                f'{entry.get("bullet_hint", "die zentralen Punkte")}: jeder Punkt auf '
                'einer eigenen Zeile, beginnend mit "- ". Leite die Liste mit einem '
                'einordnenden Satz ein und rahme sie mit Fließtext (davor und, wenn '
                'sinnvoll, danach). Hältst du eine Liste hier nicht für passend, schreibe '
                'reinen Fließtext. Kein weiteres Markdown (keine *, #, Überschriften). '
            )
        else:
            sec_rules += (
                'Reiner Fließtext -- KEINE Aufzählungszeichen und kein Markdown '
                '(keine -, *, #). '
            )
        sec_rules += (
            f'Zielumfang etwa {per_section} Wörter (grober Richtwert; Faktentreue und '
            'Qualität haben Vorrang).'
        )
        system_prompt = base + "\n\n" + sec_rules
        user_content = sources
        if written:
            vorher = "\n\n".join(f"[{h}]\n{t}" for h, t in written)
            user_content += (
                "\n\n===\nBEREITS VERFASSTE ABSCHNITTE (ihre Inhalte NICHT wiederholen "
                "-- weder wörtlich noch sinngemäß; schreibe hier nur, was zu diesem "
                "Abschnitt gehört):\n" + vorher)
        text = _call_claude_text(client, system_prompt, user_content, max_tokens=2000).strip()
        # Schutz gegen leere/abgebrochene Abschnitte: Gibt das Modell für einen
        # Abschnitt (praktisch) nichts zurück, wird der Aufruf bis zu zweimal
        # wiederholt -- sonst erschiene die Rubrik-Überschrift ohne Inhalt.
        for _retry in range(2):
            if len(text) >= 40:
                break
            text = _call_claude_text(
                client, system_prompt, user_content, max_tokens=2000).strip()
        result[key] = text
        written.append((heading, text))
    return result
