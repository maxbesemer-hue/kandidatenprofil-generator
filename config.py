"""
Zentrale Konfiguration für den Kandidatenprofil-Generator.

Alle Design-Werte (Farben, Schriftgrößen, Ränder, Logo-Positionen) wurden
1:1 aus der hochgeladenen Vorlage 'Kandidatenprofil_Jaqueline_Keck.docx'
ausgelesen (word/styles.xml, word/document.xml, word/header1.xml).
Quelle je Wert ist in den Kommentaren angegeben, damit nachvollziehbar bleibt,
woher jede Zahl stammt und nichts geschätzt wurde.
"""

from docx.shared import Pt, Inches, Emu

# -----------------------------------------------------------------------
# AGENTUR-STAMMDATEN (Cover-Seite und Kopf-/Fußzeile)
# Quelle: word/document.xml, Cover-Textblock der Vorlage
# -----------------------------------------------------------------------
AGENCY_NAME = "Vertico Executive Search GmbH"
AGENCY_ADDRESS_LINE1 = "Ahornweg 36"
AGENCY_ADDRESS_LINE2 = "73278 Schlierbach"
AGENCY_WEBSITE = "www.vertico-search.de"
AGENCY_EMAIL = "info@vertico-search.de"

# -----------------------------------------------------------------------
# SEITENLAYOUT
# Quelle: word/document.xml -> <w:pgSz w:w="12240" w:h="15840"/>
#         und <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800".../>
# 12240/1440 = 8.5 Zoll, 15840/1440 = 11 Zoll => US Letter (nicht A4!)
# -----------------------------------------------------------------------
PAGE_WIDTH = Inches(12240 / 1440)   # 8.5"
PAGE_HEIGHT = Inches(15840 / 1440)  # 11"
MARGIN_TOP = Inches(1440 / 1440)    # 1.0"
MARGIN_BOTTOM = Inches(1440 / 1440) # 1.0"
MARGIN_LEFT = Inches(1800 / 1440)   # 1.25"
MARGIN_RIGHT = Inches(1800 / 1440)  # 1.25"

# -----------------------------------------------------------------------
# SCHRIFT
# KORRIGIERT nach genauer XML-Prüfung: Der Fließtext der Vorlage verwendet
# NICHT den docDefaults-Wert (sz=22=11pt), sondern durchgängig den Absatzstil
# "StandardWeb" (styles.xml: w:styleId="StandardWeb") mit sz="24" (=12pt).
# Das gilt auch für den Block "Persönliche Daten" (dort zusätzlich direkt
# im Absatz auf sz="24" gesetzt). Quelle: word/document.xml, Absatz mit
# pStyle="StandardWeb" bzw. dem Persönliche-Daten-Absatz.
# -----------------------------------------------------------------------
FONT_NAME = "Calibri"
BODY_SIZE = Pt(12)          # korrigiert von 11pt auf 12pt (siehe oben)

# Überschrift 1 (z.B. "Persönliche Daten", "Hintergrund")
# Quelle: styles.xml Style "berschrift1": color 365F91, sz="28" (=14pt), bold,
# spacing before=480 twips (24pt). Der "after"-Wert der Vorlage selbst ist 0,
# der sichtbare Abstand zum folgenden Text entsteht dort stattdessen durch
# eine variable "auto"-Absatzabstand-Berechnung von Word, die keinen festen,
# reproduzierbaren Punktwert ergibt. Für GARANTIERT identischen Abstand nach
# jeder Überschrift wird hier stattdessen ein fester, expliziter Wert
# gesetzt (siehe auch docx_builder._add_body_paragraphs, das space_before
# bei jedem Fließtext-Absatz konsequent auf 0 setzt -- der Abstand zwischen
# Überschrift und Text wird dadurch ausschließlich hierdurch bestimmt).
H1_COLOR = "365F91"
H1_SIZE = Pt(14)
H1_SPACE_BEFORE = Pt(24)
H1_SPACE_AFTER = Pt(8)

# Überschrift 2 (z.B. "Führung und Mitarbeiterentwicklung")
# Quelle: styles.xml Style "berschrift2": color 4F81BD, sz="26" (=13pt), bold,
# spacing before=200 twips (10pt). "after" siehe Erläuterung bei H1 oben.
H2_COLOR = "4F81BD"
H2_SIZE = Pt(13)
H2_SPACE_BEFORE = Pt(10)
H2_SPACE_AFTER = Pt(6)

# Fließtext-Absätze: Blocksatz. Quelle: word/document.xml -- im gesamten
# Fließtext wird durchgängig <w:jc w:val="both"/> verwendet (72 Fundstellen),
# nirgends linksbündig.
BODY_JUSTIFY = True

# Abstände rund um Aufzählungen (List-Bullet-Absätze). Die Liste soll sich
# sichtbar vom umgebenden Fließtext absetzen (Luft davor und danach), die
# einzelnen Punkte aber eng zusammenstehen, damit die Liste als eine Einheit
# gelesen wird.
BULLET_SPACE_BEFORE = Pt(10)    # Luft zwischen Einleitungssatz und erstem Punkt
BULLET_SPACE_BETWEEN = Pt(8)    # Abstand zwischen zwei Punkten (Lesbarkeit)
BULLET_SPACE_AFTER = Pt(12)     # Luft zwischen letztem Punkt und Folgeabsatz

# Titel auf Cover-Seite "Kandidatenprofil"
# KORRIGIERT nach exaktem XML-Fund (direkte Formatierung des Titel-Absatzes,
# nicht der undefinierte "Titel"-Formatvorlagenwert): sz="72" = 36pt, bold,
# zentriert, Schriftart majorHAnsi-Theme (= Calibri, siehe theme1.xml).
COVER_TITLE_SIZE = Pt(36)
# Reservierter Abstand vor dem Titel, damit er unterhalb des freischwebenden
# Logos beginnt (das Logo nimmt selbst keinen Platz im Textfluss ein).
# Herleitung: Logo-Unterkante liegt bei ca. 3.49" von der Papierkante
# (0.083" Versatz über dem oberen Seitenrand + 3.406" Logohöhe), der
# reguläre Textfluss beginnt am oberen Seitenrand (1"). Reservierter
# Abstand ≈ (3.49 - 1) × 72 ≈ 179pt; per Render-Test (soffice) auf einen
# sauberen, überlappungsfreien Übergang feinjustiert.
COVER_TITLE_SPACE_BEFORE = Pt(179)
# Kandidatenname: sz="44" = 22pt, bold. Quelle: Absatz mit "Jaqueline Keck".
COVER_NAME_SIZE = Pt(22)
# Rollen-Untertitel (in der Vorlage per <w:br/> im selben Absatz wie der
# Name): sz="32" = 16pt, nicht fett. Gleiche Quelle.
COVER_ROLE_SIZE = Pt(16)
COVER_CONTACT_SIZE = Pt(10)

# -----------------------------------------------------------------------
# KOPFZEILE: Logo oben rechts auf allen Seiten außer Deckblatt
# Quelle: word/header1.xml -> <wp:extent cx="1059180" cy="1059180"/>
# 1059180 EMU / 914400 EMU-pro-Zoll = 1.1583... Zoll (quadratisch)
# Position: posOffset x=5455920 EMU, y=-502920 EMU (behindDoc, floating)
# -----------------------------------------------------------------------
HEADER_LOGO_SIZE = Emu(1059180)          # ≈ 1.158" x 1.158"
HEADER_LOGO_ASSET = "assets/logo_header.png"

# Position des freischwebenden Logos (relativ zur Textspalte bzw. zum
# Absatz). Quelle: word/header1.xml der Vorlage,
# <wp:positionH relativeFrom="column"><wp:posOffset>5455920</wp:posOffset>
# <wp:positionV relativeFrom="paragraph"><wp:posOffset>-502920</wp:posOffset>
# Diese Werte gelten nur exakt, weil Seitengröße und Ränder unseres
# Dokuments identisch zur Vorlage sind (siehe PAGE_WIDTH/MARGIN_* oben) --
# bei abweichender Seitengeometrie müssten sie neu berechnet werden.
HEADER_LOGO_OFFSET_X = 5455920
HEADER_LOGO_OFFSET_Y = -502920

# -----------------------------------------------------------------------
# DECKBLATT-LOGO
# Quelle: word/document.xml Cover-Bild, image1.png (2048x2048 px, quadratisch)
# In der Vorlage auf Breite 3.40625" skaliert (aus dem Markdown-Export
# 'width="3.40625in" height="3.40625in"' ersichtlich).
# -----------------------------------------------------------------------
COVER_LOGO_ASSET = "assets/logo_cover.png"
COVER_LOGO_WIDTH = Inches(3.40625)

# Position des freischwebenden Deckblatt-Logos (wie beim Kopfzeilen-Logo:
# frei schwebend statt im normalen Textfluss, dadurch bewusst weiter oben
# platziert, bis leicht über den regulären oberen Seitenrand hinausragend).
# Quelle: word/document.xml, allererste Bild-Verankerung im Dokument,
# <wp:positionH relativeFrom="column"><wp:posOffset>1193269</wp:posOffset>
# <wp:positionV relativeFrom="paragraph"><wp:posOffset>-838200</wp:posOffset>
# Gilt exakt, weil Seitengröße/Ränder identisch zur Vorlage sind.
COVER_LOGO_OFFSET_X = 1193269
COVER_LOGO_OFFSET_Y = -838200

# Maximale Größe des Kandidatenfotos auf dem Deckblatt ("contain"-Skalierung,
# Seitenverhältnis bleibt erhalten, das Foto überschreitet diesen Rahmen nie).
# Bewusst kleiner als der ursprüngliche Näherungswert (2.2in Breite ohne
# Höhenbegrenzung) gewählt: Ein hochformatiges Foto konnte zusammen mit den
# früheren 6 Leerabsätzen das Deckblatt auf eine zweite Seite drücken. Mit
# fester Maximalhöhe ist die Gesamthöhe des Deckblatts unabhängig vom
# Seitenverhältnis des hochgeladenen Fotos vorhersehbar.
# Maximale Größe des Kandidatenfotos auf dem Deckblatt ("contain"-Skalierung,
# Seitenverhältnis bleibt erhalten, das Foto überschreitet diesen Rahmen nie).
# Quelle für die Proportion: Das Originalfoto in der Vorlage ist exakt auf
# cx="1693795" cy="2023673" EMU gesetzt = 1.852in x 2.213in (word/document.xml,
# Bildanker bei "Jaqueline Keck"). Das ist allerdings eine feste, für dieses
# eine Foto von Hand gewählte Größe (kein Vorlagen-Rahmenwert) -- bei
# automatisch generierten Profilen mit wechselnden Fotoformaten muss die
# Größe stattdessen dynamisch berechnet werden, um Überlauf zu vermeiden
# (siehe _fit_photo_size). Als Rahmen wird deshalb die gleiche Proportion,
# aber mit Sicherheitsabstand zur Seitenkante verwendet.
COVER_PHOTO_MAX_WIDTH = Inches(1.85)
COVER_PHOTO_MAX_HEIGHT = Inches(1.8)

# Fester Abstand vor dem Agentur-Kontaktblock auf dem Deckblatt. Ersetzt die
# vormals 6 aneinandergereihten Leerabsätze (deren tatsächliche Höhe vom
# Zeilenabstand der Vorlage abhing und in der Praxis zu groß war -- das war
# die Ursache für den Seitenüberlauf). Empirisch mit dem Render-Test
# (soffice --convert-to pdf) auf eine Seite kalibriert -- nach der Korrektur
# der Schriftgrößen (Titel 36pt/Name 22pt/Rolle 16pt, siehe oben) musste
# dieser Wert nochmals leicht reduziert werden, da die größeren Schriften
# selbst schon mehr Platz brauchen.
COVER_CONTACT_SPACE_BEFORE = Pt(6)

# Zeilenabstand im Block "Persönliche Daten": Quelle word/document.xml,
# <w:spacing w:line="360" w:lineRule="auto"/> = 360/240 = 1.5-facher
# Zeilenabstand.
PERSONAL_DATA_LINE_SPACING = 1.5

# Fester Tabstopp zwischen Label ("Nachname:") und Wert ("Keck") im Block
# "Persönliche Daten". Die Vorlage selbst verwendet zwei aufeinanderfolgende
# Standard-Tabs ohne explizit definierten Tabstopp -- das funktioniert nur
# zufällig sauber, weil Word/die Vorlage darauf ausgelegt war. Bei variabler
# Labellänge (z.B. "Mobil:" vs. "Kündigungsfrist:") springen zwei
# Standard-Tabs je nach Ausgangsposition unterschiedlich weit, wodurch die
# Werte NICHT untereinander ausgerichtet sind. Ein einzelner, explizit
# definierter Tabstopp an einer festen Position behebt das zuverlässig.
# Position so gewählt, dass sie bequem hinter dem längsten Label
# "Kündigungsfrist:" (bei 12pt fett Calibri) liegt.
PERSONAL_DATA_TAB_STOP = Inches(1.8)

# -----------------------------------------------------------------------
# FESTE ABSCHNITTSSTRUKTUR DES KANDIDATENPROFILS
# Quelle: Gliederung der Vorlage (word/document.xml Heading-Reihenfolge).
# "heading" ist ein fester Text; {rolle} wird zur Laufzeit durch die
# Zielposition ersetzt. "level" 1 = Überschrift 1, 2 = Überschrift 2
# (nur bei "Führung und Mitarbeiterentwicklung" als Unterabschnitt von
# "Eignung für ...").
# -----------------------------------------------------------------------
# Struktur nach der Referenzvorlage "Kandidatenprofil_Anja_Krückeberg.docx":
# sechs Rubriken, alle als Überschrift 1. "desc" beschreibt Inhalt UND Abgrenzung
# jeder Rubrik und wird der KI übergeben, damit sich die Abschnitte inhaltlich
# nicht überschneiden (siehe ai_extraction._section_guide). "heading" ist ein
# fester Text; {rolle} wird zur Laufzeit durch die Zielposition ersetzt, falls
# vorhanden.
#
# Positionsabhängig: An vierter Stelle steht bei einer FÜHRUNGSPOSITION die Rubrik
# "Führungsverständnis", bei einer FACHPOSITION (ohne Führung) stattdessen
# "Fachliche Kompetenzen und Arbeitsweise". Alle übrigen Rubriken sind identisch.
_SECTION_WERDEGANG = {
    "key": "werdegang", "heading": "Beruflicher Werdegang", "level": 1,
    "bullets": True,
    "bullet_hint": "die wichtigsten beruflichen Stationen (chronologisch, je Station "
                   "eine kurze Zeile)",
    "desc": "Ausbildung, bisherige berufliche Stationen und Aufstieg, Jahre an "
            "Branchen-/Vertriebserfahrung, Führungs- und P&L-Verantwortung (falls "
            "vorhanden), fachliche Kenntnisse und eingesetzte Systeme/Tools, Sprachen "
            "sowie regionale Verankerung und über die Jahre gewachsenes Netzwerk. "
            "ABGRENZUNG: hier keine Wechselmotivation, keine ausführliche Bewertung des "
            "Arbeits-/Führungsstils und keine Gesamtbewertung."}
_SECTION_SITUATION = {
    "key": "situation_motivation", "heading": "Aktuelle Situation und Wechselmotivation",
    "level": 1,
    "desc": "Aktuelle berufliche Situation, konkreter Auslöser/Anlass für den "
            "Wechselgedanken, Wechselbereitschaft und die inhaltlichen Beweggründe, "
            "sowie die konkrete Motivation für die Zielrolle bzw. das Zielunternehmen "
            "(inkl. eigener Auseinandersetzung/Vorbereitung, falls belegt). ABGRENZUNG: "
            "den Werdegang nicht erneut erzählen; nicht das Rollenverständnis ausführen "
            "(nur der Antrieb gehört hierher)."}
_SECTION_ZIELROLLE = {
    "key": "zielrolle_verstaendnis", "heading": "Verständnis der Zielrolle", "level": 1,
    "desc": "Wie die Kandidatin/der Kandidat die Anforderungen der Zielposition "
            "versteht und einordnet: fachlich-strategisches Verständnis des "
            "Geschäftsmodells und ggf. eine eigenständig erarbeitete inhaltliche/"
            "strategische Analyse für die Rolle oder das Unternehmen. ABGRENZUNG: hier "
            "geht es um das Verständnis (Was/Wie), nicht um die Beweggründe (Warum) und "
            "nicht um Führung."}
# Vierte Rubrik – Führungsposition
_SECTION_FUEHRUNG = {
    "key": "fuehrungsverstaendnis", "heading": "Führungsverständnis", "level": 1,
    "bullets": True,
    "bullet_hint": "konkrete Beispiele der Führungspraxis im Alltag",
    "desc": "Führungsstil und -haltung, Umgang mit dem Team, Entwicklung der "
            "Mitarbeitenden, Verhalten bei Zielkonflikten oder unpopulären Vorgaben, "
            "mit konkreten Beispielen aus der Praxis. ABGRENZUNG: ausschließlich "
            "Führung; keine fachlichen oder Werdegangs-Fakten und keine Gesamtbewertung."}
# Vierte Rubrik – Fachposition (ohne Führung), Ersatz für "Führungsverständnis"
_SECTION_FACHKOMPETENZ = {
    "key": "fachkompetenz", "heading": "Fachliche Kompetenzen und Arbeitsweise", "level": 1,
    "bullets": True,
    "bullet_hint": "die zentralen fachlichen Kompetenzen und Schwerpunkte",
    "desc": "Tiefe und Anwendungspraxis der fachlichen und methodischen Kompetenzen, "
            "Spezialgebiete, Qualitäts- und Zuverlässigkeitsorientierung, "
            "Selbstorganisation und Arbeitsweise sowie Zusammenarbeit im Team ohne "
            "Führungsverantwortung, mit konkreten Beispielen. ABGRENZUNG: keine "
            "Führungsaufgaben; die im Werdegang genannten Stationen und Systeme nicht "
            "erneut aufzählen (hier geht es um das Wie und die Tiefe des fachlichen "
            "Arbeitens, nicht um die Chronologie); keine Gesamtbewertung."}
_SECTION_WIRKUNG = {
    "key": "persoenliche_wirkung", "heading": "Persönliche Wirkung", "level": 1,
    "desc": "Persönlicher Eindruck aus dem Interview: Auftreten, Kommunikationsstil, "
            "Reflexions- und Selbstwahrnehmung, Authentizität, Resilienz und "
            "Vereinbarkeit von Beruf und Privatem. ABGRENZUNG: kein Führungs- oder "
            "Arbeitsstil im fachlichen Sinn, keine fachlichen Fakten, keine Empfehlung."}
_SECTION_GESAMT = {
    "key": "gesamteinschaetzung", "heading": "Gesamteinschätzung", "level": 1,
    "desc": "Wertende Gesamtbeurteilung mit klarer Empfehlung und Aussage zur Eignung/"
            "Passung für die Zielrolle. Das ist eine Schlussfolgerung, KEINE "
            "Zusammenfassung: formuliere das übergreifende Urteil, ohne die Fakten, "
            "Beispiele und Details der vorherigen Abschnitte erneut aufzuzählen."}


def get_profile_sections(is_leadership=True):
    """Rubriken je nach Positionstyp. Führungsposition -> mit 'Führungsverständnis';
    Fachposition -> stattdessen 'Fachliche Kompetenzen und Arbeitsweise'."""
    vierte = _SECTION_FUEHRUNG if is_leadership else _SECTION_FACHKOMPETENZ
    return [_SECTION_WERDEGANG, _SECTION_SITUATION, _SECTION_ZIELROLLE, vierte,
            _SECTION_WIRKUNG, _SECTION_GESAMT]


# Kanonische Reihenfolge ALLER möglichen Rubriken (beide Positionstypen). Wird zum
# Rendern/Nachschlagen genutzt: Filtert man nach den aktiven Keys, ergibt sich immer
# die richtige Reihenfolge, egal welcher Positionstyp gewählt ist (es ist immer nur
# eine der beiden vierten Rubriken aktiv).
ALL_SECTIONS_ORDERED = [_SECTION_WERDEGANG, _SECTION_SITUATION, _SECTION_ZIELROLLE,
                        _SECTION_FUEHRUNG, _SECTION_FACHKOMPETENZ, _SECTION_WIRKUNG,
                        _SECTION_GESAMT]
ALL_SECTIONS_BY_KEY = {e["key"]: e for e in ALL_SECTIONS_ORDERED}

# Standard = Führungsposition (Rückwärtskompatibilität für Code, der die
# Modulkonstante direkt nutzt).
PROFILE_SECTIONS = get_profile_sections(True)

# Hinweis: Die frühere feste Positionsliste (Dropdown) wurde entfernt. Die
# Zielposition wird jetzt in der App als freie, konkrete Stellenbezeichnung
# eingegeben; die KI erkennt daraus die Art der Position und richtet das Profil
# inhaltlich darauf aus (siehe ai_extraction._position_block).

# Felder im Block "Persönliche Daten" -- Reihenfolge und Labels exakt wie
# in der Vorlage (word/document.xml, Abschnitt "Persönliche Daten").
PERSONAL_DATA_FIELDS = [
    ("nachname", "Nachname:"),
    ("vorname", "Vorname:"),
    ("geburtsdatum", "Geburtsdatum:"),
    ("familienstand", "Familienstand:"),
    ("wohnort", "Wohnort:"),
    ("mobil", "Mobil:"),
    ("email", "E-Mail:"),
    ("kuendigungsfrist", "Kündigungsfrist:"),
    ("ist_gehalt", "Ist-Gehalt:"),
    ("gehaltswunsch", "Gehaltswunsch:"),
]

# Personendaten-Felder im BLINDPROFIL: bewusst nur nicht-identifizierende
# Angaben, aus denen sich kein Rückschluss auf die konkrete Person ziehen
# lässt (kein Name, keine Adresse, kein Kontakt, kein Arbeitgebername).
# Statt genauem Wohnort nur eine Region, statt Arbeitgeber nur Branche und
# Unternehmensgröße. Kündigungsfrist, Gehalt und Verfügbarkeit bleiben laut
# Vorgabe erhalten.
BLIND_PERSONAL_DATA_FIELDS = [
    ("jahrgang", "Jahrgang:"),
    ("region", "Region:"),
    ("branche", "Branche:"),
    ("unternehmensgroesse", "Unternehmensgröße:"),
    ("familienstand", "Familienstand:"),
    ("kuendigungsfrist", "Kündigungsfrist:"),
    ("ist_gehalt", "Ist-Gehalt:"),
    ("gehaltswunsch", "Gehaltswunsch:"),
    ("verfuegbarkeit", "Verfügbarkeit:"),
]

# -----------------------------------------------------------------------
# LÄNGENSTEUERUNG
# Nachrechenbare Herleitung dieses Richtwerts (keine Schätzung ins Blaue):
# In der Beispieldatei 'Kandidatenprofil_Jaqueline_Keck.docx' wurde der
# Fließtext ab der Überschrift "Hintergrund" bis zum Dokumentende gezählt:
# exakt 1531 Wörter (re.findall(r'\w+', ...) auf den Pandoc-Textexport).
# Dieser Fließtext erstreckt sich von der unteren Hälfte von Seite 2 bis
# zum Ende von Seite 8 der PDF-Konvertierung, also über ca. 6,5 Seiten.
# 1531 Wörter / 6,5 Seiten = 235,5 Wörter/Seite (gerundet 236).
# WICHTIG: Dies bleibt ein Richtwert, kein exaktes Maß -- der tatsächliche
# Seitenumbruch in Word hängt zusätzlich von Silbentrennung, Schrift-
# rendering und eingebetteten Bildern ab und ist vorab nicht exakt
# berechenbar. Die UI weist diesen Wert deshalb explizit als Näherung aus.
WORDS_PER_PAGE_ESTIMATE = 236

# -----------------------------------------------------------------------
# BRANDING-PROFILE (Auswahl in der App: Vertico oder Partnerfirma VIP)
# Bestimmt Logo, Firmierung, Adresse, Kontaktblock und die Farbe des
# radialen Deckblatt-Verlaufs.
# -----------------------------------------------------------------------
# VIP-Markenfarbe: dominantes Grün, aus 'Logos/Firmen Logo VIP .png' gesampelt.
VIP_BRAND_COLOR = "B4C13E"

# Größe des VIP-Icons auf dem Deckblatt (kleiner als das große Vertico-Logo,
# da das Icon niedrig aufgelöst ist -- daneben steht der Firmenschriftzug).
COVER_ICON_SIZE = Inches(1.05)
COVER_WORDMARK_SIZE = Pt(20)
# Abstand vor dem Titel, wenn das Logo NICHT freischwebend, sondern als
# Icon+Schriftzug im normalen Textfluss oben steht (VIP-Variante).
COVER_TITLE_SPACE_BEFORE_INLINE = Pt(18)

BRANDS = {
    "vertico": {
        "label": "Vertico Executive Search",
        "name": AGENCY_NAME,
        "address": [AGENCY_ADDRESS_LINE1, AGENCY_ADDRESS_LINE2],
        "contact": [AGENCY_WEBSITE, AGENCY_EMAIL],
        "cover_logo": COVER_LOGO_ASSET,      # großes, hochauflösendes Logo
        "header_logo": HEADER_LOGO_ASSET,
        "logo_mode": "image",                # großes freischwebendes Logo (wie Vorlage)
        "wordmark": None,
        "gradient_color": H2_COLOR,          # 4F81BD (Vertico-Blau)
    },
    "vip": {
        "label": "VIP Personal Executive Search",
        "name": "VIP Personal Executive Search GmbH & Co. KG",
        "address": ["Sachsenkorso 65", "15834 Rangsdorf"],
        "contact": ["www.vip-personal.org", "glowatzki@vip-personal.org", "Mobil: 0152 31954678"],
        "cover_logo": "Logos/Firmen Logo VIP .png",
        "header_logo": "Logos/Firmen Logo VIP .png",
        "logo_mode": "icon_text",            # kleines Icon + Firmenschriftzug
        "wordmark": "VIP Personal Executive Search",
        "gradient_color": VIP_BRAND_COLOR,   # Grün, passend zum VIP-Logo
    },
}

DEFAULT_BRAND = "vertico"
