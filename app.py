"""
Kandidatenprofil-Generator -- Streamlit-App

Start lokal mit:
    streamlit run app.py

Nutzt für die KI-Unterstützung die Anthropic API (Claude). Ein API-Key ist
erforderlich (siehe README, Abschnitt "API-Key dauerhaft speichern" für
automatisches Vorausfüllen statt manueller Neueingabe bei jedem Start).

Ablauf:
1. Interviewprotokoll (.txt aus ECHO KI) und optional Lebenslauf (.docx/.pdf)
   per Drag & Drop hochladen. Ein Kandidatenfoto wird dabei automatisch aus
   dem Lebenslauf erkannt.
2. Gewünschten Seitenumfang einstellen.
3. "Entwurf erstellen" -> Claude extrahiert Personendaten und entwirft die
   Fließtexte je Abschnitt.
4. Alle Felder und Texte sind editierbar, bevor das finale .docx erzeugt wird.
5. "Kandidatenprofil erzeugen" -> fertiges .docx im Vertico-Design, lokal
   speicherbar.
"""

import os
import io
import datetime
import tempfile
import streamlit as st
from PIL import Image

import config
import ai_extraction
import docx_builder
import image_extraction
from file_reading import extract_text

st.set_page_config(page_title="Kandidatenprofil-Generator", layout="wide")

_APP_DIR = os.path.dirname(os.path.abspath(__file__))

# -----------------------------------------------------------------------
# DESIGN: Variante C "Dossier" -- editorial, markenstark. Tiefes Navy mit
# Messing-Akzent auf Elfenbein, Serifen-Display (Cormorant Garamond) mit
# Karla als Textschrift, Abschnitte als ruhige Dossier-Karten. Rein
# optisch -- keine funktionale Änderung an der Logik der App.
# -----------------------------------------------------------------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@500;600;700&family=Karla:wght@400;500;600;700&display=swap');

    :root{
        --navy:#16233A; --navy-2:#22324d; --brass:#B08D4C; --ink:#1C2A3A;
        --mut:#6B6152; --ln:#E6DED0; --ivory:#F6F3EC; --card:#FFFDF9;
        --sans:"Karla", -apple-system, "Helvetica Neue", Arial, sans-serif;
        --serif:"Cormorant Garamond", Georgia, "Times New Roman", serif;
    }

    /* --- Schrift überall: Karla als Grundschrift (per Vererbung + Bedienelemente) --- */
    html, body, .stApp, [data-testid="stAppViewContainer"] { font-family:var(--sans); }
    button, input, textarea, select,
    [data-baseweb], [data-baseweb] * ,
    [data-testid="stWidgetLabel"], [data-testid="stMarkdownContainer"] { font-family:var(--sans); }
    .stApp { background:var(--ivory); color:var(--ink); }

    /* --- Überschriften: ausschließlich Serifen-Display (Cormorant) --- */
    h1, h2, h3, h4,
    [data-testid="stHeading"] h1, [data-testid="stHeading"] h2, [data-testid="stHeading"] h3 {
        font-family:var(--serif) !important; color:var(--navy);
        font-weight:600; letter-spacing:.01em;
    }
    /* Abschnitts-/Kartenüberschriften mit feiner Messing-Leiste davor */
    [data-testid="stHeading"] h2, [data-testid="stHeading"] h3 {
        display:flex; align-items:center; gap:.6rem;
    }
    [data-testid="stHeading"] h2::before, [data-testid="stHeading"] h3::before {
        content:""; flex:none; width:24px; height:1.5px; background:var(--brass);
    }

    /* Streamlit-Chrome ausblenden */
    #MainMenu {visibility:hidden;}
    footer {visibility:hidden;}
    header[data-testid="stHeader"] {background:transparent;}

    /* --- Marken-Kopf --- */
    .dossier-head {
        background:var(--navy); color:#F4EFE6; border-radius:8px;
        padding:24px 30px; margin:0 0 10px; position:relative; overflow:hidden;
    }
    .dossier-head::after {
        content:""; position:absolute; left:0; bottom:0; height:2px; width:56%;
        background:linear-gradient(90deg, var(--brass), transparent);
    }
    .dossier-head .eyebrow {
        font-family:var(--sans); font-size:.72rem; letter-spacing:.28em;
        text-transform:uppercase; color:var(--brass);
    }
    .dossier-head h1 {
        font-family:var(--serif) !important; color:#F7F2E9 !important;
        font-size:2.4rem; margin:.25rem 0 .15rem; letter-spacing:.01em;
    }
    .dossier-head .sub { color:#B9C2D0; font-size:.95rem; }

    /* --- Buttons --- */
    .stButton > button {
        border-radius:4px; border:1px solid var(--ln); background:transparent;
        color:var(--navy); font-family:var(--sans); font-weight:600;
        padding:.5rem 1.2rem; transition:all .15s ease;
    }
    .stButton > button:hover { border-color:var(--brass); color:var(--navy); }
    .stButton > button[kind="primary"] {
        background:var(--navy); border:1px solid var(--navy); color:#F4EFE6;
        text-transform:uppercase; letter-spacing:.14em; font-size:.8rem; padding:.7rem 1.4rem;
    }
    .stButton > button[kind="primary"]:hover {
        background:var(--navy-2); border-color:var(--navy-2); color:#ffffff;
    }

    /* --- Eingabefelder: reduziert, nur Unterstrich statt Kasten --- */
    .stTextInput > div > div, .stTextArea > div > div {
        background:transparent !important; border:none !important;
        border-bottom:1.5px solid var(--ln) !important; border-radius:0 !important;
        box-shadow:none !important;
    }
    .stTextInput > div > div:focus-within, .stTextArea > div > div:focus-within {
        border-bottom-color:var(--brass) !important;
    }
    .stTextInput input, .stTextArea textarea {
        background:transparent !important; color:var(--ink) !important; font-family:var(--sans) !important;
    }

    /* --- Feldbeschriftungen: kleine, gesperrte Großbuchstaben (dezent) --- */
    [data-testid="stWidgetLabel"] p,
    [data-testid="stWidgetLabel"] label {
        font-size:.72rem !important; font-weight:700 !important;
        letter-spacing:.11em !important; text-transform:uppercase !important;
        color:var(--mut) !important; line-height:1.4 !important;
    }

    /* --- Upload-Felder dezent (bleiben Ablagefelder) --- */
    [data-testid="stFileUploaderDropzone"] {
        background:var(--ivory); border:1px dashed var(--ln); border-radius:6px;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] { font-size:.9rem !important; }

    /* --- Dossier-Karten: flach, dünne Linie, kein Schatten --- */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background:var(--card); border:1px solid var(--ln) !important;
        border-radius:6px; box-shadow:none;
    }
    [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stHeading"] h3 {
        font-size:1.5rem; margin-top:.1rem;
    }

    /* --- Rubriken: feine gepunktete Trennlinien --- */
    [data-testid="stCheckbox"] {
        border-bottom:1px dotted var(--ln); padding-bottom:.45rem;
    }

    /* Hinweistexte / Divider zurückgenommen */
    [data-testid="stCaptionContainer"] { color:var(--mut) !important; }
    hr { border-color:var(--ln); }

    /* Status-Pille (KI-Zugang) */
    .pill {
        display:inline-block; font-family:var(--sans); font-size:.9rem; font-weight:600;
        padding:5px 12px; border-radius:20px; border:1px solid var(--ln);
        background:var(--ivory); color:var(--ink);
    }
    .pill.ok { color:var(--navy); border-color:var(--brass); }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------
# Server-Secrets (nur in der gehosteten Version gesetzt):
#   APP_PASSWORD       -> aktiviert den Passwortschutz (Zugang nur mit Passwort)
#   ANTHROPIC_API_KEY  -> hinterlegt den KI-Zugang zentral, sodass Nutzer selbst
#                         keinen Key eingeben müssen und keinen zu sehen bekommen
# Lokal (ohne diese Secrets) verhält sich die App wie bisher.
# ---------------------------------------------------------------------
def _secret(name):
    try:
        return st.secrets.get(name, "")
    except Exception:
        return ""


HOSTED = bool(_secret("APP_PASSWORD"))


def _require_password():
    """Einfacher Passwortschutz für die gehostete Version. Ist kein
    APP_PASSWORD-Secret gesetzt (lokale Nutzung), bleibt der Schutz inaktiv."""
    expected = _secret("APP_PASSWORD")
    if not expected:
        return True
    if st.session_state.get("_auth_ok"):
        return True
    st.markdown("### Kandidatenprofil-Generator")
    st.caption("Bitte das Passwort eingeben, um fortzufahren.")
    pw = st.text_input("Passwort", type="password", key="_pw_input")
    if st.button("Anmelden", type="primary"):
        if pw == expected:
            st.session_state["_auth_ok"] = True
            st.rerun()
        else:
            st.error("Falsches Passwort.")
    return False


if not _require_password():
    st.stop()

# Marken-Kopf im Dossier-Stil (Navy-Band, Serifen-Titel, Messing-Leiste).
st.markdown(
    '<div class="dossier-head">'
    '<div class="eyebrow">Vertico Executive Search</div>'
    '<h1>Kandidatenprofil-Generator</h1>'
    '<div class="sub">Interview &amp; Lebenslauf → fertiges Kandidatenprofil im Vertico-Design</div>'
    '</div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------
# Session State Initialisierung
# ---------------------------------------------------------------------
if "personal_data" not in st.session_state:
    st.session_state.personal_data = {k: "" for k, _ in config.PERSONAL_DATA_FIELDS}
    st.session_state.personal_data.update({"interviewpartner": "", "rolle": "", "rolle_kurz": ""})
if "sections" not in st.session_state:
    st.session_state.sections = {e["key"]: "" for e in config.PROFILE_SECTIONS}
if "active_sections" not in st.session_state:
    # Welche Rubriken tatsächlich erzeugt/angezeigt/exportiert werden.
    # Wird beim Klick auf "Entwurf erstellen" aus den Checkboxen gesetzt.
    st.session_state.active_sections = [e["key"] for e in config.PROFILE_SECTIONS]
if "is_blind" not in st.session_state:
    # Ob der zuletzt erstellte Entwurf ein Blindprofil ist (bestimmt die
    # angezeigten Personendaten-Felder in Schritt 3 und den Export).
    st.session_state.is_blind = False
if "is_leadership" not in st.session_state:
    # Ob der zuletzt erstellte Entwurf eine Führungsposition ist (bestimmt die
    # vierte Rubrik: Führungsverständnis vs. Fachliche Kompetenzen).
    st.session_state.is_leadership = True
if "photo_path" not in st.session_state:
    st.session_state.photo_path = None
if "draft_version" not in st.session_state:
    # Wird bei jeder programmatischen Neubefüllung (KI-Entwurf oder
    # "Antwort übernehmen") hochgezählt. Die Eingabefelder unten hängen
    # ihren key von dieser Nummer ab -- das ist nötig, weil Streamlit bei
    # Feldern mit festem 'key' den 'value'-Parameter bei jedem weiteren
    # Durchlauf ignoriert (bekannte Streamlit-Falle). Ohne diesen Trick
    # bleiben neu geladene Texte in der Anzeige leer, obwohl sie im
    # Hintergrund korrekt ankommen.
    st.session_state.draft_version = 0

# ---------------------------------------------------------------------
# API-Key: Vorrang hat ein serverseitig hinterlegter Key (Secret bzw.
# Umgebungsvariable). In der gehosteten Version ist dadurch der Key zentral
# hinterlegt -- Nutzer brauchen selbst keinen und sehen kein Key-Feld.
# Nur wenn kein Key hinterlegt ist (lokale Nutzung), erscheint ein
# Eingabefeld, dessen Wert dauerhaft in .streamlit/secrets.toml gespeichert
# wird (einmal eingeben genügt).
# ---------------------------------------------------------------------
def _save_api_key(key: str) -> bool:
    """Legt den Key dauerhaft in .streamlit/secrets.toml ab (lokale Nutzung)."""
    try:
        secrets_dir = os.path.join(_APP_DIR, ".streamlit")
        os.makedirs(secrets_dir, exist_ok=True)
        safe = key.replace("\\", "\\\\").replace('"', '\\"')
        with open(os.path.join(secrets_dir, "secrets.toml"), "w", encoding="utf-8") as f:
            f.write("# Automatisch von der App gespeichert.\n")
            f.write("# Enthaelt ein Geheimnis -- Datei nicht weitergeben.\n")
            f.write(f'ANTHROPIC_API_KEY = "{safe}"\n')
        return True
    except Exception:
        return False


_BRAND_LABELS = {
    "Vertico Executive Search": "vertico",
    "VIP Personal Executive Search": "vip",
}
_hinterlegter_key = _secret("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_API_KEY", "")

# ---------------------------------------------------------------------
# Einstellungen als kompakte Kopf-Leiste (statt Seitenleiste): KI-Zugang,
# Profiltyp, Branding, Positionstyp -- gelten für Entwurf und Export.
# ---------------------------------------------------------------------
with st.container(border=True):
    _cset = st.columns([1.2, 1, 1.25, 1.4])
    with _cset[0]:
        if _hinterlegter_key:
            # Gehostet bzw. lokal bereits hinterlegt: kein Key-Feld, nur Status.
            api_key = _hinterlegter_key
            st.markdown("**KI-Zugang**")
            st.markdown('<span class="pill ok">🔐 hinterlegt</span>', unsafe_allow_html=True)
        else:
            api_key = st.text_input(
                "Anthropic API-Key", type="password", value="",
                help="Wird nur für Aufrufe an die Anthropic API verwendet und nach "
                     "der ersten Eingabe dauerhaft gespeichert.")
            if api_key and st.session_state.get("_persisted_key") != api_key:
                if _save_api_key(api_key):
                    st.success("API-Key gespeichert.")
                st.session_state["_persisted_key"] = api_key
            if not api_key:
                st.warning("Kein API-Key hinterlegt.")
    with _cset[1]:
        profile_type = st.radio(
            "Profiltyp", ["Vollprofil", "Blindprofil"], key="profile_type",
            help="Blindprofil: vollständig anonym – kein Name, kein Foto, keine "
                 "identifizierenden Angaben; der Kandidat wird trotzdem "
                 "vollumfänglich beschrieben.")
        is_blind_selected = (profile_type == "Blindprofil")
    with _cset[2]:
        brand_choice = st.radio(
            "Branding", list(_BRAND_LABELS.keys()), key="brand_choice",
            help="Bestimmt Logo, Firmierung, Adresse und Deckblatt-Farbe des Berichts.")
        brand_key = _BRAND_LABELS[brand_choice]
    with _cset[3]:
        positionstyp = st.radio(
            "Positionstyp", ["Führungsposition", "Fachposition (ohne Führung)"],
            key="positionstyp",
            help="Führungsposition → Rubrik „Führungsverständnis“; Fachposition → "
                 "„Fachliche Kompetenzen und Arbeitsweise“.")
        is_leadership_selected = (positionstyp == "Führungsposition")
    blind_reference = ""
    if is_blind_selected:
        blind_reference = st.text_input(
            "Chiffre / Referenz (statt Name)", key="blind_reference",
            placeholder="z. B. Ref. VIP-2026-001",
            help="Erscheint auf dem Deckblatt anstelle des Namens. Kann leer bleiben.")

# ---------------------------------------------------------------------
# Unterlagen (Interviewprotokoll, Lebenslauf, Foto) -- als Karte, zwei Spalten
# ---------------------------------------------------------------------
if is_blind_selected:
    st.info("🔒 **Blindprofil aktiv:** Name, Foto und identifizierende Angaben "
            "(Firma, genauer Wohnort, Kontakt) werden im Bericht weggelassen – "
            "ein hochgeladenes Foto wird ignoriert.")

tmp_dir = "tmp_uploads"
os.makedirs(tmp_dir, exist_ok=True)

with st.container(border=True):
    st.subheader("Unterlagen")
    u1, u2 = st.columns(2)
    with u1:
        transcript_file = st.file_uploader(
            "Interviewprotokoll (.txt / .docx / .pdf)", type=["txt", "docx", "pdf"])
        transcript_paste = st.text_area(
            "…oder Text direkt einfügen (falls keine Datei vorliegt)",
            height=120, key="transcript_paste",
            help="Wird verwendet, wenn oben keine Datei hochgeladen wurde. Bei "
                 "hochgeladener Datei hat die Datei Vorrang.")
    with u2:
        cv_file = st.file_uploader("Lebenslauf (.docx / .pdf)", type=["docx", "pdf"])
        # Foto automatisch aus dem Lebenslauf erkennen -- immer als Vorschau, nie
        # ungesehen ins Dokument übernommen.
        extracted_photo_path = None
        if cv_file is not None:
            photo_bytes = image_extraction.extract_photo_from_cv(cv_file.getvalue(), cv_file.name)
            if photo_bytes:
                candidate_path = os.path.join(tmp_dir, "foto_aus_lebenslauf.png")
                try:
                    with Image.open(io.BytesIO(photo_bytes)) as img:
                        img.convert("RGB").save(candidate_path)
                    extracted_photo_path = candidate_path
                except Exception:
                    extracted_photo_path = None
        if extracted_photo_path:
            st.image(extracted_photo_path, width=104)
            st.caption("Foto automatisch aus dem Lebenslauf erkannt. Falls es nicht "
                       "passt, hier ein eigenes hochladen:")
            manual_photo_file = st.file_uploader(
                "Anderes Foto verwenden (optional)", type=["jpg", "jpeg", "png"],
                label_visibility="collapsed")
        else:
            if cv_file is not None:
                st.caption("Im Lebenslauf wurde kein passendes Foto gefunden – bitte manuell hochladen.")
            manual_photo_file = st.file_uploader(
                "Kandidatenfoto für Deckblatt (optional, .jpg/.png)", type=["jpg", "jpeg", "png"])

# Auswertung nach den Spalten (Variablen aus den with-Blöcken bleiben gültig).
transcript_pasted_text = (transcript_paste or "").strip()
has_transcript = bool(transcript_file is not None or transcript_pasted_text)
if transcript_file is not None and transcript_pasted_text:
    st.caption("ℹ️ Datei UND Text vorhanden – verwendet wird die hochgeladene Datei.")

if manual_photo_file is not None:
    photo_path = os.path.join(tmp_dir, manual_photo_file.name)
    with open(photo_path, "wb") as f:
        f.write(manual_photo_file.getbuffer())
    st.session_state.photo_path = photo_path
else:
    st.session_state.photo_path = extracted_photo_path

# ---------------------------------------------------------------------
# Zielposition (Pflichtfeld-Titel + optionale Beschreibung) -- Karte, zwei Spalten
# ---------------------------------------------------------------------
with st.container(border=True):
    st.subheader("Zielposition")
    z1, z2 = st.columns(2)
    with z1:
        position_titel = st.text_input(
            "Zu besetzende Position (genaue Bezeichnung) *",
            key="position_titel",
            placeholder="z. B. Vertriebsleitung Neuanlage · Montageleitung · Servicetechniker",
            help="Pflichtfeld. Konkrete Stellenbezeichnung. Das Programm erkennt daraus "
                 "die Art der Position und richtet das Profil inhaltlich darauf aus "
                 "(ein Vertriebsprofil wird anders formuliert als z. B. eine "
                 "Montageleitung). Erscheint zudem als Untertitel auf dem Deckblatt.")
    with z2:
        ziel_beschreibung = st.text_area(
            "Beschreibung der Zielposition (optional: Firma, Aufgaben, Anforderungen)",
            key="ziel_beschreibung", height=130,
            placeholder="Firma und Branche, konkrete Aufgaben und Verantwortung, "
                        "Anforderungen …",
            help="Optional, aber hilfreich: schärft die inhaltliche Ausrichtung. "
                 "Über den Kandidaten werden keine Fakten erfunden.")

# ---------------------------------------------------------------------
# Feinjustierung und Rubriken nebeneinander -- zwei Karten
# ---------------------------------------------------------------------
_fj, _rb = st.columns(2)
with _fj:
    with st.container(border=True):
        st.subheader("Feinjustierung")
        rating = st.slider(
            "Bewertung des Kandidaten", min_value=1, max_value=10, value=5, step=1,
            help="1 = schwach, 10 = perfekt; solider Durchschnitt 5–6. Steuert Tonfall "
                 "und Gewichtung, nie die Fakten.")
        target_pages = st.slider(
            "Umfang (Seiten, Richtwert)", min_value=2, max_value=6, value=4, step=1)
        st.caption(
            f"≈ {config.WORDS_PER_PAGE_ESTIMATE} Wörter/Seite → ca. "
            f"{target_pages * config.WORDS_PER_PAGE_ESTIMATE} Wörter Fließtext. "
            f"Der Seitenumbruch in Word kann abweichen.")
with _rb:
    with st.container(border=True):
        st.subheader("Rubriken")
        st.caption("Häkchen entfernen, um eine Rubrik nicht aufzunehmen.")
        selected_keys = []
        for entry in config.get_profile_sections(is_leadership_selected):
            heading = entry["heading"].format(rolle=(position_titel or "Zielrolle"))
            if st.checkbox(heading, value=True, key=f"secsel_{entry['key']}"):
                selected_keys.append(entry["key"])

can_draft = bool(has_transcript and api_key and selected_keys and position_titel.strip())
_bl, _br = st.columns([3, 1])
with _br:
    _draft_clicked = st.button("Entwurf erstellen", type="primary",
                               disabled=not can_draft, use_container_width=True)
if _draft_clicked:
    try:
        with st.spinner("Lese Dokumente ein..."):
            if transcript_file is not None:
                transcript_text = extract_text(transcript_file)
            else:
                transcript_text = transcript_pasted_text
            cv_text = extract_text(cv_file) if cv_file is not None else ""

        with st.spinner("Extrahiere Personendaten..."):
            personal_data = ai_extraction.extract_personal_data(
                api_key, transcript_text, cv_text, blind=is_blind_selected)
            st.session_state.personal_data.update(personal_data)

        # Deckblatt-Titel festlegen: die explizite Positionsbezeichnung hat Vorrang,
        # sonst die Kategorie-Auswahl. Bleibt in Schritt 3 editierbar.
        cover_title = position_titel.strip()
        if cover_title:
            st.session_state.personal_data["rolle"] = cover_title
            st.session_state.personal_data["rolle_kurz"] = cover_title

        with st.spinner("Entwerfe Profiltexte (das kann einen Moment dauern)..."):
            sections = ai_extraction.draft_sections(
                api_key, transcript_text, cv_text, target_pages,
                rating=rating, zielposition=position_titel, section_keys=selected_keys,
                blind=is_blind_selected, zielbeschreibung=ziel_beschreibung)
            # Nur die gewählten Abschnitte des aktuellen Positionstyps behalten
            # (verhindert, dass z.B. eine Führungs-Rubrik aus einem früheren Entwurf
            # zurückbleibt, wenn nun eine Fachposition erstellt wird).
            st.session_state.sections = {k: "" for k in selected_keys}
            st.session_state.sections.update(sections)

        st.session_state.active_sections = list(selected_keys)
        st.session_state.is_blind = is_blind_selected
        st.session_state.is_leadership = is_leadership_selected
        st.session_state.draft_version += 1
        st.success("Entwurf erstellt. Bitte unten prüfen und bei Bedarf bearbeiten.")
    except Exception as e:
        msg = str(e)
        low = msg.lower()
        if "529" in msg or "overloaded" in low:
            st.warning(
                "⏳ Die KI-Server von Anthropic sind gerade überlastet (Fehler 529). "
                "Das ist vorübergehend und liegt **nicht** an deinem Zugang oder Guthaben. "
                "Die App hat es bereits mehrfach automatisch wiederholt – bitte kurz "
                "warten und dann erneut auf „Entwurf erstellen“ klicken."
            )
        elif "429" in msg or "rate limit" in low:
            st.warning(
                "⏳ Zu viele Anfragen in kurzer Zeit (Rate-Limit). Bitte einen Moment "
                "warten und erneut versuchen."
            )
        elif "401" in msg or "authentication" in low or "invalid x-api-key" in low:
            st.error(
                "Der hinterlegte API-Key ist ungültig. Bitte den Anthropic-API-Key prüfen."
            )
        elif "402" in msg or "credit" in low or "billing" in low or "insufficient" in low:
            st.error(
                "Kein Guthaben auf dem Anthropic-Konto. Bitte in der Anthropic Console "
                "unter Billing Guthaben aufladen."
            )
        elif "unterminated" in low or "expecting" in low or "jsondecode" in low:
            st.warning(
                "⏳ Die KI-Antwort kam unvollständig an (meist wegen kurzzeitiger "
                "Serverüberlastung). Die App hat es bereits mehrfach automatisch "
                "versucht – bitte kurz warten und erneut auf „Entwurf erstellen“ klicken."
            )
        else:
            st.error(
                f"Der Entwurf konnte nicht erstellt werden: {e}\n\n"
                f"Mögliche Ursachen: vorübergehende Störung, Netzwerkproblem, "
                f"ungültiger API-Key oder fehlendes Guthaben. Bitte erneut versuchen."
            )

if not api_key:
    st.info("Bitte oben den Anthropic API-Key eingeben (oder dauerhaft speichern, siehe README).")
elif not has_transcript:
    st.info("Bitte zunächst ein Interviewprotokoll hochladen oder einfügen.")
elif not position_titel.strip():
    st.info("Bitte die zu besetzende Position eintragen (Pflichtfeld).")
elif not selected_keys:
    st.info("Bitte mindestens eine Rubrik auswählen.")

# ---------------------------------------------------------------------
# Prüfen & bearbeiten: Personendaten (Karte)
# ---------------------------------------------------------------------
v = st.session_state.draft_version
with st.container(border=True):
    st.subheader("Personendaten prüfen")
    pd_cols = st.columns(2)
    # Im Blindprofil die anonymisierten Felder anzeigen (Jahrgang, Region, …).
    field_items = (config.BLIND_PERSONAL_DATA_FIELDS
                   if st.session_state.get("is_blind") else config.PERSONAL_DATA_FIELDS)
    for i, (key, label) in enumerate(field_items):
        col = pd_cols[i % 2]
        st.session_state.personal_data[key] = col.text_input(
            label, value=st.session_state.personal_data.get(key, ""), key=f"pd_{key}_{v}")

    extra_cols = st.columns(3)
    st.session_state.personal_data["interviewpartner"] = extra_cols[0].text_input(
        "Interviewpartner", value=st.session_state.personal_data.get("interviewpartner", ""),
        key=f"pd_interviewpartner_{v}")
    st.session_state.personal_data["rolle"] = extra_cols[1].text_input(
        "Zielposition (voll)", value=st.session_state.personal_data.get("rolle", ""),
        key=f"pd_rolle_{v}")
    st.session_state.personal_data["rolle_kurz"] = extra_cols[2].text_input(
        "Zielposition (kurz, für Deckblatt)", value=st.session_state.personal_data.get("rolle_kurz", ""),
        key=f"pd_rolle_kurz_{v}")

# ---------------------------------------------------------------------
# Prüfen & bearbeiten: Profiltexte (Karte)
# ---------------------------------------------------------------------
with st.container(border=True):
    st.subheader("Profiltexte prüfen")
    _active = st.session_state.get("active_sections") or [e["key"] for e in config.PROFILE_SECTIONS]
    # Über die kanonische Gesamtreihenfolge iterieren und nach aktiven Rubriken filtern,
    # damit unabhängig vom Positionstyp die richtige Rubrik (Führung vs. Fachkompetenz)
    # in der richtigen Reihenfolge erscheint.
    for entry in config.ALL_SECTIONS_ORDERED:
        if entry["key"] not in _active:
            continue
        heading = entry["heading"].format(rolle=st.session_state.personal_data.get("rolle", ""))
        text_vorhanden = bool(st.session_state.sections.get(entry["key"], "").strip())
        label_suffix = "" if text_vorhanden else "  ⚠️ noch leer"
        with st.expander(heading + label_suffix, expanded=not text_vorhanden):
            st.session_state.sections[entry["key"]] = st.text_area(
                "Text", value=st.session_state.sections.get(entry["key"], ""),
                height=200, key=f"section_{entry['key']}_{v}", label_visibility="collapsed")

# ---------------------------------------------------------------------
# Kandidatenprofil erzeugen (Karte)
# ---------------------------------------------------------------------
with st.container(border=True):
    st.subheader("Kandidatenprofil erzeugen")
    if HOSTED:
        # Gehostet: kein Serverpfad-Feld -- die Datei landet in einem temporären
        # Ordner und wird ausschließlich über den Download-Button ausgeliefert.
        output_dir = os.path.join(tempfile.gettempdir(), "kpg_output")
    else:
        output_dir = st.text_input("Zielordner (lokal)", value=os.path.join(os.getcwd(), "output"))

    if st.button("Kandidatenprofil erzeugen", type="primary"):
        os.makedirs(output_dir, exist_ok=True)
        if st.session_state.get("is_blind"):
            ref = (blind_reference or "").strip() or "Blindprofil"
            filename = f"Kandidatenprofil_{ref}.docx".replace(" ", "_")
        else:
            vorname = st.session_state.personal_data.get("vorname", "").strip() or "Vorname"
            nachname = st.session_state.personal_data.get("nachname", "").strip() or "Nachname"
            filename = f"Kandidatenprofil_{vorname}_{nachname}.docx".replace(" ", "_")
        output_path = os.path.join(output_dir, filename)

        docx_builder.build_candidate_profile(
            data=st.session_state.personal_data,
            sections=st.session_state.sections,
            output_path=output_path,
            photo_path=st.session_state.photo_path,
            section_keys=st.session_state.get("active_sections"),
            brand=brand_key,
            blind=st.session_state.get("is_blind", False),
            reference=blind_reference,
        )
        if HOSTED:
            st.success("Kandidatenprofil erstellt – bitte jetzt herunterladen:")
        else:
            st.success(f"Kandidatenprofil gespeichert unter: {output_path}")
        with open(output_path, "rb") as f:
            st.download_button("Herunterladen", data=f.read(), file_name=filename)
