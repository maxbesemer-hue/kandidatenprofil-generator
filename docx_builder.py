"""
Erzeugt ein Kandidatenprofil-Word-Dokument im Vertico-Design.

Alle Design-Parameter kommen aus config.py und wurden dort mit Quellenangabe
aus der Originalvorlage referenziert. Dieses Modul enthält keine eigenen
Design-Entscheidungen "aus dem Bauch heraus".
"""

import os
import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml

import config


def _set_run_font(run, size=None, bold=None, color_hex=None, italic=None):
    run.font.name = config.FONT_NAME
    # Ostasiatische Schriftart ebenfalls setzen, sonst greift bei manchen
    # Word-Versionen ein Fallback-Font für Sonderzeichen
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), config.FONT_NAME)
    if size is not None:
        run.font.size = size
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color_hex is not None:
        run.font.color.rgb = RGBColor.from_string(color_hex)


def _add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    # "keep_with_next" sorgt dafür, dass Word/LibreOffice zwischen dieser
    # Überschrift und dem direkt folgenden Absatz NIE umbricht -- eine
    # Überschrift kann dadurch nicht mehr allein am Seitenende stehen,
    # während der zugehörige Text erst auf der Folgeseite beginnt.
    p.paragraph_format.keep_with_next = True
    if level == 1:
        p.paragraph_format.space_before = config.H1_SPACE_BEFORE
        p.paragraph_format.space_after = config.H1_SPACE_AFTER
        run = p.add_run(text)
        _set_run_font(run, size=config.H1_SIZE, bold=True, color_hex=config.H1_COLOR)
    else:
        p.paragraph_format.space_before = config.H2_SPACE_BEFORE
        p.paragraph_format.space_after = config.H2_SPACE_AFTER
        run = p.add_run(text)
        _set_run_font(run, size=config.H2_SIZE, bold=True, color_hex=config.H2_COLOR)
    return p


def _add_body_paragraphs(doc, text):
    """Fügt Fließtext ein. Leerzeilen im Eingabetext trennen Absätze,
    exakt wie in der Vorlage (mehrere kurze Absätze statt einem Block).
    Ausrichtung: Blocksatz, wie in der gesamten Vorlage durchgängig
    verwendet (word/document.xml: <w:jc w:val="both"/>, 72 Fundstellen).

    space_before wird bei JEDEM Absatz explizit auf 0 gesetzt: Dadurch wird
    der Abstand zwischen einer Überschrift und dem folgenden Text
    ausschließlich durch H1_SPACE_AFTER/H2_SPACE_AFTER bestimmt -- und ist
    dadurch garantiert bei jeder Rubrik exakt gleich groß, statt vom
    (nicht überall gleich gesetzten) Standardabstand des Normal-Stils
    abzuhängen."""
    # Zeilenweise verarbeiten: Zeilen, die mit einem Aufzählungszeichen beginnen,
    # werden als echte Word-Liste (Stil "List Bullet") gesetzt, der Rest als
    # Blocksatz-Absätze. Leerzeilen trennen Absätze.
    #
    # Abstände: Eine Aufzählung soll sich als Einheit vom Fließtext absetzen --
    # deutlich Luft davor (BULLET_SPACE_BEFORE) und danach (BULLET_SPACE_AFTER),
    # die Punkte untereinander aber eng (BULLET_SPACE_BETWEEN).
    #
    # Seitenumbrüche: Der einleitende Satz bleibt beim ersten Aufzählungspunkt
    # (keep_with_next), und die Punkte einer Liste bleiben untereinander zusammen
    # (keep_with_next auf allen außer dem letzten). Dadurch wird eine Liste nie
    # mitten auseinandergerissen und ihr Einleitungssatz steht nie verwaist am
    # Seitenende.
    bullet_prefixes = ("- ", "• ", "* ", "– ", "‐ ", "· ", "•\t", "-\t")
    prose = []
    group = []          # aktuelle, noch offene Aufzählungsgruppe
    prev_was_bullet = False

    def _close_group():
        # Innerhalb der Gruppe zusammenhalten, aber die Liste vom Folgeabsatz
        # lösbar lassen (letzter Punkt ohne keep_with_next).
        if not group:
            return
        for bp in group[:-1]:
            bp.paragraph_format.keep_with_next = True
        group[-1].paragraph_format.keep_with_next = False
        group.clear()

    def _flush_prose(keep_with_next=False, after_bullets=False):
        nonlocal prev_was_bullet
        block = "\n".join(prose).strip()
        prose.clear()
        if not block:
            return
        p = doc.add_paragraph()
        p.paragraph_format.space_before = (
            config.BULLET_SPACE_AFTER if after_bullets else Pt(0))
        if keep_with_next:
            p.paragraph_format.keep_with_next = True
        if config.BODY_JUSTIFY:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_run_font(p.add_run(block), size=config.BODY_SIZE)
        prev_was_bullet = False

    for raw in text.split("\n"):
        s = raw.strip()
        if not s:
            _close_group()
            _flush_prose(after_bullets=prev_was_bullet)
            continue
        marker = next((pref for pref in bullet_prefixes if s.startswith(pref)), None)
        if marker:
            # Einleitungssatz direkt über der Liste an den ersten Punkt binden.
            _flush_prose(keep_with_next=bool(prose and not prev_was_bullet))
            item = s[len(marker):].strip()
            bp = doc.add_paragraph(style="List Bullet")
            bp.paragraph_format.space_before = (
                config.BULLET_SPACE_BETWEEN if prev_was_bullet
                else config.BULLET_SPACE_BEFORE)
            bp.paragraph_format.space_after = Pt(0)
            _set_run_font(bp.add_run(item), size=config.BODY_SIZE)
            group.append(bp)
            prev_was_bullet = True
        else:
            # Erste Fließtextzeile nach einer Liste beendet die Gruppe.
            if prev_was_bullet:
                _close_group()
            prose.append(s)
    _close_group()
    _flush_prose(after_bullets=prev_was_bullet)


def _set_page_geometry(section):
    section.page_width = config.PAGE_WIDTH
    section.page_height = config.PAGE_HEIGHT
    section.top_margin = config.MARGIN_TOP
    section.bottom_margin = config.MARGIN_BOTTOM
    section.left_margin = config.MARGIN_LEFT
    section.right_margin = config.MARGIN_RIGHT


def _make_floating(run, offset_x, offset_y, behind_doc=False, relative_height=1):
    """Wandelt ein per run.add_picture() *inline* eingefügtes Bild in ein
    frei schwebendes ("wp:anchor") Bild um, mit fester Position relativ zur
    Textspalte (horizontal) und zum Absatz (vertikal). Wird verwendet, um
    Bildpositionen exakt wie in der Vorlage nachzubilden (siehe config.py
    für die jeweiligen Quellenwerte je Bild)."""
    drawing = run._element.find(qn('w:drawing'))
    inline = drawing.find(qn('wp:inline'))

    extent_el = inline.find(qn('wp:extent'))
    effect_extent_el = inline.find(qn('wp:effectExtent'))
    docpr_el = inline.find(qn('wp:docPr'))
    cnv_el = inline.find(qn('wp:cNvGraphicFramePr'))
    graphic_el = inline.find(qn('a:graphic'))

    anchor = OxmlElement('wp:anchor')
    for attr, value in {
        'distT': '0', 'distB': '0', 'distL': '114300', 'distR': '114300',
        'simplePos': '0', 'relativeHeight': str(relative_height),
        'behindDoc': '1' if behind_doc else '0',
        'locked': '0', 'layoutInCell': '1', 'allowOverlap': '1',
    }.items():
        anchor.set(attr, value)

    simple_pos = OxmlElement('wp:simplePos')
    simple_pos.set('x', '0')
    simple_pos.set('y', '0')
    anchor.append(simple_pos)

    pos_h = OxmlElement('wp:positionH')
    pos_h.set('relativeFrom', 'column')
    pos_h_offset = OxmlElement('wp:posOffset')
    pos_h_offset.text = str(offset_x)
    pos_h.append(pos_h_offset)
    anchor.append(pos_h)

    pos_v = OxmlElement('wp:positionV')
    pos_v.set('relativeFrom', 'paragraph')
    pos_v_offset = OxmlElement('wp:posOffset')
    pos_v_offset.text = str(offset_y)
    pos_v.append(pos_v_offset)
    anchor.append(pos_v)

    if extent_el is not None:
        anchor.append(extent_el)
    if effect_extent_el is not None:
        anchor.append(effect_extent_el)
    anchor.append(OxmlElement('wp:wrapNone'))
    if docpr_el is not None:
        anchor.append(docpr_el)
    if cnv_el is not None:
        anchor.append(cnv_el)
    if graphic_el is not None:
        anchor.append(graphic_el)

    drawing.remove(inline)
    drawing.append(anchor)


def _add_header_logo(section, assets_dir, brand):
    """Setzt das Logo oben rechts, freischwebend hinter dem Text --
    exakt wie in header1.xml der Vorlage (Position/Größe siehe config.py).
    Das Logo richtet sich nach der gewählten Marke (brand["header_logo"]).

    WICHTIG: python-docx's run.add_picture() fügt Bilder nur *inline* ein
    (in den normalen Textfluss, begrenzt durch den Satzspiegel/Rand). Die
    Vorlage positioniert das Logo dagegen frei schwebend (floating/
    "anchored"), wodurch es bewusst über den rechten Satzspiegelrand hinaus
    bis nah an die Papierkante reicht. Eine reine Rechtsbündig-Ausrichtung
    (wie zuvor) bleibt dadurch sichtbar weiter innen als im Original."""
    header = section.header
    header.is_linked_to_previous = False
    p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    run = p.add_run()
    logo_path = os.path.join(assets_dir, brand["header_logo"])
    run.add_picture(logo_path, width=config.HEADER_LOGO_SIZE, height=config.HEADER_LOGO_SIZE)
    _make_floating(run, config.HEADER_LOGO_OFFSET_X, config.HEADER_LOGO_OFFSET_Y,
                    behind_doc=True, relative_height=1)
    # Absatzausrichtung spielt für ein frei schwebendes Bild keine Rolle
    # mehr, schadet aber auch nicht als Fallback für Programme, die Anker
    # nicht unterstützen.
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def _add_footer_page_number(section):
    """Seitenzahl rechtsbündig im Footer -- wie in footer1.xml der Vorlage."""
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = 'PAGE   \\* MERGEFORMAT'
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')

    run = p.add_run()
    _set_run_font(run, size=config.BODY_SIZE)
    run._element.append(fld_begin)
    run._element.append(instr)
    run._element.append(fld_sep)
    run._element.append(fld_end)


def _fit_photo_size(photo_path):
    """Berechnet Breite/Höhe für das Deckblatt-Foto so, dass es innerhalb
    von config.COVER_PHOTO_MAX_WIDTH x config.COVER_PHOTO_MAX_HEIGHT bleibt
    und dabei das Seitenverhältnis erhält ('contain'-Skalierung). Das
    verhindert, dass ein hochformatiges Foto das Deckblatt auf eine zweite
    Seite drückt -- unabhängig davon, welches Seitenverhältnis das
    hochgeladene Foto hat."""
    from PIL import Image
    with Image.open(photo_path) as img:
        px_w, px_h = img.size
    aspect = px_w / px_h
    max_w = config.COVER_PHOTO_MAX_WIDTH
    max_h = config.COVER_PHOTO_MAX_HEIGHT
    if max_w / max_h > aspect:
        # Höhe ist der begrenzende Faktor
        height = max_h
        width = Emu(int(max_h * aspect))
    else:
        # Breite ist der begrenzende Faktor
        width = max_w
        height = Emu(int(max_w / aspect))
    return width, height


# Radialer Blauverlauf auf dem Deckblatt. Quelle: word/document.xml der
# Vorlage, eine frei schwebende Rechteck-Form (behindDoc="1") mit
# Farbverlauf-Füllung ("gradFill"), die bewusst größer als die Seite ist
# und mit negativem Versatz positioniert wird, sodass nur die obere linke
# Ecke sichtbar in den blauen Farbverlauf übergeht. Alle Positions- und
# Geometriewerte (positionH/V, extent, effectExtent, Gradient-Stopps,
# Pfadtyp) sind 1:1 aus der Vorlage übernommen. Einzige bewusste Änderung:
# Die Farben sind dort als Theme-Referenzen ("bg1"/"accent1") hinterlegt,
# die vom Theme des jeweiligen Dokuments abhängen. Da unser generiertes
# Dokument nicht das Vertico-Theme der Vorlage enthält, sind sie hier als
# feste RGB-Werte (Weiß bzw. #4F81BD, identisch zu H2_COLOR) hinterlegt,
# damit die Farbe unabhängig vom Theme exakt gleich bleibt.
_COVER_BACKGROUND_XML = """
<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
     xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
     xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing">
<w:rPr><w:sz w:val="2"/></w:rPr>
<w:drawing>
<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"
  relativeHeight="1" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="column"><wp:posOffset>-1619250</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="paragraph"><wp:posOffset>-1017270</wp:posOffset></wp:positionV>
<wp:extent cx="8458200" cy="10206990"/>
<wp:effectExtent l="57150" t="19050" r="57150" b="80010"/>
<wp:wrapNone/>
<wp:docPr id="900001" name="Deckblatt-Hintergrund"/>
<wp:cNvGraphicFramePr/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp>
<wps:cNvSpPr/>
<wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="8458200" cy="10206990"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:gradFill flip="none" rotWithShape="1">
<a:gsLst>
<a:gs pos="80000"><a:srgbClr val="FFFFFF"/></a:gs>
<a:gs pos="100000"><a:srgbClr val="4F81BD"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:srgbClr></a:gs>
</a:gsLst>
<a:path path="circle"><a:fillToRect l="100000" t="100000"/></a:path>
<a:tileRect r="-100000" b="-100000"/>
</a:gradFill>
<a:ln><a:noFill/></a:ln>
</wps:spPr>
<wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow"
  vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"
  numCol="1" spcCol="0" rtlCol="0" fromWordArt="0" anchor="ctr" anchorCtr="0"
  forceAA="0" compatLnSpc="1">
<a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:noAutofit/>
</wps:bodyPr>
</wps:wsp>
</a:graphicData></a:graphic>
<wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>0</wp14:pctWidth></wp14:sizeRelH>
<wp14:sizeRelV relativeFrom="margin"><wp14:pctHeight>0</wp14:pctHeight></wp14:sizeRelV>
</wp:anchor>
</w:drawing>
</w:r>
"""


def _add_cover_branding(doc, assets_dir, brand):
    """Fügt das Deckblatt-Branding ein: den radialen Verlauf in der
    Markenfarbe sowie das Logo -- abhängig vom Branding-Modus:

    - brand['logo_mode'] == 'image': großes, freischwebendes Logo, exakt wie
      in der Vorlage (Vertico, hochauflösendes Logo).
    - brand['logo_mode'] == 'icon_text': kleines Icon zentriert oben im
      Textfluss, darunter der Firmenschriftzug (Partnermarke mit niedrig
      aufgelöstem Icon ohne eingebauten Namen).

    Der Verlauf liegt hinter dem Text (behindDoc=1). Die Absatzhöhe des
    technischen ersten Absatzes wird mit sz=2 (1pt) minimiert."""
    p = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

    # Verlauf in der Markenfarbe (Vertico-Blau 4F81BD bzw. VIP-Grün).
    run_xml = parse_xml(_COVER_BACKGROUND_XML.replace("4F81BD", brand["gradient_color"]))
    p._p.append(run_xml)

    if brand["logo_mode"] == "image":
        # Großes, freischwebend positioniertes Logo (siehe config.COVER_LOGO_OFFSET_X/Y)
        logo_run = p.add_run()
        logo_path = os.path.join(assets_dir, brand["cover_logo"])
        logo_run.add_picture(logo_path, width=config.COVER_LOGO_WIDTH, height=config.COVER_LOGO_WIDTH)
        _make_floating(logo_run, config.COVER_LOGO_OFFSET_X, config.COVER_LOGO_OFFSET_Y,
                        behind_doc=False, relative_height=2)
    else:
        # Kleines Icon zentriert im Textfluss, darunter der Firmenschriftzug.
        icon_p = doc.add_paragraph()
        icon_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        icon_p.paragraph_format.space_before = Pt(8)
        icon_p.paragraph_format.space_after = Pt(2)
        icon_run = icon_p.add_run()
        icon_path = os.path.join(assets_dir, brand["cover_logo"])
        icon_run.add_picture(icon_path, width=config.COVER_ICON_SIZE, height=config.COVER_ICON_SIZE)
        if brand.get("wordmark"):
            wm_p = doc.add_paragraph()
            wm_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            wm_p.paragraph_format.space_after = Pt(0)
            wm_run = wm_p.add_run(brand["wordmark"])
            _set_run_font(wm_run, size=config.COVER_WORDMARK_SIZE, bold=True,
                          color_hex=brand["gradient_color"])


def build_candidate_profile(data: dict, sections: dict, output_path: str,
                             assets_dir: str = None, photo_path: str = None,
                             section_keys=None, brand: str = "vertico",
                             blind: bool = False, reference: str = ""):
    """
    Baut das vollständige Kandidatenprofil.

    data: dict mit den Personendaten-Feldern (Keys siehe config.PERSONAL_DATA_FIELDS
          bzw. config.BLIND_PERSONAL_DATA_FIELDS im Blindprofil)
          plus 'rolle' (Zielposition, für Cover-Untertitel und Eignungs-Überschrift)
          plus 'interviewpartner'
    sections: dict {section_key: Fließtext} für jeden Eintrag aus config.PROFILE_SECTIONS
    output_path: Ziel-Dateipfad für die erzeugte .docx
    assets_dir: Verzeichnis mit den Logo-Assets (Standard: Verzeichnis dieses Skripts)
    photo_path: Optionaler Pfad zu einem Kandidatenfoto für das Deckblatt
    section_keys: Optionale Liste der einzuschließenden Abschnitts-Keys. None =
                  alle Abschnitte aus config.PROFILE_SECTIONS. Abgewählte
                  Abschnitte werden mitsamt ihrer Überschrift ausgelassen.
    brand: Branding-Schlüssel ("vertico" oder "vip") -- steuert Logo, Firmierung,
           Adresse, Kontaktblock und Verlaufsfarbe (config.BRANDS).
    blind: True erzeugt ein Blindprofil (kein Name, kein Foto, anonymisierte
           Personendaten, Titelzusatz "(anonymisiert)", Chiffre statt Name).
    reference: Chiffre/Referenz, die im Blindprofil anstelle des Namens erscheint.
    """
    if assets_dir is None:
        assets_dir = os.path.dirname(os.path.abspath(__file__))

    brand = config.BRANDS.get(brand, config.BRANDS[config.DEFAULT_BRAND])

    doc = Document()

    # Basis-Stil (Normal) auf Calibri/12pt setzen, damit Fallbacks korrekt sind
    normal = doc.styles['Normal']
    normal.font.name = config.FONT_NAME
    normal.font.size = config.BODY_SIZE

    section = doc.sections[0]
    _set_page_geometry(section)

    # --- Deckblatt ---------------------------------------------------
    _add_cover_branding(doc, assets_dir, brand)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Bei freischwebendem Logo (image-Modus) reserviert dieser Abstand Platz,
    # damit der Titel unter dem Logo beginnt. Bei Icon+Schriftzug im
    # Textfluss (icon_text-Modus) ist nur ein kleiner Abstand nötig.
    title_p.paragraph_format.space_before = (
        config.COVER_TITLE_SPACE_BEFORE if brand["logo_mode"] == "image"
        else config.COVER_TITLE_SPACE_BEFORE_INLINE)
    title_p.paragraph_format.space_after = Pt(0)
    title_text = "Kandidatenprofil (anonymisiert)" if blind else "Kandidatenprofil"
    title_run = title_p.add_run(title_text)
    _set_run_font(title_run, size=config.COVER_TITLE_SIZE, bold=True)

    # Vollprofil: Kandidatenname. Blindprofil: Chiffre/Referenz (falls
    # angegeben) statt Name -- kein Rückschluss auf die Person. Darunter je
    # die Rolle, per Zeilenumbruch im selben Absatz (wie in der Vorlage).
    if blind:
        identifier = (reference or "").strip()
    else:
        identifier = f"{data.get('vorname','')} {data.get('nachname','')}".strip()
    role_text = data.get('rolle_kurz', data.get('rolle', ''))
    name_p = doc.add_paragraph()
    name_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_p.paragraph_format.space_before = Pt(12)
    if identifier:
        id_run = name_p.add_run(identifier)
        _set_run_font(id_run, size=config.COVER_NAME_SIZE, bold=True)
        name_p.add_run().add_break()
    role_run = name_p.add_run(role_text)
    _set_run_font(role_run, size=config.COVER_ROLE_SIZE)

    # Foto nur im Vollprofil -- im Blindprofil bewusst kein Bild.
    if not blind and photo_path and os.path.exists(photo_path):
        photo_p = doc.add_paragraph()
        photo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        photo_p.paragraph_format.space_before = Pt(12)
        photo_run = photo_p.add_run()
        photo_width, photo_height = _fit_photo_size(photo_path)
        photo_run.add_picture(photo_path, width=photo_width, height=photo_height)

    # Fester, kontrollierter Abstand vor dem Kontaktblock (statt mehrerer
    # Leerabsätze mit variabler, vom Zeilenabstand abhängiger Höhe -- das
    # hatte zuvor je nach Fotogröße zu einem Seitenüberlauf geführt).
    spacer_p = doc.add_paragraph()
    spacer_p.paragraph_format.space_before = config.COVER_CONTACT_SPACE_BEFORE

    for line in [brand["name"]] + list(brand["address"]):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(line)
        _set_run_font(r, size=config.COVER_CONTACT_SIZE)

    contact_p = doc.add_paragraph()
    contact_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact_p.paragraph_format.space_before = Pt(6)
    for i, line in enumerate(brand["contact"]):
        if i > 0:
            contact_p.add_run().add_break()
        r = contact_p.add_run(line)
        _set_run_font(r, size=config.COVER_CONTACT_SIZE)

    # --- Neue Sektion für Innenseiten (eigener Header mit Logo) ------
    new_section = doc.add_section(WD_SECTION.NEW_PAGE)
    _set_page_geometry(new_section)
    _add_header_logo(new_section, assets_dir, brand)
    _add_footer_page_number(new_section)

    # --- Persönliche Daten --------------------------------------------
    # Struktur an die Vorlage angelehnt: EIN Absatz für den gesamten Block,
    # Felder durch Zeilenumbrüche (<w:br/>) getrennt. Label und Wert durch
    # EINEN Tab an einem fest definierten Tabstopp getrennt (siehe
    # config.PERSONAL_DATA_TAB_STOP) -- das ergibt eine garantiert saubere
    # Spaltenausrichtung unabhängig von der jeweiligen Labellänge; die
    # Vorlage selbst nutzt zwei variable Standard-Tabs, was zu leicht
    # unterschiedlichen Spaltenpositionen führen kann. Zeilenabstand
    # 1,5-fach (Quelle: <w:spacing w:line="360" w:lineRule="auto"/>).
    _add_heading(doc, "Persönliche Daten", level=1)
    pd_p = doc.add_paragraph()
    pd_p.paragraph_format.space_before = Pt(0)
    pd_p.paragraph_format.line_spacing = config.PERSONAL_DATA_LINE_SPACING
    pd_p.paragraph_format.tab_stops.add_tab_stop(config.PERSONAL_DATA_TAB_STOP)
    pd_fields = config.BLIND_PERSONAL_DATA_FIELDS if blind else config.PERSONAL_DATA_FIELDS
    for i, (key, label) in enumerate(pd_fields):
        if i > 0:
            pd_p.add_run().add_break()
        label_run = pd_p.add_run(label + "\t")
        _set_run_font(label_run, size=config.BODY_SIZE, bold=True)
        value_run = pd_p.add_run(str(data.get(key, "")))
        _set_run_font(value_run, size=config.BODY_SIZE)

    ip_p = doc.add_paragraph()
    ip_p.paragraph_format.space_before = Pt(12)
    ip_label = ip_p.add_run("Interviewpartner: ")
    _set_run_font(ip_label, size=config.BODY_SIZE, bold=True)
    ip_value = ip_p.add_run(data.get("interviewpartner", ""))
    _set_run_font(ip_value, size=config.BODY_SIZE)

    # --- Inhaltliche Abschnitte ----------------------------------------
    rolle = data.get("rolle", "")
    # Über die kanonische Gesamtreihenfolge iterieren und nach den aktiven Keys
    # filtern -- so stimmt die Reihenfolge unabhängig vom Positionstyp (nur eine der
    # beiden vierten Rubriken ist aktiv). None = Standard (Führungsposition).
    active_keys = (section_keys if section_keys is not None
                   else [e["key"] for e in config.PROFILE_SECTIONS])
    for entry in config.ALL_SECTIONS_ORDERED:
        if entry["key"] not in active_keys:
            continue
        heading_text = entry["heading"].format(rolle=rolle)
        _add_heading(doc, heading_text, level=entry["level"])
        text = sections.get(entry["key"], "")
        if text:
            _add_body_paragraphs(doc, text)

    # --- Abschluss: Monat/Jahr, Berater, Agentur (wie in der Vorlage) ---
    _monate = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
               "August", "September", "Oktober", "November", "Dezember"]
    heute = datetime.date.today()
    datum_p = doc.add_paragraph()
    datum_p.paragraph_format.space_before = Pt(24)
    _set_run_font(datum_p.add_run(f"{_monate[heute.month - 1]} {heute.year}"),
                  size=config.BODY_SIZE)
    berater = (data.get("interviewpartner") or "").strip()
    sig_p = doc.add_paragraph()
    sig_p.paragraph_format.space_before = Pt(6)
    sig_lines = ([berater] if berater else []) + [brand["name"]]
    for i, line in enumerate(sig_lines):
        if i > 0:
            sig_p.add_run().add_break()
        _set_run_font(sig_p.add_run(line), size=config.BODY_SIZE)

    doc.save(output_path)
    return output_path
