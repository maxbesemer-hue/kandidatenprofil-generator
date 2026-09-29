"""
Extrahiert automatisch ein Kandidatenfoto aus einem hochgeladenen
Lebenslauf (.docx oder .pdf), damit es nicht zusätzlich manuell
hochgeladen werden muss.

Heuristik: Von allen im Lebenslauf eingebetteten Bildern, die eine
Mindestgröße überschreiten (um Icons, Aufzählungszeichen-Grafiken o.ä.
auszuschließen) und kein extremes Seitenverhältnis haben (kein Banner,
kein dünner Trennstrich), wird das wahrscheinlichste Bewerbungsfoto
gewählt. Dabei gilt:

1. Ein Bewerbungsfoto ist praktisch immer QUADRATISCH oder HOCHFORMATIG
   (Passbild-/Portraitformat). Deko- und Bannerbilder sind dagegen breit
   im Querformat. Deshalb werden zuerst nur quadratische/hochformatige
   Bilder betrachtet; nur wenn es davon keines gibt, wird ersatzweise auch
   ein Querformat-Bild zugelassen.
2. Innerhalb der bevorzugten Gruppe wird das flächengrößte Bild gewählt.

Diese Portrait-Bevorzugung verhindert den Fall, dass ein großes Deko-
Querformatbild (z.B. ein Stockfoto als Gestaltungselement) das kleinere,
echte Bewerbungsfoto verdrängt.

WICHTIG: Diese Heuristik ist NICHT unfehlbar -- bei Lebensläufen mit
mehreren portrait-förmigen Bildern (z.B. ein hochformatiges Firmenlogo)
kann weiterhin das falsche Bild gewählt werden. Die App zeigt das erkannte
Foto deshalb immer zur Kontrolle an und erlaubt, stattdessen manuell ein
anderes Foto hochzuladen -- es wird nie automatisch final übernommen, ohne
dass es sichtbar war.
"""

import io
from PIL import Image

MIN_DIMENSION_PX = 120   # schließt kleine Icons/Symbole aus
MIN_ASPECT = 0.4         # schließt extrem schmale/hohe Grafiken aus
MAX_ASPECT = 2.2         # schließt extrem breite Banner/Trennlinien aus
# Bis zu diesem Seitenverhältnis (Breite/Höhe) gilt ein Bild als
# quadratisch/hochformatig und damit als bevorzugtes Portrait. Ein echtes
# Bewerbungsfoto ist quadratisch (1.0) oder hochkant (3:4=0.75, 2:3=0.67);
# etwas Toleranz nach oben fängt leicht breite Zuschnitte ab.
PORTRAIT_MAX_ASPECT = 1.15


def _qualifies_as_photo(image_bytes: bytes):
    """Prüft Mindestgröße und Seitenverhältnis. Gibt (width, height) bei
    Eignung zurück, sonst None."""
    try:
        with Image.open(io.BytesIO(image_bytes)) as img:
            w, h = img.size
    except Exception:
        return None
    if w < MIN_DIMENSION_PX or h < MIN_DIMENSION_PX:
        return None
    aspect = w / h
    if not (MIN_ASPECT <= aspect <= MAX_ASPECT):
        return None
    return (w, h)


def _pick_best(candidates):
    """Wählt aus geeigneten Bildern das wahrscheinlichste Bewerbungsfoto.

    candidates: Liste von (width, height, bytes).

    Bevorzugt quadratische/hochformatige Bilder (Portraitformat); nur wenn
    es davon keines gibt, wird auf Querformat zurückgegriffen. Innerhalb der
    gewählten Gruppe entscheidet die größte Fläche."""
    if not candidates:
        return None
    portrait = [c for c in candidates if (c[0] / c[1]) <= PORTRAIT_MAX_ASPECT]
    pool = portrait if portrait else candidates
    best = max(pool, key=lambda c: c[0] * c[1])
    return best[2]


def _extract_from_docx(file_bytes: bytes):
    from docx import Document
    doc = Document(io.BytesIO(file_bytes))
    candidates = []
    for rel in doc.part.rels.values():
        if "image" not in rel.reltype:
            continue
        try:
            data = rel.target_part.blob
        except Exception:
            continue
        dims = _qualifies_as_photo(data)
        if dims is None:
            continue
        candidates.append((dims[0], dims[1], data))
    return _pick_best(candidates)


def _extract_from_pdf(file_bytes: bytes):
    from pypdf import PdfReader
    reader = PdfReader(io.BytesIO(file_bytes))
    candidates = []
    for page in reader.pages:
        for img in getattr(page, "images", []):
            try:
                data = img.data
            except Exception:
                continue
            dims = _qualifies_as_photo(data)
            if dims is None:
                continue
            candidates.append((dims[0], dims[1], data))
    return _pick_best(candidates)


def extract_photo_from_cv(file_bytes: bytes, filename: str):
    """Gibt die Bilddaten des wahrscheinlichsten Kandidatenfotos aus dem
    Lebenslauf zurück (als bytes), oder None, wenn kein geeignetes Bild
    gefunden wurde (z.B. bei .txt-Dateien, textbasierten PDFs ohne
    eingebettete Bilder, oder wenn kein Bild die obigen Kriterien erfüllt)."""
    name = filename.lower()
    if name.endswith(".docx"):
        return _extract_from_docx(file_bytes)
    if name.endswith(".pdf"):
        return _extract_from_pdf(file_bytes)
    return None
