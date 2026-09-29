"""Hilfsfunktionen zum Einlesen von Text aus verschiedenen Dateiformaten.

Unterstützt: .txt (Interviewprotokolle aus ECHO KI), .docx und .pdf (Lebensläufe).
"""

import io


def extract_text(uploaded_file) -> str:
    """uploaded_file: Streamlit UploadedFile-Objekt (hat .name und .read())"""
    name = uploaded_file.name.lower()
    data = uploaded_file.read()

    if name.endswith(".txt"):
        return data.decode("utf-8", errors="replace")

    if name.endswith(".docx"):
        from docx import Document
        doc = Document(io.BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs)

    if name.endswith(".pdf"):
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(data))
        return "\n".join(page.extract_text() or "" for page in reader.pages)

    raise ValueError(f"Nicht unterstütztes Dateiformat: {name}. "
                      f"Unterstützt werden .txt, .docx und .pdf.")
