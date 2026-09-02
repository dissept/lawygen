"""
extractor.py
Extrae texto plano de PDF y DOCX. No usa ningun servicio externo: todo corre
en local.
"""
import os
import re

PARTY_ROLE_PATTERNS = {
    "Demandado": r"\bdemandad[oa]s?\b|\bparte\s+demandada\b",
    "Demandante": r"\bdemandantes?\b|\bparte\s+demandante\b|\bactor(?:a)?\b",
}

def extract_text_from_pdf(path, max_pages=60):
    """Devuelve el texto de un PDF. Limita paginas para no colgar el proceso
    con documentos enormes (escaneados de cientos de paginas)."""
    text_parts = []
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            for i, page in enumerate(pdf.pages):
                if i >= max_pages:
                    break
                try:
                    t = page.extract_text() or ""
                except Exception:
                    t = ""
                text_parts.append(t)
    except Exception as e:
        return "", f"Error leyendo PDF ({os.path.basename(path)}): {e}"
    return "\n".join(text_parts), None


def extract_text_from_docx(path):
    try:
        import docx
        d = docx.Document(path)
        paras = [p.text for p in d.paragraphs]
        return "\n".join(paras), None
    except Exception as e:
        return "", f"Error leyendo DOCX ({os.path.basename(path)}): {e}"


def extract_text_from_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return extract_text_from_pdf(path)
    if ext in (".docx",):
        return extract_text_from_docx(path)
    if ext in (".txt",):
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read(), None
        except Exception as e:
            return "", f"Error leyendo TXT ({os.path.basename(path)}): {e}"
    return "", None  # tipo no soportado, se ignora silenciosamente


def list_documents(folder_path):
    """Devuelve lista de rutas absolutas a documentos soportados dentro de
    una carpeta (no recursivo en subcarpetas de segundo nivel para mantener
    la logica simple: 1 carpeta = 1 asunto)."""
    supported = (".pdf", ".docx", ".txt")
    docs = []
    for root, _dirs, files in os.walk(folder_path):
        for f in files:
            if not f.lower().endswith(supported) or f.startswith("~$"):
                continue
            if f.endswith("_Resumen.docx") or f.endswith("_Esquema.docx") or f.endswith("_Resumen_Esquema.docx"):
                continue  # excluye los resultados generados por esta misma app
            docs.append(os.path.join(root, f))
    return docs


def pick_main_pdf(folder_path):
    """Elige el PDF 'principal' de la carpeta: heuristica = el PDF de mayor
    tamano de archivo (suele ser la demanda/sentencia principal frente a
    notificaciones cortas). Devuelve None si no hay PDFs."""
    pdfs = [p for p in list_documents(folder_path) if p.lower().endswith(".pdf")]
    if not pdfs:
        return None
    return max(pdfs, key=lambda p: os.path.getsize(p))


def party_role_for_path(path, root_folder):
    """Si el documento vive dentro de una subcarpeta tipo 'Demandado' o
    'Demandante' (a cualquier profundidad bajo la carpeta del asunto),
    devuelve ese rol ('Demandado'/'Demandante'); si no, None. Permite
    separar qué documentos pertenecen a cada parte cuando el despacho
    organiza así sus carpetas."""
    try:
        rel = os.path.relpath(os.path.dirname(path), root_folder)
    except ValueError:
        return None
    parts = [p for p in rel.split(os.sep) if p not in ("", ".")]
    for part in parts:
        low = part.lower()
        for role, pat in PARTY_ROLE_PATTERNS.items():
            if re.search(pat, low, re.IGNORECASE):
                return role
    return None
