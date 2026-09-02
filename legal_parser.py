"""
legal_parser.py
Extraccion basada en reglas (regex + heuristicas) de:
  - Juzgado
  - Partes
  - Procedimiento(s)
  - Resumen (extractivo)
  - Esquema (estructura del documento)

No usa ninguna API externa: 100% gratis y offline. Si un campo no se puede
determinar con confianza razonable, se marca "revisar" en vez de inventarlo.
"""
import re
from collections import Counter

REVISAR = "revisar"

# Fragmento reutilizable para el tratamiento honorifico delante de un nombre
# propio: admite "D.", "Dña.", "D.ª", "Dª." (orden de puntos/ordinal
# variable segun el escrito) y variantes sin punto final.
NAME_PREFIX = r"(?:Dña\.?|D\.?ª\.?|D\.?)\s*"

# ---------------------------------------------------------------------------
# JUZGADO
# ---------------------------------------------------------------------------
JUZGADO_PATTERNS = [
    r"(JUZGADO\s+DE\s+[A-ZÁÉÍÓÚÑ ,°ºª0-9]+(?:N[ºO°]\s?\d+)?)",
    r"(JUZGADO\s+[A-ZÁÉÍÓÚÑ ]+\s+N[ºO°]?\s?\d+\s+DE\s+[A-ZÁÉÍÓÚÑ ]+)",
    r"(TRIBUNAL\s+[A-ZÁÉÍÓÚÑ ,°ºª0-9]+)",
    r"(AUDIENCIA\s+PROVINCIAL\s+DE\s+[A-ZÁÉÍÓÚÑ ]+)",
    r"(SALA\s+DE\s+LO\s+[A-ZÁÉÍÓÚÑ ]+\s+DE[A-ZÁÉÍÓÚÑ ]*)",
]

def extract_juzgado(text):
    text = _flat(text)
    for pat in JUZGADO_PATTERNS:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            candidate = " ".join(m.group(1).split())
            # corta antes de que se cuele la direccion postal del membrete
            candidate = re.split(
                r"\s+(?:Calle|C/|Avda\.?|Avenida|Plaza|Pza\.?|Paseo|Procedimiento|Autos|Expediente|"
                r"Demandante|Demandado|Actor|Juicio|Diligencias)\b",
                candidate, flags=re.IGNORECASE
            )[0]
            return candidate.strip(" .,:;")
    return REVISAR


def extract_juzgado_multi(doc_texts):
    """Vota entre los documentos de la carpeta: el juzgado correcto es el
    que aparece de forma consistente en varios documentos del mismo asunto."""
    counts = Counter()
    for text in doc_texts:
        j = extract_juzgado(text)
        if j != REVISAR:
            counts[j] += 1
    if not counts:
        return REVISAR
    return counts.most_common(1)[0][0]


# ---------------------------------------------------------------------------
# PROCEDIMIENTO
# ---------------------------------------------------------------------------
PROCEDIMIENTO_PATTERNS = [
    r"(PROCEDIMIENTO[:\s]+[^\n\.]{3,110})",
    r"(JUICIO\s+(?:ORDINARIO|VERBAL|MONITORIO|CAMBIARIO|EJECUTIVO)[^\n\.]{0,60})",
    r"(DILIGENCIAS\s+PREVIAS[^\n\.]{0,60})",
    r"(AUTOS\s+(?:DE\s+)?[A-ZÁÉÍÓÚÑ ]+\s+N[ºO°]?\s?\d+[/\-]\d+)",
    r"(EXPEDIENTE\s+N[ºO°]?\s?\d+[/\-]\d+)",
    r"(PROCEDIMIENTO\s+ORDINARIO\s+N[ºO°]?\s?\d+[/\-]\d+)",
    r"(EJECUCI[ÓO]N\s+[A-ZÁÉÍÓÚÑ ]*N[ºO°]?\s?\d+[/\-]\d+)",
    r"(RECURSO\s+DE\s+[A-ZÁÉÍÓÚÑ]+\s+N[ºO°]?\s?\d+[/\-]\d+)",
]

def _balance_parens(val):
    """Si la captura corta a mitad de un parentesis (ej. '(Procedimiento'),
    lo completa o lo elimina para que quede legible."""
    if val.count("(") > val.count(")"):
        idx = val.rfind("(")
        # intenta cerrar con una palabra mas si parece razonable, si no, recorta
        return val[:idx].strip(" .,:;")
    return val

def extract_procedimientos(text):
    """Devuelve una lista de procedimientos distintos detectados en el texto
    (puede haber mas de uno: p.ej. varios expedientes en la misma carpeta)."""
    text = clean_text(text)
    found = []
    seen = set()
    for pat in PROCEDIMIENTO_PATTERNS:
        for m in re.finditer(pat, text, re.IGNORECASE):
            val = " ".join(m.group(1).split()).strip(" .,:;")
            val = _balance_parens(val)
            key = val.lower()
            if key not in seen and len(val) > 4:
                seen.add(key)
                found.append(val)
    if not found:
        return [REVISAR]
    return found


CASE_NUMBER_PATTERN = r"\b(\d{1,6})\s?[/\-]\s?(\d{2,4})\b"
CASE_TYPE_PATTERN = (
    r"(PROCEDIMIENTO\s+(?:ORDINARIO|VERBAL|MONITORIO|CAMBIARIO|EJECUTIVO)"
    r"|JUICIO\s+(?:ORDINARIO|VERBAL)"
    r"|DILIGENCIAS\s+PREVIAS"
    r"|EJECUCI[ÓO]N(?:\s+DE\s+T[ÍI]TULOS?\s+JUDICIALES)?"
    r"|PIEZA\s+DE\s+MEDIDAS\s+CAUTELARES"
    r"|EXPEDIENTE"
    r"|RECURSO\s+DE\s+[A-ZÁÉÍÓÚÑ]+"
    r"|AUTOS)"
)

CITATION_PREFIX_PATTERN = r"\b(?:STS|SAP|SJM|SAN|STSJ|TC|ATS|AAP|RJ|ROJ)\b"

def extract_procedimientos_multi(doc_texts):
    """Identifica los procedimientos por su NUMERO real de autos/expediente
    (ej. 936/2024), agrupando todas las menciones de ese mismo numero en
    los distintos documentos de la carpeta en una unica fila — en vez de
    listar cada mencion suelta como si fuera un procedimiento distinto.

    Descarta agresivamente los falsos positivos: las sentencias citan
    constantemente jurisprudencia con el mismo formato numero/año (p.ej.
    "la STS 257/2008 establece que...") que NO es el numero de autos de
    este expediente. Solo se acepta un numero como procedimiento real si
    tiene una etiqueta de tipo (Procedimiento/Autos/Diligencias/etc.) MUY
    cerca delante, y ademas aparece repetido en el documento o esta cerca
    del principio (cabecera), que es donde SIEMPRE figura el numero de
    autos propio del escrito."""
    labels_by_number = {}
    order = []
    for text in doc_texts:
        candidates = []  # (norm, label, position)
        counts = Counter()
        for m in re.finditer(CASE_NUMBER_PATTERN, text):
            main_num, year = m.group(1), m.group(2)
            if len(main_num) < 1 or len(year) < 2:
                continue
            # normaliza el año a 4 digitos (24 -> 2024) para que la misma
            # causa citada como "936/24" en un sitio y "936/2024" en otro
            # se reconozca como el MISMO procedimiento, no como dos.
            year_full = ("20" + year) if len(year) == 2 else year
            norm = f"{int(main_num)}/{year_full}"
            if len(re.sub(r'\D', '', norm)) < 5:
                continue  # numero demasiado corto para ser fiable (ruido)
            num_display = f"{main_num}/{year}"

            # descarta citas de jurisprudencia (STS, SAP, ROJ...) justo antes del numero
            pre = text[max(0, m.start() - 12):m.start()]
            if re.search(CITATION_PREFIX_PATTERN, pre, re.IGNORECASE):
                continue

            # la etiqueta de tipo debe estar INMEDIATAMENTE antes (ventana corta):
            # evita que un "Procedimiento" mencionado varias frases antes
            # "adopte" por error un numero de cita jurisprudencial posterior
            context = text[max(0, m.start() - 35):m.start()]
            type_m = re.search(CASE_TYPE_PATTERN, context, re.IGNORECASE)
            if not type_m:
                continue  # sin etiqueta de tipo pegada al numero -> no fiable, se descarta

            label = " ".join(f"{type_m.group(1).title()} {num_display}".split())
            candidates.append((norm, label, m.start()))
            counts[norm] += 1

        # ademas de la etiqueta de tipo, exige que el numero se repita en el
        # documento (cabecera + pie de pagina, tipico) o que aparezca cerca
        # del inicio (primeros 700 caracteres, donde va la cabecera) — asi
        # se descartan menciones aisladas de jurisprudencia sin prefijo
        # reconocible que igualmente hubieran colado una etiqueta de tipo.
        accepted = {norm for norm, _, pos in candidates if counts[norm] >= 2 or pos < 700}

        for norm, label, _ in candidates:
            if norm not in accepted:
                continue
            if norm not in labels_by_number:
                labels_by_number[norm] = label
                order.append(norm)
            elif len(label) > len(labels_by_number[norm]):
                labels_by_number[norm] = label  # prefiere la etiqueta mas informativa
    if not labels_by_number:
        return [REVISAR]
    return [labels_by_number[n] for n in order]


# ---------------------------------------------------------------------------
# PARTES
# ---------------------------------------------------------------------------
PARTES_LABELS = [
    "DEMANDANTE", "DEMANDADO", "DEMANDADA", "ACTOR", "ACTORA",
    "QUERELLANTE", "QUERELLADO", "RECURRENTE", "RECURRIDO",
    "EJECUTANTE", "EJECUTADO", "APELANTE", "APELADO",
]

def extract_partes(text):
    partes = []
    for label in PARTES_LABELS:
        # Requiere que la etiqueta vaya seguida de ":" y no sea parte de otra
        # palabra (evita falsos positivos como "la demandada incumplio...").
        # No depende de saltos de linea: el texto puede venir ya aplanado.
        pat = rf"(?<![A-Za-zÁÉÍÓÚÑáéíóúñ]){label}\s*:\s*({NAME_PREFIX}[A-ZÁÉÍÓÚÑ][^\n\.]{{2,80}})"
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = " ".join(m.group(1).split()).strip(" .,:;")
            # si se colo la siguiente etiqueta o un encabezado de seccion
            # (ej. sin punto de por medio), recorta ahi
            boundary_labels = PARTES_LABELS + ["ABOGADO", "ABOGADA", "LETRADO", "LETRADA",
                                                "PROCURADOR", "PROCURADORA"] + ESQUEMA_HEADERS
            next_label_m = re.search(
                r"\b(?:" + "|".join(boundary_labels) + r")\b\s*:?",
                val, re.IGNORECASE
            )
            if next_label_m:
                val = val[:next_label_m.start()].strip(" .,:;")
            partes.append(f"{label.capitalize()}: {val}")
    if partes:
        return " | ".join(partes)
    # fallback: patron "X contra Y" / "X vs Y" — se admiten particulas de
    # union en apellidos compuestos ("de la Joya", "del Campo"...), muy
    # habituales en español, para no truncar el nombre a mitad.
    NAME_CHUNK = (
        NAME_PREFIX +
        r"[A-ZÁÉÍÓÚÑ][a-záéíóúñ\.]+"
        r"(?:\s+(?:de|del|la|las|los|y)\b|\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ\.]+){1,6}"
    )
    m = re.search(rf"({NAME_CHUNK})\s+(?:contra|vs\.?)\s+({NAME_CHUNK})", text)
    if m:
        return f"{m.group(1).strip()} contra {m.group(2).strip()}"
    return REVISAR


def extract_partes_multi(doc_texts):
    """Recorre los documentos de la carpeta y usa el primero que tenga
    etiquetas claras de partes (el escrito principal suele tenerlas; las
    notificaciones/anexos normalmente no)."""
    for text in doc_texts:
        p = extract_partes(text)
        if p != REVISAR:
            return p
    return REVISAR


# ---------------------------------------------------------------------------
# REPRESENTANTES (abogados/procuradores de cada parte)
# ---------------------------------------------------------------------------
REPRESENTANTE_LABELS = [
    "ABOGADO", "ABOGADA", "LETRADO", "LETRADA",
    "PROCURADOR", "PROCURADORA",
]

def extract_representantes(text):
    """Busca etiquetas tipo 'Abogado:', 'Procurador:', etc. y las devuelve
    como 'Rol: Nombre' unidas por ' | '. Marca 'revisar' si no encuentra
    ninguna con confianza razonable."""
    found = []
    seen = set()
    for label in REPRESENTANTE_LABELS:
        pat = rf"(?<![A-Za-zÁÉÍÓÚÑáéíóúñ]){label}\s*:\s*({NAME_PREFIX}[A-ZÁÉÍÓÚÑ][^\n\.]{{2,80}})"
        for m in re.finditer(pat, text, re.IGNORECASE):
            val = " ".join(m.group(1).split()).strip(" .,:;")
            boundary_labels = REPRESENTANTE_LABELS + PARTES_LABELS + ESQUEMA_HEADERS
            next_label_m = re.search(
                r"\b(?:" + "|".join(boundary_labels) + r")\b\s*:?",
                val, re.IGNORECASE
            )
            if next_label_m:
                val = val[:next_label_m.start()].strip(" .,:;")
            if len(val) < 3:
                continue
            entry = f"{label.capitalize()}: {val}"
            key = entry.lower()
            if key not in seen:
                seen.add(key)
                found.append(entry)
    if found:
        return " | ".join(found)

    # Fallback narrativo: los escritos judiciales españoles rara vez usan
    # etiquetas "Procurador:"; lo habitual es prosa tipo "el Procurador de
    # los Tribunales D. Fernando Pérez Cruz en nombre y representación de
    # D. Francisco Javier Montero Bartolomesanz". Busca ese patrón.
    NAME_CHUNK = (
        NAME_PREFIX +
        r"[A-ZÁÉÍÓÚÑ][a-záéíóúñ\.]+"
        r"(?:\s+(?:de|del|la|las|los|y)\b|\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ\.]+){1,6}"
    )
    narrative_pat = re.compile(
        rf"(?i:Procurador(?:a)?)(?:(?i:\s+de\s+los\s+Tribunales))?\s+({NAME_CHUNK})"
        rf"(?i:\s+en\s+(?:nombre\s+y\s+)?representaci[oó]n\s+de)\s+({NAME_CHUNK})"
    )
    for m in narrative_pat.finditer(text):
        rep_name, party_name = m.group(1).strip(" .,:;"), m.group(2).strip(" .,:;")
        entry = f"Procurador de {party_name}: {rep_name}"
        key = entry.lower()
        if key not in seen:
            seen.add(key)
            found.append(entry)

    # Segunda direccion narrativa: "D. Adrián ... representado por la
    # Procuradora Dña. Silvia..." (muy habitual para la parte demandada).
    narrative_pat2 = re.compile(
        rf"({NAME_CHUNK})(?i:\s+representad[oa]s?\s+por\s+(?:el|la)\s+Procurador(?:a)?)\s+({NAME_CHUNK})"
    )
    for m in narrative_pat2.finditer(text):
        party_name, rep_name = m.group(1).strip(" .,:;"), m.group(2).strip(" .,:;")
        entry = f"Procurador de {party_name}: {rep_name}"
        key = entry.lower()
        if key not in seen:
            seen.add(key)
            found.append(entry)

    # "Letrado" (el abogado que dirige/asiste, distinto del procurador que
    # representa formalmente) — importante no perderlo: "asistido del/por
    # el Letrado D. X" o "bajo la dirección del Letrado D. X".
    letrado_pat = re.compile(
        rf"(?i:asistid[oa]s?\s+(?:del|por\s+el)|bajo\s+la\s+direcci[oó]n\s+(?:del|de\s+la)|"
        rf"defendid[oa]s?\s+por\s+el)\s+(?i:Letrad[oa])\s+({NAME_CHUNK})"
    )
    for m in letrado_pat.finditer(text):
        rep_name = m.group(1).strip(" .,:;")
        entry = f"Letrado: {rep_name}"
        key = entry.lower()
        if key not in seen:
            seen.add(key)
            found.append(entry)

    if found:
        return " | ".join(found)
    return REVISAR


def extract_representantes_multi(doc_texts):
    for text in doc_texts:
        r = extract_representantes(text)
        if r != REVISAR:
            return r
    return REVISAR


# ---------------------------------------------------------------------------
# MATERIA (civil, penal, laboral, contencioso-administrativo, mercantil)
# ---------------------------------------------------------------------------
MATERIA_KEYWORDS = [
    ("Penal", [
        "diligencias previas", "código penal", "codigo penal", "querella",
        "denunciado", "fiscalía", "fiscalia", "instrucción penal",
        "juicio de faltas", "procedimiento abreviado", "delito",
        "gabinete de la acusación", "atestado policial",
    ]),
    ("Laboral", [
        "juzgado de lo social", "estatuto de los trabajadores", "despido",
        "reclamación de cantidad laboral", "conciliación laboral", "seguridad social",
        "expediente de regulación de empleo",
    ]),
    ("Contencioso-Administrativo", [
        "juzgado de lo contencioso", "sala de lo contencioso",
        "recurso contencioso-administrativo", "acto administrativo",
        "administración pública", "administracion publica",
    ]),
    ("Mercantil", [
        "juzgado de lo mercantil", "concurso de acreedores", "sociedad mercantil",
        "propiedad industrial", "competencia desleal", "registro mercantil",
    ]),
    ("Civil", [
        "ley de enjuiciamiento civil", "juicio ordinario", "juicio verbal",
        "arrendamiento", "divorcio", "separación matrimonial", "sucesiones",
        "reclamación de cantidad", "reclamacion de cantidad", "desahucio",
        "custodia", "herencia",
    ]),
]

def extract_materia_multi(doc_texts):
    scores = Counter()
    for text in doc_texts:
        low = _flat(text).lower()
        for materia, keywords in MATERIA_KEYWORDS:
            for kw in keywords:
                if kw in low:
                    scores[materia] += 1
    if not scores:
        return REVISAR
    return scores.most_common(1)[0][0]


# ---------------------------------------------------------------------------
# OBJETO / CAUSA JUZGADA
# ---------------------------------------------------------------------------
# Encabezados de seccion que NO deben colarse dentro del contenido de OTRO
# campo (Objeto, Fundamentos...): si aparecen dentro de la ventana de texto
# que estamos leyendo, es que ya cruzamos a la siguiente seccion.
SECTION_BOUNDARY_HEADERS = [
    r"\bHECHOS\b", r"\bFUNDAMENTOS\s+DE\s+DERECHO\b", r"\bFUNDAMENTOS?\s+JUR[IÍ]DICOS\b",
    r"\bSUPLICO\b", r"\bOTROS[IÍ]\s+DIGO\b", r"\bFALLO\b", r"\bANTECEDENTES\s+DE\s+HECHO\b",
    r"\bPETICIONES?\b",
]


def _cut_before_next_header(snippet):
    """Si dentro del fragmento aparece el inicio de otra seccion (buscando a
    partir del caracter 5, para no confundir con el propio encabezado que ya
    usamos para localizar esta seccion), recorta ahi: ese texto ya pertenece
    a la siguiente seccion, no a la que estamos leyendo."""
    earliest = None
    for pat in SECTION_BOUNDARY_HEADERS:
        m = re.search(pat, snippet[5:], re.IGNORECASE)
        if m and (earliest is None or m.start() < earliest):
            earliest = m.start()
    return snippet[:earliest + 5] if earliest is not None else snippet


def _drop_incomplete_tail(sentences):
    """Si el ultimo fragmento no termina en puntuacion de cierre de frase,
    la ventana de texto lo corto a mitad de palabra/frase -- se descarta en
    vez de devolver algo incompleto."""
    if sentences and not sentences[-1].rstrip().endswith((".", "!", "?", "…")):
        sentences = sentences[:-1]
    return sentences

OBJETO_HEADERS = [
    r"OBJETO\s+DEL\s+PROCEDIMIENTO",
    r"OBJETO\s+DE\s+LA\s+DEMANDA",
    r"^\s*OBJETO\b",
    r"SUPLICO",
    r"PETICIONES?\b",
    r"solicit[oó]\s+que\s+se\s+dictar?a?\s+[Ss]entencia\s+por\s+la\s+que",
    r"\bFALLO\b",
]

def extract_objeto(text):
    text = clean_text(text)
    for pat in OBJETO_HEADERS:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            snippet = text[m.end():m.end() + 260].strip(" .:;-")
            snippet = _cut_before_next_header(snippet)
            parts = re.split(r"(?<=[\.\!\?])\s", snippet)
            parts = _drop_incomplete_tail(parts)
            snippet = " ".join(parts[:2]).strip()
            if len(snippet) > 15:
                return snippet[:280]
    return REVISAR


def extract_objeto_multi(doc_texts):
    for text in doc_texts:
        o = extract_objeto(text)
        if o != REVISAR:
            return o
    return REVISAR


# ---------------------------------------------------------------------------
# PLAZOS
# ---------------------------------------------------------------------------
NUM_WORD = (
    r"(?:\d+|un|una|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez|once|doce|"
    r"trece|catorce|quince|dieciséis|dieciseis|diecisiete|dieciocho|diecinueve|"
    r"veinte|treinta|cuarenta|cincuenta|sesenta)"
)
PLAZO_PATTERNS = [
    rf"plazo\s+de\s+({NUM_WORD}\s+(?:d[ií]as?|meses?|a[ñn]os?)[^\n\.,]{{0,40}})",
    rf"t[ée]rmino\s+de\s+({NUM_WORD}\s+(?:d[ií]as?|meses?|a[ñn]os?)[^\n\.,]{{0,40}})",
    rf"en\s+el\s+plazo\s+improrrogable\s+de\s+({NUM_WORD}\s+(?:d[ií]as?|meses?)[^\n\.,]{{0,40}})",
]

def extract_plazos(text):
    found = []
    seen = set()
    for pat in PLAZO_PATTERNS:
        for m in re.finditer(pat, text, re.IGNORECASE):
            val = " ".join(m.group(0).split()).strip(" .,:;")
            key = val.lower()
            if key not in seen and len(val) > 5:
                seen.add(key)
                found.append(val)
    if not found:
        return REVISAR
    return "; ".join(found[:5])


def extract_plazos_multi(doc_texts):
    all_found = []
    seen = set()
    for text in doc_texts:
        p = extract_plazos(text)
        if p != REVISAR:
            for item in p.split("; "):
                key = item.lower()
                if key not in seen:
                    seen.add(key)
                    all_found.append(item)
    if not all_found:
        return REVISAR
    return "; ".join(all_found[:6])


# ---------------------------------------------------------------------------
# FUNDAMENTOS DE DERECHO (base jurídica/procesal alegada — sustituye a
# "Plazos" como campo de interés, que resultaba menos útil en la práctica)
# ---------------------------------------------------------------------------
FUNDAMENTOS_HEADERS = [
    r"FUNDAMENTOS\s+DE\s+DERECHO",
    r"FUNDAMENTOS?\s+JUR[IÍ]DICOS",
    r"\bALEGACIONES\b",
    r"EXCEPCI[OÓ]N\s+PROCESAL",
    r"\bOTROS[IÍ]\s+DIGO\b",
    r"\bDIGO\b",
]

def extract_fundamentos_derecho(text, max_bullets=4):
    """Busca la sección de fundamentos jurídicos/procesales (Fundamentos de
    Derecho, Alegaciones, Excepción Procesal, Otrosí Digo...) y devuelve
    las primeras frases como viñetas — el equivalente para la base legal
    de lo que extract_objeto hace para la petición concreta."""
    text = clean_text(text)
    for pat in FUNDAMENTOS_HEADERS:
        m = re.search(pat, text, re.IGNORECASE)
        if not m:
            continue
        snippet = text[m.end():m.end() + 500].strip(" .:;-")
        snippet = _cut_before_next_header(snippet)
        sentences = re.split(r"(?<=[\.\!\?])\s+", snippet)
        sentences = _drop_incomplete_tail(sentences)
        bullets = [s.strip() for s in sentences if len(s.strip()) > 20][:max_bullets]
        if bullets:
            return bullets
    return [REVISAR]


def extract_fundamentos_derecho_multi(doc_texts):
    for text in doc_texts:
        b = extract_fundamentos_derecho(text)
        if b != [REVISAR]:
            return b
    return [REVISAR]



CUANTIA_PATTERNS = [
    r"cuant[ií]a[:\s]+([\d\.,]+\s*(?:euros|EUR|€)?)",
    r"por\s+la\s+cantidad\s+de\s+([\d\.,]+\s*(?:euros|EUR|€))",
    r"a\s+la\s+cantidad\s+de\s+([\d\.,]+\s*(?:euros|EUR|€))",
    r"abono\s+de\s+la\s+cantidad\s+de\s+([\d\.,]+\s*(?:euros|EUR|€))",
    r"cantidad\s+de\s+([\d\.,]+\s*(?:euros|EUR|€))\s+en\s+concepto\s+de\s+principal",
]
COSTAS_PATTERNS = [
    r"(con\s+(?:expresa\s+)?imposici[oó]n\s+de\s+costas[^\n\.]{0,60})",
    r"(sin\s+(?:expresa\s+)?imposici[oó]n\s+de\s+costas[^\n\.]{0,60})",
    r"(cada\s+parte\s+abonar[aá]\s+las\s+costas[^\n\.]{0,60})",
]

def extract_cuantia_costas(text):
    parts = []
    for pat in CUANTIA_PATTERNS:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            parts.append(f"Cuantía: {' '.join(m.group(1).split()).strip(' .,:;')}")
            break
    for pat in COSTAS_PATTERNS:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            parts.append(f"Costas: {' '.join(m.group(1).split()).strip(' .,:;')}")
            break
    if not parts:
        return REVISAR
    return " | ".join(parts)


def extract_cuantia_costas_multi(doc_texts):
    for text in doc_texts:
        c = extract_cuantia_costas(text)
        if c != REVISAR:
            return c
    return REVISAR


# ---------------------------------------------------------------------------
# ESTADO (sugerencia, no vinculante; el usuario ajusta con el desplegable)
# ---------------------------------------------------------------------------
ESTADOS_VALIDOS = ["Abierto", "En trámite", "Suspendido", "Archivado", "Resuelto"]

ESTADO_KEYWORDS = {
    "Resuelto": ["sentencia firme", "se resuelve", "fallo", "se declara", "cosa juzgada",
                 "sentencia n", "queda firme", "ejecutoria", "se ejecuta la sentencia"],
    "Archivado": ["archivo de las actuaciones", "se archiva", "sobreseimiento libre",
                  "archivado", "se tiene por desistido", "se declara caducado"],
    "Suspendido": ["se suspende", "suspensión del procedimiento", "en suspenso"],
    "En apelación": ["recurso de apelación", "se interpone recurso", "se eleva a la audiencia provincial",
                      "en apelación", "recurso de casación", "designando sección y ponente"],
    "En trámite": ["se admite a trámite", "traslado a las partes", "pendiente de",
                   "en tramitación", "se da traslado", "señalamiento para vista", "se cita a las partes"],
}

def suggest_estado(text):
    low = _flat(text).lower()
    for estado, keywords in ESTADO_KEYWORDS.items():
        for kw in keywords:
            if kw in low:
                return estado
    return "Abierto"  # valor por defecto conservador; el usuario debe revisar


def suggest_estado_multi(doc_texts):
    """Combina todos los documentos de la carpeta, ponderando los últimos
    algo más que los primeros: en un expediente, los documentos más
    recientes (a menudo los últimos en añadirse a la carpeta) suelen
    reflejar mejor el estado ACTUAL que el primer escrito presentado.
    No es perfecto (depende del orden de los archivos), pero es mucho
    mejor que quedarse solo con la primera coincidencia encontrada."""
    if not doc_texts:
        return "Abierto"
    scores = Counter()
    n = len(doc_texts)
    for i, text in enumerate(doc_texts):
        low = _flat(text).lower()
        weight = 1.0 + (i / max(n - 1, 1))  # de 1.0 (primer doc) a 2.0 (ultimo doc)
        for estado, keywords in ESTADO_KEYWORDS.items():
            for kw in keywords:
                if kw in low:
                    scores[estado] += weight
    if not scores:
        return "Abierto"
    return scores.most_common(1)[0][0]


# ---------------------------------------------------------------------------
# LIMPIEZA DE RUIDO (cabeceras, pies de página, códigos de verificación,
# texto corrupto que a veces generan los PDFs judiciales al extraer texto)
# ---------------------------------------------------------------------------
NOISE_PATTERNS = [
    r"C[oó]digo\s+seguro\s+de\s+verificaci[oó]n[^\n]{0,200}",
    r"Este\s+documento\s+es\s+una\s+copia\s+aut[eé]ntica[^\n]{0,200}",
    r"La\s+difusi[oó]n\s+del\s+texto\s+de\s+esta\s+resoluci[oó]n[^\.]{0,400}\.",
    r"firmado\s+electr[oó]nicamente\s+por[^\n\.]{0,80}",
    r"Tfno:?\s*\d[\d\s\-]{5,}",
    r"Fax:?\s*\d[\d\s\-]{5,}",
    r"NIG:?\s*[\d\.\-\/]+",
    r"\b[\w\.-]+@[\w\.-]+\.\w+\b",
    r"\bwww\.[^\s]+",
    r"\b\d{12,}\b",
    r"\b\d{1,3}\s+de\s+\d{1,3}\s*(?=\s|$)",  # numeración "1 de 2" de páginas
]

def _is_reversed_boilerplate(sentence):
    """Detecta lineas de pie de pagina que quedaron invertidas al extraer el
    texto del PDF (frecuente en escritos judiciales escaneados/firmados
    digitalmente). Se comprueba invirtiendo la frase y buscando patrones
    reconocibles (dominios, palabras clave de verificacion)."""
    reversed_s = sentence[::-1]
    return bool(re.search(r"www\.|\.org|\.es\b|c[oó]digo|verificaci[oó]n", reversed_s, re.IGNORECASE))

def clean_text(text):
    # Normaliza saltos de linea de Windows/Mac a "\n" simple.
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    for pat in NOISE_PATTERNS:
        text = re.sub(pat, " ", text, flags=re.IGNORECASE)
    # Colapsa espacios/tabulaciones DENTRO de una linea, pero conserva los
    # saltos de linea reales: son la unica señal de donde acaba una etiqueta
    # (ej. "Procurador: Fulano") y empieza lo siguiente. El re.sub(r"\s+", " ")
    # anterior los destruia a todos, y "Procurador: Fulano\nOBJETO DE LA
    # DEMANDA" se convertia en "Procurador: Fulano OBJETO DE LA DEMANDA" sin
    # ninguna frontera para detectar.
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

def _flat(text):
        """Aplana los saltos de linea a espacios SOLO para busquedas de
        palabra/frase que no dependen de una etiqueta (Materia, Estado, nombre
        del Juzgado): una frase clave puede aparecer partida en dos lineas por
        el ajuste de texto del PDF, y sin aplanar no se reconoceria."""
        return text.replace("\n", " ")


# ---------------------------------------------------------------------------
# RESUMEN (extractivo, por frecuencia de palabras - sin dependencias externas)
# ---------------------------------------------------------------------------
STOPWORDS_ES = set("""
de la que el en y a los del se las por un para con no una su al lo como más
pero sus le ya o este sí porque esta entre cuando muy sin sobre también me
hasta hay donde quien desde todo nos durante todos uno les ni contra otros
ese eso ante ellos e esto mí antes algunos qué unos yo otro otras otra él
tanto esa estos mucho quienes nada muchos cual poco ella estar estas algunas
algo nosotros mi mis tú te ti tu tus ellas nosotras vosotros vosotras os
mío mía míos mías tuyo tuya tuyos tuyas suyo suya suyos suyas nuestro nuestra
nuestros nuestras vuestro vuestra vuestros vuestras esos esas
""".split())

def summarize_bullets(text, max_bullets=5, max_chars=220):
    """Devuelve una LISTA de frases clave (viñetas), limpias de ruido, en
    vez de un unico bloque de texto. Cada viñeta se recorta a max_chars."""
    text = clean_text(text)
    if not text:
        return [REVISAR]

    split_pat = r"(?<=[\.\!\?])(?<!\bD\.)(?<!\bDña\.)(?<!\bSr\.)(?<!\bSra\.)(?<!\bD\.ª)\s+"
    sentences = re.split(split_pat, text)
    clean_sentences = []
    for s in sentences:
        s = s.strip(" .;-")
        if len(s) < 25 or len(s) > 500:
            continue
        if _is_reversed_boilerplate(s):
            continue
        digits = sum(1 for c in s if c.isdigit())
        if digits / max(len(s), 1) > 0.12:
            continue  # probable cabecera/membrete con codigos, direcciones, numeros de autos
        clean_sentences.append(s)

    if not clean_sentences:
        return [REVISAR]

    if len(clean_sentences) <= max_bullets:
        picked = clean_sentences
    else:
        words = re.findall(r"[a-záéíóúñA-ZÁÉÍÓÚÑ]{3,}", text.lower())
        freq = Counter(w for w in words if w not in STOPWORDS_ES)
        max_freq = max(freq.values()) if freq else 1
        for w in freq:
            freq[w] /= max_freq

        scored = []
        for idx, s in enumerate(clean_sentences):
            s_words = re.findall(r"[a-záéíóúñA-ZÁÉÍÓÚÑ]{3,}", s.lower())
            score = sum(freq.get(w, 0) for w in s_words) / max(len(s_words), 1)
            position_bonus = 0.5 if idx < 3 else 0.0
            scored.append((score + position_bonus, idx, s))

        top = sorted(scored, key=lambda x: x[0], reverse=True)[:max_bullets]
        picked = [s for _, _, s in sorted(top, key=lambda x: x[1])]

    bullets = []
    for s in picked:
        s = s[0].upper() + s[1:] if s else s
        if len(s) > max_chars:
            s = s[:max_chars].rsplit(" ", 1)[0] + "…"
        if not s.endswith((".", "…")):
            s += "."
        bullets.append(s)
    return bullets


def summarize(text, max_sentences=6):
    """Mantiene compatibilidad: version en un solo bloque de texto."""
    bullets = summarize_bullets(text, max_bullets=max_sentences)
    if bullets == [REVISAR]:
        return REVISAR
    return " ".join(bullets)


def _candidate_sentences(text):
    # evita partir frases justo despues de abreviaturas como "D." o "Dña."
    split_pat = r"(?<=[\.\!\?])(?<!\bD\.)(?<!\bDña\.)(?<!\bSr\.)(?<!\bSra\.)(?<!\bD\.ª)\s+"
    sentences = re.split(split_pat, text)
    out = []
    for s in sentences:
        s = s.strip(" .;-")
        if len(s) < 30 or len(s) > 400:
            continue
        if _is_reversed_boilerplate(s):
            continue
        digits = sum(1 for c in s if c.isdigit())
        if digits / max(len(s), 1) > 0.12:
            continue
        out.append(s)
    return out


def _sentence_signature(s):
    return re.sub(r"\W+", "", s.lower())[:50]


def summarize_bullets_multi(doc_texts, max_bullets=6, max_chars=220, max_per_doc=2):
    """Version multi-documento: descarta cualquier frase que aparezca
    (casi) igual en MAS DE UN documento de la carpeta — eso es la señal de
    que es membrete/pie de pagina repetido, no contenido real del caso.
    Solo lo que es propio de un unico documento se considera 'punto clave'."""
    doc_sentences = [_candidate_sentences(clean_text(t)) for t in doc_texts]

    sig_counts = Counter()
    for sents in doc_sentences:
        for k in {_sentence_signature(s) for s in sents}:
            sig_counts[k] += 1

    candidates = []
    for doc_idx, sents in enumerate(doc_sentences):
        for s in sents:
            if sig_counts[_sentence_signature(s)] > 1:
                continue  # se repite en mas de un documento -> ruido/membrete
            candidates.append((doc_idx, s))

    if not candidates:
        return [REVISAR]

    all_text = " ".join(s for _, s in candidates)
    words = re.findall(r"[a-záéíóúñA-ZÁÉÍÓÚÑ]{3,}", all_text.lower())
    freq = Counter(w for w in words if w not in STOPWORDS_ES)
    max_freq = max(freq.values()) if freq else 1
    for w in freq:
        freq[w] /= max_freq

    scored = []
    for i, (doc_idx, s) in enumerate(candidates):
        s_words = re.findall(r"[a-záéíóúñA-ZÁÉÍÓÚÑ]{3,}", s.lower())
        score = sum(freq.get(w, 0) for w in s_words) / max(len(s_words), 1)
        scored.append((score, doc_idx, i, s))

    scored.sort(key=lambda x: x[0], reverse=True)
    picked = []
    per_doc = Counter()
    for score, doc_idx, i, s in scored:
        if per_doc[doc_idx] >= max_per_doc:
            continue
        picked.append((doc_idx, i, s))
        per_doc[doc_idx] += 1
        if len(picked) >= max_bullets:
            break
    picked.sort(key=lambda x: (x[0], x[1]))

    bullets = []
    for _, _, s in picked:
        s2 = s[0].upper() + s[1:] if s else s
        if len(s2) > max_chars:
            s2 = s2[:max_chars].rsplit(" ", 1)[0] + "…"
        if not s2.endswith((".", "…")):
            s2 += "."
        bullets.append(s2)
    return bullets


# ---------------------------------------------------------------------------
# ESQUEMA (estructura detectada: secciones tipicas de escritos juridicos)
# ---------------------------------------------------------------------------
ESQUEMA_HEADERS = [
    "HECHOS", "FUNDAMENTOS DE DERECHO", "FUNDAMENTOS JURIDICOS",
    "SUPLICO", "PETICIONES", "ANTECEDENTES DE HECHO", "FALLO",
    "RAZONAMIENTOS JURIDICOS", "PRUEBA", "OTROSI", "CONSIDERANDOS",
]

def build_esquema(text):
    """Detecta encabezados tipicos y arma un esquema jerarquico simple (texto)."""
    lines = text.split("\n")
    esquema = []
    for header in ESQUEMA_HEADERS:
        pat = re.compile(rf"^\s*{header}\b", re.IGNORECASE)
        for line in lines:
            if pat.match(line.strip()):
                esquema.append(header.title())
                break
    numerales = re.findall(r"\b(PRIMERO|SEGUNDO|TERCERO|CUARTO|QUINTO|SEXTO|SÉPTIMO|OCTAVO)\.?\s*[-–.]", text, re.IGNORECASE)
    if numerales:
        esquema.append(f"Clausulas numeradas detectadas: {', '.join(dict.fromkeys(n.title() for n in numerales))}")
    if not esquema:
        return REVISAR
    return "\n".join(f"- {e}" for e in esquema)


def extract_esquema_steps(text):
    """Version 'rica' del esquema para el diagrama visual: una lista ordenada
    de pasos, cada uno con un fragmento real de contenido (no solo el nombre
    de la seccion), en el orden en que aparecen en el documento."""
    text = clean_text(text)
    matches = []

    for header in ESQUEMA_HEADERS:
        m = re.search(rf"\b{header}\b", text, re.IGNORECASE)
        if m:
            snippet = text[m.end():m.end() + 140].strip(" .:;-")
            snippet = re.split(r"(?<=[\.\!\?])\s", snippet)[0]
            label = header.title()
            content = f"{label}: {snippet}" if snippet else label
            matches.append((m.start(), content[:160]))

    for m in re.finditer(
        r"\b(PRIMERO|SEGUNDO|TERCERO|CUARTO|QUINTO|SEXTO|SÉPTIMO|OCTAVO)\.?\s*[-–.]\s*([^\.\n]{10,140})",
        text, re.IGNORECASE
    ):
        label = m.group(1).title()
        snippet = m.group(2).strip()
        matches.append((m.start(), f"{label}: {snippet}"))

    if not matches:
        return None
    matches.sort(key=lambda x: x[0])
    steps, seen = [], set()
    for _, content in matches:
        key = content.lower()[:40]
        if key not in seen:
            seen.add(key)
            steps.append(content)
    return steps[:10]


def extract_esquema_steps_multi(doc_texts):
    """En vez de mezclar las clausulas de TODOS los documentos de la
    carpeta (lo que produce un diagrama confuso), elige el documento que
    parece ser el escrito principal (el que tiene mas fases/clausulas
    numeradas detectadas) y construye el diagrama solo a partir de ese."""
    best_steps, best_score = None, -1
    for text in doc_texts:
        steps = extract_esquema_steps(text)
        score = len(steps) if steps else 0
        if score > best_score:
            best_score = score
            best_steps = steps
    return best_steps


# ---------------------------------------------------------------------------
# RESUMEN CORTO (una frase tipo titular, para la columna "Resumen" del Excel)
# ---------------------------------------------------------------------------
def _is_mostly_uppercase(s):
    letters = [c for c in s if c.isalpha()]
    if not letters:
        return False
    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    return upper_ratio > 0.6


def resumen_corto_from_bullets(bullets, procedimientos=None, max_chars=140):
    """Sin IA no se puede redactar un titular nuevo con garantías, así que
    se aproxima recortando la primera viñeta de prosa normal (no un
    fragmento de cabecera/formulario en mayúsculas, típico de portadas de
    informes) a una longitud de titular. Con IA activada, este valor se
    sustituye por el 'resumen_corto' que redacta el modelo (ver
    ai_extractor.py), mucho más fiable."""
    if not bullets or bullets == [REVISAR]:
        return REVISAR
    candidate = next((b for b in bullets if not _is_mostly_uppercase(b)), bullets[0])
    candidate = candidate.strip()
    if len(candidate) > max_chars:
        candidate = candidate[:max_chars].rsplit(" ", 1)[0] + "…"
    return candidate


# ---------------------------------------------------------------------------
# ASUNTO (nombre del asunto = nombre de la carpeta, normalizado)
# ---------------------------------------------------------------------------
def asunto_from_folder(folder_name):
    name = folder_name.replace("_", " ").replace("-", " ")
    name = " ".join(name.split())
    return name if name else REVISAR
