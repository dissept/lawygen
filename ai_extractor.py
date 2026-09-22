"""
ai_extractor.py
Enriquecimiento OPCIONAL mediante la API de Anthropic (Claude).

Este modulo NUNCA se activa por si solo: solo se usa si el usuario ha
introducido su propia clave de API en la pantalla de "Ajustes de IA" de la
app (guardada en local con config_store.py) y ha dejado la casilla "Usar IA"
activada. Si no hay clave configurada, LawyGen sigue funcionando 100% con
las reglas de texto de legal_parser.py (gratis y offline).

Cuando la IA esta activada, se envia el texto de los documentos de cada
carpeta a la API de Anthropic (api.anthropic.com) para obtener una
extraccion mas precisa de: partes con sus representantes (abogado/
procurador), materia, argumentos de cada parte, objeto/causa juzgada,
plazos, cuantia y costas, y un resumen ejecutivo en viñetas.

Cualquier fallo (sin conexion, clave invalida, limite de tasa, JSON mal
formado, etc.) se captura y se devuelve None: el llamador (processor.py)
debe entonces recurrir a las reglas de texto como respaldo.
"""
import json
import re
import unicodedata
import urllib.request
import urllib.error

import excel_builder
import legal_parser as lp

API_URL = "https://api.anthropic.com/v1/messages"
ANTHROPIC_VERSION = "2023-06-01"

MAX_CHARS_PER_DOC = 30000     # recorte por documento (prioriza calidad: menos recorte)
MAX_TOTAL_CHARS = 120000      # recorte total del conjunto de documentos de la carpeta
REQUEST_TIMEOUT = 180         # segundos (mas margen: respuestas mas completas tardan mas)


MATERIAS_VALIDAS = [m for m, _ in lp.MATERIA_KEYWORDS]


def _normalize(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.strip().lower()


def _match_whitelist(value, whitelist):
    if not isinstance(value, str) or not value.strip():
        return None
    norm = _normalize(value)
    for option in whitelist:
        if _normalize(option) == norm:
            return option
    return None

SYSTEM_PROMPT = (
    "Eres un asistente jurídico que extrae datos estructurados de expedientes "
    "legales españoles (demandas, sentencias, diligencias, notificaciones, etc.). "
    "Debes responder ÚNICAMENTE con un objeto JSON válido, sin texto adicional, "
    "sin explicaciones, sin markdown ni bloques de código, empezando directamente "
    "por '{' y terminando en '}'. "
    "Sé exhaustivo: relee el texto con atención antes de rendirte con un campo. "
    "La prosa jurídica española rara vez usa etiquetas tipo 'Demandante:'; los "
    "datos suelen venir en frases narrativas (p.ej. 'el Procurador D. X en "
    "representación de D. Y contra D. Z, representado por la Procuradora Dña. "
    "W'), en cabeceras de sentencia, o repartidos entre varios documentos de la "
    "misma carpeta — combina la información de TODOS los documentos antes de "
    "decidir que un dato falta. "
    "Dicho esto, usa \"revisar\" (o [] para listas) cuando el dato genuinamente "
    "no aparece en el texto proporcionado — nunca inventes nombres, cifras, "
    "fechas ni cláusulas que no estén respaldadas por el contenido. "
    "Escapa correctamente cualquier salto de línea, comilla o carácter especial "
    "dentro de los valores de texto (usa \\n, \\\", etc.) para que el JSON sea "
    "válido; no incluyas saltos de línea sin escapar dentro de una cadena. "
    "El JSON debe tener EXACTAMENTE estas claves:\n"
    "{\n"
    '  "resumen_corto": string,  // UNA frase (máx. 25 palabras) que resuma el asunto, estilo "Reclamación de deuda contra X" o "Despido improcedente de Y"\n'
    '  "juzgado": string,\n'
    '  "materia": string,  // uno de: Civil, Penal, Laboral, Mercantil, Contencioso-Administrativo, u "revisar"\n'
    '  "partes": string,  // formato "Demandante: Nombre | Demandado: Nombre"\n'
    '  "representantes": string,  // formato "Abogado: Nombre | Procurador: Nombre" (de cada parte si consta)\n'
    '  "procedimientos": [ {"numero": string, "tipo": string, "grupo_id": string} ],\n'
    '    // Un objeto por cada número de procedimiento/autos/expediente real encontrado (ignora citas de jurisprudencia).\n'
    '    // "numero" ej. "936/2024". "tipo" ej. "Procedimiento Ordinario", "Pieza de Medidas Cautelares", "Recurso de Apelación", "Ejecución".\n'
    '    // "grupo_id" es CLAVE: usa el MISMO grupo_id (p.ej. el número del procedimiento principal, como "936/2024") para\n'
    '    // todos los procedimientos que sean la MISMA disputa vista desde distintas piezas — el procedimiento principal,\n'
    '    // su pieza de medidas cautelares, su recurso de apelación o su ejecución de sentencia son el MISMO caso aunque\n'
    '    // tengan números de expediente distintos (esto se ve porque un escrito de apelación o de medidas cautelares cita\n'
    '    // expresamente el número de autos del procedimiento principal del que dimana). Si un número de procedimiento es\n'
    '    // una disputa independiente y no está relacionado con ningún otro de esta carpeta, usa su propio número como grupo_id.\n'
    '  "objeto": string,  // breve descripción (1-2 frases) de qué se pide/juzga\n'
    '  "fundamentos_derecho": [string],  // 2-5 viñetas con la base jurídica/procesal alegada (Fundamentos de Derecho, Alegaciones, Excepción Procesal, Otrosí Digo...)\n'
    '  "cuantia_costas": string,  // formato "Cuantía: X € | Costas: ..." \n'
    '  "estado_sugerido": string,  // uno de: Abierto, En trámite, En apelación, Suspendido, Archivado, Resuelto\n'
    '  "resumen_ejecutivo": [string],  // 4-8 viñetas claras y concisas, tono ejecutivo, con los datos clave del caso\n'
    '  "argumentos": [ {"parte": string, "puntos": [string]} ],  // 2-5 puntos por parte, resumen de sus argumentos/pretensiones\n'
    '  "documentos_por_parte": {"demandado": [string], "demandante": [string]}  // SOLO si los documentos vienen etiquetados con "(Carpeta: Demandado)" / "(Carpeta: Demandante)": 2-4 viñetas resumiendo el contenido de esos documentos por cada parte. Si no hay etiquetas de carpeta, deja ambas listas vacías []\n'
    "}"
)


def _truncate_docs(doc_texts, doc_labels=None):
    parts = []
    total = 0
    for i, t in enumerate(doc_texts, start=1):
        t = t[:MAX_CHARS_PER_DOC]
        if total + len(t) > MAX_TOTAL_CHARS:
            t = t[: max(0, MAX_TOTAL_CHARS - total)]
        if not t:
            break
        label = doc_labels[i - 1] if doc_labels and i - 1 < len(doc_labels) else None
        tag = f" (Carpeta: {label})" if label else ""
        parts.append(f"--- DOCUMENTO {i}{tag} ---\n{t}")
        total += len(t)
        if total >= MAX_TOTAL_CHARS:
            break
    return "\n\n".join(parts)


def _extract_json(raw_text):
    """El modelo deberia devolver JSON puro, pero por robustez:
    1) probamos parseo estricto,
    2) probamos con strict=False (tolera saltos de línea/control chars
       sueltos dentro de las cadenas, un fallo muy comun al generar JSON
       con un LLM que no rompe realmente la estructura),
    3) si viniera envuelto en texto o markdown, extraemos el primer
       bloque {...} y repetimos los dos intentos anteriores sobre él."""
    raw_text = raw_text.strip()
    for candidate in _json_candidates(raw_text):
        for strict in (True, False):
            try:
                return json.loads(candidate, strict=strict)
            except Exception:
                continue
    return None


def _json_candidates(raw_text):
    candidates = [raw_text]
    m = re.search(r"\{.*\}", raw_text, re.DOTALL)
    if m and m.group(0) != raw_text:
        candidates.append(m.group(0))
    return candidates


def extract_fields_ai(doc_texts, api_key, model="claude-sonnet-5", log=None, doc_labels=None):
    """Llama a la API de Anthropic para extraer los campos enriquecidos de
    una carpeta de caso. Devuelve un dict (ver SYSTEM_PROMPT) o None si
    falla por cualquier motivo (se debe hacer fallback a legal_parser).
    doc_labels (opcional): lista paralela a doc_texts con el nombre de la
    subcarpeta de parte (p.ej. "Demandado"/"Demandante") de cada documento,
    si existe — permite rellenar 'documentos_por_parte'."""
    if not api_key or not doc_texts:
        return None

    combined = _truncate_docs(doc_texts, doc_labels=doc_labels)
    if not combined.strip():
        return None

    body = {
        "model": model,
        "max_tokens": 16000,
        # La prioridad ahora es la máxima calidad posible (evitar cualquier
        # revisión manual), no el coste ni la velocidad. Se deja el
        # "adaptive thinking" de Sonnet 5 activado (comportamiento por
        # defecto) para que el modelo pueda razonar sobre expedientes con
        # prosa ambigua o repartida entre muchos documentos antes de
        # responder. Con max_tokens alto y el reintento automático por
        # truncamiento (más abajo) como red de seguridad, esto ya no corta
        # la respuesta a mitad como pasaba con un presupuesto más bajo.
        "system": SYSTEM_PROMPT,
        "messages": [
            {"role": "user", "content": f"Extrae los datos del siguiente expediente:\n\n{combined}"}
        ],
    }

    data, err = _call_messages_api(body, api_key, log=log)
    if data is None:
        return None

    # Si la respuesta se cortó por falta de espacio (carpetas con muchos
    # documentos/procedimientos generan un JSON más largo), reintenta una
    # vez con bastante más presupuesto antes de rendirse.
    if data.get("stop_reason") == "max_tokens":
        if log:
            log("  [AVISO] IA: respuesta cortada por límite de tokens, reintentando con más margen...")
        body["max_tokens"] = 32000
        retry_data, retry_err = _call_messages_api(body, api_key, log=log)
        if retry_data is not None:
            data = retry_data

    try:
        text_blocks = [b.get("text", "") for b in data.get("content", []) if b.get("type") == "text"]
        raw_text = "\n".join(text_blocks)
    except Exception:
        raw_text = ""

    if data.get("stop_reason") == "max_tokens" and log:
        log("  [AVISO] IA: la respuesta sigue cortándose por límite de tokens; puede quedar incompleta.")

    parsed = _extract_json(raw_text)
    if not parsed:
        if log:
            log("  [AVISO] IA: la respuesta no fue JSON válido. Se usan reglas de texto.")
        return None

    return parsed


def _call_messages_api(body, api_key, log=None):
    """Devuelve (data, error_code). error_code es None si fue bien, o
    'thinking_not_supported' si el modelo rechazo thinking:disabled, o
    'error' para cualquier otro fallo (ya registrado en el log)."""
    req = urllib.request.Request(
        API_URL,
        data=json.dumps(body).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": ANTHROPIC_VERSION,
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as resp:
            return json.loads(resp.read().decode("utf-8")), None
    except urllib.error.HTTPError as e:
        try:
            err_body = e.read().decode("utf-8")
        except Exception:
            err_body = str(e)
        if log:
            log(f"  [AVISO] IA: error HTTP {e.code} de la API de Anthropic: {err_body[:200]}")
        return None, "error"
    except Exception as e:
        if log:
            log(f"  [AVISO] IA: no se pudo contactar con la API de Anthropic ({e}). Se usan reglas de texto.")
        return None, "error"


def merge_ai_into_analysis(a, ai_data, REVISAR="revisar"):
    """Combina el resultado de la IA con el analisis por reglas (a). La IA
    solo sobreescribe un campo si aporta un valor no vacio / distinto de
    'revisar'; si la IA no supo un dato, se conserva lo que ya habia
    calculado legal_parser.py (o 'revisar' si tampoco eso se encontro)."""
    if not ai_data:
        return a

    def pick(ai_val, current):
        if isinstance(ai_val, str) and ai_val.strip() and ai_val.strip().lower() != REVISAR:
            return ai_val.strip()
        return current

    a["juzgado"] = pick(ai_data.get("juzgado"), a.get("juzgado", REVISAR))
    a["partes"] = pick(ai_data.get("partes"), a.get("partes", REVISAR))
    a["materia"] = _match_whitelist(ai_data.get("materia"), MATERIAS_VALIDAS) or a.get("materia", REVISAR)
    a["representantes"] = pick(ai_data.get("representantes"), a.get("representantes", REVISAR))
    a["objeto"] = pick(ai_data.get("objeto"), a.get("objeto", REVISAR))
    a["cuantia_costas"] = pick(ai_data.get("cuantia_costas"), a.get("cuantia_costas", REVISAR))
    a["resumen_corto"] = pick(ai_data.get("resumen_corto"), a.get("resumen_corto", REVISAR))

    ai_estado = _match_whitelist(ai_data.get("estado_sugerido"), excel_builder.ESTADOS)
    if ai_estado:
        a["estado"] = ai_estado

    ai_procs = ai_data.get("procedimientos")
    if isinstance(ai_procs, list) and ai_procs:
        flat = []
        groups = {}   # grupo_id -> list[label] en orden de aparicion
        group_order = []
        for item in ai_procs:
            if not isinstance(item, dict):
                continue
            numero = str(item.get("numero", "")).strip()
            tipo = str(item.get("tipo", "")).strip()
            if not numero:
                continue
            label = f"{tipo} {numero}".strip() if tipo else numero
            flat.append(label)
            grupo_id = str(item.get("grupo_id", "")).strip() or numero
            if grupo_id not in groups:
                groups[grupo_id] = []
                group_order.append(grupo_id)
            groups[grupo_id].append(label)
        if flat:
            a["procedimientos"] = flat
            a["procedimientos_grouped"] = [groups[g] for g in group_order]

    ai_resumen = ai_data.get("resumen_ejecutivo")
    if isinstance(ai_resumen, list) and ai_resumen:
        cleaned = [b.strip() for b in ai_resumen if isinstance(b, str) and b.strip()]
        if cleaned:
            a["resumen_bullets"] = cleaned

    ai_fundamentos = ai_data.get("fundamentos_derecho")
    if isinstance(ai_fundamentos, list) and ai_fundamentos:
        cleaned = [b.strip() for b in ai_fundamentos if isinstance(b, str) and b.strip()]
        if cleaned:
            a["fundamentos_derecho"] = cleaned

    ai_argumentos = ai_data.get("argumentos")
    if isinstance(ai_argumentos, list) and ai_argumentos:
        bullets = []
        for item in ai_argumentos:
            if not isinstance(item, dict):
                continue
            parte = str(item.get("parte", "")).strip()
            puntos = item.get("puntos") or []
            if not parte or not isinstance(puntos, list):
                continue
            for punto in puntos:
                if isinstance(punto, str) and punto.strip():
                    bullets.append(f"{parte}: {punto.strip()}")
        if bullets:
            a["argumentos_bullets"] = bullets

    ai_por_parte = ai_data.get("documentos_por_parte")
    if isinstance(ai_por_parte, dict):
        for role_key, out_key in (("demandado", "documentos_demandado_bullets"),
                                   ("demandante", "documentos_demandante_bullets")):
            items = ai_por_parte.get(role_key)
            if isinstance(items, list) and items:
                cleaned = [b.strip() for b in items if isinstance(b, str) and b.strip()]
                if cleaned:
                    a[out_key] = cleaned

    a["ai_used"] = True
    return a
