"""
processor.py
Recorre las subcarpetas de un directorio raiz (cada subcarpeta = un asunto),
extrae texto de los documentos y aplica las reglas de legal_parser.py.

Tres modos independientes (llamados desde los tres botones de la app):
  - MODE_EXCEL:   solo calcula las filas para el Excel maestro (no escribe docx)
  - MODE_RESUMEN: genera UN SOLO Word "Resumen_General.docx" en la carpeta
                   raiz, con una seccion organizada por cada asunto
  - MODE_ESQUEMA: genera un Word "<Carpeta>_Esquema.docx" por asunto, con un
                   diagrama de flujo visual (imagen) de las fases detectadas
"""
import os
import tempfile
import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

import extractor
import legal_parser as lp
import diagram_builder
import config_store
import ai_extractor

MODE_EXCEL = "excel"
MODE_RESUMEN = "resumen"
MODE_ESQUEMA = "esquema"

NOTE_TEXT_REGLAS = (
    "Nota: este documento fue generado automáticamente mediante reglas de texto "
    "(sin inteligencia artificial de pago). Los campos marcados como \"revisar\" "
    "no pudieron determinarse con certeza a partir del contenido disponible y "
    "deben verificarse manualmente."
)

NOTE_TEXT_IA = (
    "Nota: este documento fue generado automáticamente combinando reglas de "
    "texto y un análisis asistido por IA (Claude, Anthropic) sobre el "
    "contenido de los documentos de cada carpeta. Los campos marcados como "
    "\"revisar\" no pudieron determinarse con certeza y deben verificarse "
    "manualmente."
)


def _analyze_folder(folder_path, log=print):
    """Extrae todo lo necesario de una carpeta. Devuelve un dict con los
    campos calculados (juzgado, partes, procedimientos, resumen, esquema...).
    Procesa cada documento POR SEPARADO y luego combina de forma
    inteligente — evita mezclar todo en un unico bloque de texto, que es
    lo que producia procedimientos duplicados/garbled y puntos clave
    repetidos con membretes."""
    folder_name = os.path.basename(folder_path.rstrip(os.sep))
    asunto = lp.asunto_from_folder(folder_name)

    doc_paths = extractor.list_documents(folder_path)
    if not doc_paths:
        log(f"  [AVISO] Sin documentos soportados en '{folder_name}'.")
        return {
            "folder_name": folder_name, "asunto": asunto, "doc_paths": [],
            "juzgado": lp.REVISAR, "partes": lp.REVISAR, "representantes": lp.REVISAR,
            "materia": lp.REVISAR, "objeto": lp.REVISAR, "fundamentos_derecho": [lp.REVISAR],
            "cuantia_costas": lp.REVISAR,
            "procedimientos": [lp.REVISAR], "estado": "Abierto",
            "resumen_bullets": [lp.REVISAR], "argumentos_bullets": [lp.REVISAR],
            "resumen_corto": lp.REVISAR, "esquema": lp.REVISAR,
            "esquema_steps": None, "main_pdf": None, "ai_used": False, "doc_labels": [],
        }

    doc_texts = []  # texto limpio, uno por documento (no concatenado)
    doc_labels = []  # rol de parte ("Demandado"/"Demandante") o None, paralelo a doc_texts
    party_texts = {"Demandado": [], "Demandante": []}  # para el fallback sin IA
    for p in doc_paths:
        text, err = extractor.extract_text_from_file(p)
        if err:
            log(f"  [AVISO] {err}")
        if text and text.strip():
            cleaned = lp.clean_text(text)
            doc_texts.append(cleaned)
            role = extractor.party_role_for_path(p, folder_path)
            doc_labels.append(role)
            if role:
                party_texts[role].append(cleaned)

    if not doc_texts:
        log(f"  [AVISO] No se pudo extraer texto legible de '{folder_name}' (¿PDFs escaneados sin OCR?).")

    has_text = bool(doc_texts)
    juzgado = lp.extract_juzgado_multi(doc_texts) if has_text else lp.REVISAR
    partes = lp.extract_partes_multi(doc_texts) if has_text else lp.REVISAR
    representantes = lp.extract_representantes_multi(doc_texts) if has_text else lp.REVISAR
    materia = lp.extract_materia_multi(doc_texts) if has_text else lp.REVISAR
    objeto = lp.extract_objeto_multi(doc_texts) if has_text else lp.REVISAR
    fundamentos_derecho = lp.extract_fundamentos_derecho_multi(doc_texts) if has_text else [lp.REVISAR]
    cuantia_costas = lp.extract_cuantia_costas_multi(doc_texts) if has_text else lp.REVISAR
    procedimientos = lp.extract_procedimientos_multi(doc_texts) if has_text else [lp.REVISAR]
    estado = lp.suggest_estado_multi(doc_texts) if has_text else "Abierto"
    resumen_bullets = lp.summarize_bullets_multi(doc_texts) if has_text else [lp.REVISAR]
    resumen_corto = lp.resumen_corto_from_bullets(resumen_bullets) if has_text else lp.REVISAR
    argumentos_bullets = [lp.REVISAR]
    esquema_steps = lp.extract_esquema_steps_multi(doc_texts) if has_text else None
    esquema = ("\n".join(f"- {s.split(':')[0]}" for s in esquema_steps[:6]) if esquema_steps else lp.REVISAR)

    # Documentos por parte (si el despacho organiza subcarpetas "Demandado" /
    # "Demandante"): sin IA, un resumen extractivo de esos documentos; con
    # IA, se sustituye por algo mejor redactado en merge_ai_into_analysis.
    documentos_demandado_bullets = (
        lp.summarize_bullets_multi(party_texts["Demandado"], max_bullets=4)
        if party_texts["Demandado"] else None
    )
    documentos_demandante_bullets = (
        lp.summarize_bullets_multi(party_texts["Demandante"], max_bullets=4)
        if party_texts["Demandante"] else None
    )

    main_pdf = extractor.pick_main_pdf(folder_path)

    a = {
        "folder_name": folder_name, "asunto": asunto, "doc_paths": doc_paths,
        "juzgado": juzgado, "partes": partes, "representantes": representantes,
        "materia": materia, "objeto": objeto, "fundamentos_derecho": fundamentos_derecho,
        "cuantia_costas": cuantia_costas,
        "procedimientos": procedimientos,
        "estado": estado, "resumen_bullets": resumen_bullets, "resumen_corto": resumen_corto,
        "argumentos_bullets": argumentos_bullets, "esquema": esquema,
        "esquema_steps": esquema_steps, "main_pdf": main_pdf, "ai_used": False,
        "doc_labels": doc_labels,
    }
    if documentos_demandado_bullets:
        a["documentos_demandado_bullets"] = documentos_demandado_bullets
    if documentos_demandante_bullets:
        a["documentos_demandante_bullets"] = documentos_demandante_bullets

    # ------------------------------------------------------------------
    # Enriquecimiento OPCIONAL con IA (Claude/Anthropic). Solo se activa
    # si el usuario ha guardado su propia clave de API en Ajustes de IA.
    # Cualquier fallo cae de vuelta a los valores calculados por reglas.
    # ------------------------------------------------------------------
    if has_text and config_store.is_ai_enabled():
        api_key = config_store.get_api_key()
        model = config_store.get_model()
        try:
            ai_data = ai_extractor.extract_fields_ai(doc_texts, api_key, model=model, log=log, doc_labels=doc_labels)
            if ai_data:
                a = ai_extractor.merge_ai_into_analysis(a, ai_data, REVISAR=lp.REVISAR)
                log(f"  IA: análisis completado para '{folder_name}'.")
        except Exception as e:
            log(f"  [AVISO] IA: fallo inesperado ({e}). Se mantienen las reglas de texto.")

    return a


def _rows_from_analysis(a):
    doc_info = (os.path.basename(a["main_pdf"]), os.path.abspath(a["main_pdf"])) if a["main_pdf"] else None

    # Lista de los demas documentos de la carpeta (adjuntos aparte del
    # principal), para responder a "cuantos adjuntos hay" sin perder el
    # hipervinculo al documento principal.
    other_docs = [p for p in a.get("doc_paths", []) if p != a.get("main_pdf")]
    adjuntos = "; ".join(os.path.basename(p) for p in other_docs) if other_docs else "—"

    argumentos_txt = (
        " | ".join(a.get("argumentos_bullets", [lp.REVISAR]))
        if a.get("argumentos_bullets") not in (None, [], [lp.REVISAR])
        else lp.REVISAR
    )

    fundamentos_txt = (
        " | ".join(a.get("fundamentos_derecho", [lp.REVISAR]))
        if a.get("fundamentos_derecho") not in (None, [], [lp.REVISAR])
        else lp.REVISAR
    )

    demandado_bullets = a.get("documentos_demandado_bullets")
    demandado_txt = (
        " | ".join(demandado_bullets)
        if demandado_bullets not in (None, [], [lp.REVISAR])
        else "—"
    )
    demandante_bullets = a.get("documentos_demandante_bullets")
    demandante_txt = (
        " | ".join(demandante_bullets)
        if demandante_bullets not in (None, [], [lp.REVISAR])
        else "—"
    )

    # Si la IA agrupó los procedimientos relacionados (principal + su pieza
    # de medidas cautelares + su apelación = mismo caso), generamos UNA fila
    # por grupo en vez de una por cada número suelto. Sin esa agrupación
    # (modo sin IA), mantenemos el comportamiento anterior: una fila por
    # número de procedimiento encontrado, ya que las reglas de texto no
    # pueden determinar con fiabilidad qué números pertenecen al mismo caso.
    grouped = a.get("procedimientos_grouped")
    proc_rows = ["; ".join(g) for g in grouped] if grouped else a["procedimientos"]

    rows = []
    for proc in proc_rows:
        rows.append({
            "Asunto": a["asunto"], "Resumen": a.get("resumen_corto", lp.REVISAR),
            "Procedimiento": proc,
            "Materia": a.get("materia", lp.REVISAR),
            "Juzgado": a["juzgado"], "Partes": a["partes"],
            "Representantes": a.get("representantes", lp.REVISAR),
            "Argumentos": argumentos_txt,
            "Objeto": a.get("objeto", lp.REVISAR),
            "FundamentosDerecho": fundamentos_txt,
            "CuantiaCostas": a.get("cuantia_costas", lp.REVISAR),
            "Estado": a["estado"], "Carpeta": a["folder_name"],
            "Documento": doc_info,
            "Adjuntos": adjuntos,
            "DocumentosDemandado": demandado_txt,
            "DocumentosDemandante": demandante_txt,
        })
    return rows


def _add_note(d, ai_used=False):
    note = d.add_paragraph()
    run = note.add_run(NOTE_TEXT_IA if ai_used else NOTE_TEXT_REGLAS)
    run.italic = True
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)


def _add_datos_table(d, a):
    other_docs = [p for p in a.get("doc_paths", []) if p != a.get("main_pdf")]
    adjuntos = ("; ".join(os.path.basename(p) for p in other_docs)
                if other_docs else "Sin adjuntos adicionales")
    table = d.add_table(rows=0, cols=2)
    table.style = "Light Grid Accent 1"
    for label, value in [
        ("Carpeta", a["folder_name"]),
        ("Resumen", a.get("resumen_corto", lp.REVISAR)),
        ("Materia", a.get("materia", lp.REVISAR)),
        ("Juzgado", a["juzgado"]),
        ("Partes", a["partes"]),
        ("Representantes (abogado/procurador)", a.get("representantes", lp.REVISAR)),
        ("Procedimiento(s)", "; ".join(a["procedimientos"])),
        ("Objeto / causa juzgada", a.get("objeto", lp.REVISAR)),
        ("Fundamentos de Derecho", "; ".join(a.get("fundamentos_derecho", [lp.REVISAR])) or lp.REVISAR),
        ("Cuantía y costas", a.get("cuantia_costas", lp.REVISAR)),
        ("Estado sugerido", a["estado"]),
        ("Documento principal", os.path.basename(a["main_pdf"]) if a.get("main_pdf") else lp.REVISAR),
        ("Otros documentos/adjuntos", adjuntos),
    ] + ([("Documentos del demandado", "; ".join(a["documentos_demandado_bullets"]))]
         if a.get("documentos_demandado_bullets") else []) + (
        [("Documentos del demandante", "; ".join(a["documentos_demandante_bullets"]))]
        if a.get("documentos_demandante_bullets") else []
    ):
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value or lp.REVISAR
    return table


# ---------------------------------------------------------------------------
# RESUMEN — un unico documento consolidado para toda la carpeta raiz
# ---------------------------------------------------------------------------
def _write_master_resumen_docx(root_path, analyses):
    root_name = os.path.basename(root_path.rstrip(os.sep)) or "Casos"
    d = Document()

    d.add_heading("Resumen General de Procedimientos Legales", level=0)
    meta = d.add_paragraph()
    meta.add_run(f"Carpeta analizada: {root_name}\n").italic = True
    meta.add_run(f"Generado: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}\n").italic = True
    meta.add_run(f"Asuntos incluidos: {len(analyses)}").italic = True

    # Indice rapido
    d.add_heading("Índice de asuntos", level=1)
    for a in analyses:
        p = d.add_paragraph(style="List Number")
        p.add_run(a["asunto"]).bold = True
        p.add_run(f"  ({a['folder_name']})").italic = True

    d.add_page_break()

    any_ai_used = False
    for a in analyses:
        if a.get("ai_used"):
            any_ai_used = True
        d.add_heading(a["asunto"], level=1)
        d.add_heading("Datos identificados", level=2)
        _add_datos_table(d, a)

        d.add_heading("Puntos clave (resumen ejecutivo)", level=2)
        bullets = a["resumen_bullets"] or [lp.REVISAR]
        if bullets == [lp.REVISAR]:
            d.add_paragraph(lp.REVISAR)
        else:
            for b in bullets:
                d.add_paragraph(b, style="List Bullet")

        argumentos = a.get("argumentos_bullets") or [lp.REVISAR]
        if argumentos != [lp.REVISAR]:
            d.add_heading("Argumentos de las partes", level=2)
            for arg in argumentos:
                d.add_paragraph(arg, style="List Bullet")

        d.add_paragraph()  # espaciado antes del siguiente asunto

    _add_note(d, ai_used=any_ai_used)

    out_path = os.path.join(root_path, "Resumen_General.docx")
    d.save(out_path)
    return out_path


# ---------------------------------------------------------------------------
# ESQUEMA — diagrama de flujo visual por asunto
# ---------------------------------------------------------------------------
def _write_esquema_docx(folder_path, a):
    d = Document()
    d.add_heading(f"Esquema — {a['asunto']}", level=1)
    meta = d.add_paragraph()
    meta.add_run(f"Carpeta: {a['folder_name']}\n").italic = True
    meta.add_run(f"Generado: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}").italic = True

    d.add_heading("Diagrama del procedimiento", level=2)

    steps = a.get("esquema_steps")
    if steps:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            diagram_builder.build_flowchart(steps, tmp_path, title=a["asunto"][:60])
            pic_p = d.add_paragraph()
            pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pic_p.add_run().add_picture(tmp_path, width=Inches(6.2))
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
    else:
        d.add_paragraph(
            "No se detectaron suficientes fases o cláusulas numeradas en el "
            "documento para construir un diagrama. Revisar manualmente."
        )

    if a["esquema"] and a["esquema"] != lp.REVISAR:
        d.add_heading("Detalle textual", level=2)
        for line in a["esquema"].split("\n"):
            d.add_paragraph(line.lstrip("- ").strip(), style="List Bullet")

    _add_note(d, ai_used=a.get("ai_used", False))

    out_path = os.path.join(folder_path, f"{a['folder_name']}_Esquema.docx")
    d.save(out_path)
    return out_path


def process_root_folder(root_path, mode, progress_callback=None, log=print):
    """Recorre todas las subcarpetas directas de root_path segun el modo:
      MODE_EXCEL   -> devuelve (rows, [])
      MODE_RESUMEN -> devuelve (None, [ruta del Resumen_General.docx])
      MODE_ESQUEMA -> devuelve (None, [rutas de los .docx generados, uno por asunto])
    """
    subfolders = sorted([
        os.path.join(root_path, name) for name in os.listdir(root_path)
        if os.path.isdir(os.path.join(root_path, name)) and not name.startswith(".")
    ])

    all_rows = []
    generated_files = []
    analyses = []
    total = len(subfolders)
    for i, folder in enumerate(subfolders, start=1):
        name = os.path.basename(folder)
        log(f"[{i}/{total}] Procesando: {name}")
        try:
            a = _analyze_folder(folder, log=log)
            if mode == MODE_EXCEL:
                all_rows.extend(_rows_from_analysis(a))
            elif mode == MODE_RESUMEN:
                analyses.append(a)
            elif mode == MODE_ESQUEMA:
                generated_files.append(_write_esquema_docx(folder, a))
        except Exception as e:
            log(f"  [ERROR] Fallo procesando '{name}': {e}")
            if mode == MODE_EXCEL:
                all_rows.append({
                    "Asunto": lp.asunto_from_folder(name), "Resumen": lp.REVISAR,
                    "Procedimiento": lp.REVISAR,
                    "Materia": lp.REVISAR, "Juzgado": lp.REVISAR, "Partes": lp.REVISAR,
                    "Representantes": lp.REVISAR, "Argumentos": lp.REVISAR,
                    "Objeto": lp.REVISAR, "FundamentosDerecho": lp.REVISAR, "CuantiaCostas": lp.REVISAR,
                    "Estado": "Abierto", "Carpeta": name, "Documento": None,
                    "Adjuntos": "—", "DocumentosDemandado": "—", "DocumentosDemandante": "—",
                })
        if progress_callback:
            progress_callback(i, total)

    if mode == MODE_RESUMEN and analyses:
        log("\nConsolidando Resumen General...")
        path = _write_master_resumen_docx(root_path, analyses)
        generated_files.append(path)

    return all_rows, generated_files
