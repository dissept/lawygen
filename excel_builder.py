"""
excel_builder.py
Genera el Excel de seguimiento con:
  - Una sola hoja
  - Columnas: Asunto, Procedimiento, Juzgado, Partes, Estado, Carpeta, Documento
  - Estado: desplegable real (validacion de datos de Excel)
  - Documento: hipervinculo directo al PDF principal (ruta absoluta local)
"""
import os
import re
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

HEADERS = [
    "Asunto", "Resumen", "Procedimiento", "Materia", "Juzgado", "Partes",
    "Representantes (abogado/procurador)", "Argumentos de las partes",
    "Objeto / causa juzgada", "Fundamentos de Derecho", "Cuantía y costas", "Estado",
    "Carpeta", "Documento principal", "Otros documentos adjuntos",
    "Documentos del demandado", "Documentos del demandante",
]
# Columnas y su key en el dict de fila, en el mismo orden que HEADERS
# (excepto "Documento principal", que se trata aparte por el hipervinculo).
FIELD_KEYS = [
    "Asunto", "Resumen", "Procedimiento", "Materia", "Juzgado", "Partes",
    "Representantes", "Argumentos", "Objeto", "FundamentosDerecho", "CuantiaCostas",
    "Estado",
]
CARPETA_COL = len(FIELD_KEYS) + 1          # 13
DOCUMENTO_COL = CARPETA_COL + 1            # 14
ADJUNTOS_COL = DOCUMENTO_COL + 1           # 15
DEMANDADO_COL = ADJUNTOS_COL + 1           # 16
DEMANDANTE_COL = DEMANDADO_COL + 1         # 17
ESTADO_COL = FIELD_KEYS.index("Estado") + 1  # 12

ESTADOS = ["Abierto", "En trámite", "En apelación", "Suspendido", "Archivado", "Resuelto"]

META_SHEET_NAME = "_LawyGenMeta"

# Columnas "de contenido" que se protegen si el usuario las edito a mano.
# Carpeta y Documento quedan fuera: son datos estructurales, siempre deben
# reflejar la carpeta/archivo real.
MERGE_COLUMNS = FIELD_KEYS + ["Adjuntos", "DocumentosDemandado", "DocumentosDemandante"]

# Igual que HEADERS pero con los nombres internos de cada columna, en el
# mismo orden -- se usa para dejar una marca oculta en cada encabezado (ver
# el bucle de escritura de encabezados y _load_previous).
COLUMN_KEYS = FIELD_KEYS + ["Carpeta", "Documento", "Adjuntos", "DocumentosDemandado", "DocumentosDemandante"]

def _row_key(row):
    proc = row.get("Procedimiento", "") or ""
    m = re.search(r"\d{1,6}\s*[/\-]\s*\d{2,4}", proc)
    proc_key = m.group(0).replace(" ", "") if m else proc
    return f"{row.get('Carpeta', '')}||{proc_key}"


def _load_previous(output_path):
    old_values, old_meta = {}, {}
    if not os.path.exists(output_path):
        return old_values, old_meta
    try:
        wb_old = load_workbook(output_path, data_only=True)
    except Exception:
        return old_values, old_meta

    if "Seguimiento Legal" in wb_old.sheetnames:
        ws_old = wb_old["Seguimiento Legal"]
        # Ya no asumimos que "la columna N siempre es tal campo": el usuario
        # puede renombrar, reordenar o editar las columnas del Excel a su
        # gusto. En vez de eso, leemos la marca oculta ("LAWYGEN_KEY:...")
        # que build_workbook deja como comentario en cada celda de
        # encabezado -- ese comentario viaja con la celda si se mueve o
        # reordena la columna en Excel, y no le afecta que se renombre el
        # texto visible del encabezado.
        col_index = {}
        for cell in ws_old[1]:
            comment = getattr(cell, "comment", None)
            if not comment or not comment.text:
                continue
            first_line = comment.text.split("\n", 1)[0]
            if first_line.startswith("LAWYGEN_KEY:"):
                col_index[first_line[len("LAWYGEN_KEY:"):]] = cell.column

        if not col_index:
            # Archivo de una version anterior a esta marca (o alguien borro
            # las notas a mano): no hay forma fiable de saber que columna es
            # cada cosa. Mejor no arriesgar un merge mal hecho -- se
            # regenera todo desde cero esta vez, como si no hubiera archivo
            # previo.
            return old_values, old_meta

        n_cols_old = ws_old.max_column
        for row_cells in ws_old.iter_rows(min_row=2):
            values = {
                name: (row_cells[idx - 1].value if idx <= n_cols_old else None)
                for name, idx in col_index.items()
            }
            key = _row_key({"Carpeta": values.get("Carpeta", ""), "Procedimiento": values.get("Procedimiento", "")})
            old_values[key] = values

    if META_SHEET_NAME in wb_old.sheetnames:
        ws_meta = wb_old[META_SHEET_NAME]
        meta_header = [c.value for c in ws_meta[1]]
        meta_index = {name: i for i, name in enumerate(meta_header) if name}
        for row_cells in ws_meta.iter_rows(min_row=2):
            values = {name: row_cells[idx].value for name, idx in meta_index.items()}
            key = values.get("_row_key")
            if key:
                old_meta[key] = values

    return old_values, old_meta


def _merge_row(row, old_values, old_meta):
    key = _row_key(row)
    prev_values = old_values.get(key, {})
    prev_meta = old_meta.get(key, {})

    row_final = dict(row)
    row_meta = {}
    for col in MERGE_COLUMNS:
        new_auto = row.get(col, "")
        row_meta[col] = new_auto

        if key in old_values:
            old_cell = prev_values.get(col, "")
            old_auto = prev_meta.get(col, None)
            edited_by_user = (old_auto is not None) and (old_cell != old_auto)
            if edited_by_user:
                row_final[col] = old_cell
    return row_final, row_meta

def build_workbook(rows, output_path):
    """rows: lista de dicts con keys = FIELD_KEYS + 'Carpeta' + 'Documento'
    (tupla (texto_visible, ruta_absoluta) o None) + 'Adjuntos' (texto)."""
    old_values, old_meta = _load_previous(output_path)

    wb = Workbook()
    ws = wb.active
    ws.title = "Seguimiento Legal"

    header_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
        c.comment = Comment(
            f"LAWYGEN_KEY:{COLUMN_KEYS[col - 1]}\n"
            "(No borrar esta nota: permite que LawyGen reconozca esta "
            "columna aunque renombres el título o la muevas de sitio.)",
            "LawyGen",
        )

    total_cols = len(HEADERS)
    meta_rows = []
    for r_idx, row in enumerate(rows, start=2):
        merged_row, row_meta = _merge_row(row, old_values, old_meta)
        meta_rows.append((_row_key(row), row_meta))
        for col, key in enumerate(FIELD_KEYS, start=1):
            value = merged_row.get(key, "")
            ws.cell(row=r_idx, column=col, value=value if key != "Estado" else (value or "Abierto")).border = border

        ws.cell(row=r_idx, column=CARPETA_COL, value=row.get("Carpeta", "")).border = border

        doc_cell = ws.cell(row=r_idx, column=DOCUMENTO_COL)
        doc_info = row.get("Documento")
        if doc_info:
            display_name, abs_path = doc_info
            # file:/// URI para que el hipervinculo abra el archivo local
            uri = "file:///" + abs_path.replace("\\", "/").lstrip("/")
            doc_cell.value = display_name
            doc_cell.hyperlink = uri
            doc_cell.font = Font(color="0563C1", underline="single")
        else:
            doc_cell.value = "revisar"
        doc_cell.border = border

        ws.cell(row=r_idx, column=ADJUNTOS_COL, value=merged_row.get("Adjuntos", "—")).border = border
        ws.cell(row=r_idx, column=DEMANDADO_COL, value=merged_row.get("DocumentosDemandado", "—")).border = border
        ws.cell(row=r_idx, column=DEMANDANTE_COL, value=merged_row.get("DocumentosDemandante", "—")).border = border


        for col in range(1, total_cols + 1):
            ws.cell(row=r_idx, column=col).alignment = Alignment(vertical="top", wrap_text=True)

    # Validacion de datos (desplegable real) en columna Estado
    estado_letter = get_column_letter(ESTADO_COL)
    dv = DataValidation(
        type="list",
        formula1='"' + ",".join(ESTADOS) + '"',
        allow_blank=False,
        showDropDown=False,  # False = SI muestra la flecha del desplegable (comportamiento openpyxl)
    )
    dv.error = "Selecciona un valor de la lista."
    dv.errorTitle = "Estado no válido"
    dv.prompt = "Elige el estado del procedimiento"
    dv.promptTitle = "Estado"
    last_row = max(len(rows) + 1, 2)
    dv.add(f"{estado_letter}2:{estado_letter}{last_row + 200}")  # margen para filas añadidas a mano
    ws.add_data_validation(dv)

    # Anchos de columna
    widths = [26, 40, 30, 16, 28, 32, 32, 42, 34, 36, 26, 16, 22, 30, 34, 36, 36]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(total_cols)}{last_row}"
    ws_meta = wb.create_sheet(META_SHEET_NAME)
    ws_meta.append(["_row_key"] + MERGE_COLUMNS)
    for key, row_meta in meta_rows:
            ws_meta.append([key] + [row_meta.get(c, "") for c in MERGE_COLUMNS])
    ws_meta.sheet_state = "hidden"
    wb.save(output_path)
    return output_path
