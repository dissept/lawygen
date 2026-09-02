"""
excel_builder.py
Genera el Excel de seguimiento con:
  - Una sola hoja
  - Columnas: Asunto, Procedimiento, Juzgado, Partes, Estado, Carpeta, Documento
  - Estado: desplegable real (validacion de datos de Excel)
  - Documento: hipervinculo directo al PDF principal (ruta absoluta local)
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

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


def build_workbook(rows, output_path):
    """rows: lista de dicts con keys = FIELD_KEYS + 'Carpeta' + 'Documento'
    (tupla (texto_visible, ruta_absoluta) o None) + 'Adjuntos' (texto)."""
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

    total_cols = len(HEADERS)
    for r_idx, row in enumerate(rows, start=2):
        for col, key in enumerate(FIELD_KEYS, start=1):
            value = row.get(key, "")
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

        ws.cell(row=r_idx, column=ADJUNTOS_COL, value=row.get("Adjuntos", "—")).border = border
        ws.cell(row=r_idx, column=DEMANDADO_COL, value=row.get("DocumentosDemandado", "—")).border = border
        ws.cell(row=r_idx, column=DEMANDANTE_COL, value=row.get("DocumentosDemandante", "—")).border = border

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

    wb.save(output_path)
    return output_path
