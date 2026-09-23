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
import datetime
from urllib.parse import quote
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule

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

def _file_uri(abs_path):
    """Convierte una ruta local en un URI file:/// valido (espacios, tildes,
    ñ, '#', '%'... codificados). Excel 2007 rechaza el archivo entero si un
    hipervinculo contiene una ruta sin codificar; las versiones modernas lo
    toleran, por eso el fallo solo aparecia en equipos con Office antiguo."""
    path = abs_path.replace("\\", "/")
    if path.startswith("//"):          # ruta de red \\servidor\carpeta
        return "file:" + quote(path, safe="/:")
    return "file:///" + quote(path.lstrip("/"), safe="/:")


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
        # La fila de encabezados es la que lleva las marcas: la fila 1 en
        # versiones antiguas, la fila HEADER_ROW en el diseno actual (con
        # banda de titulo encima). Se busca en las primeras filas.
        col_index = {}
        header_row = 1
        for r in range(1, min(ws_old.max_row, 8) + 1):
            for cell in ws_old[r]:
                comment = getattr(cell, "comment", None)
                if not comment or not comment.text:
                    continue
                first_line = comment.text.split("\n", 1)[0]
                if first_line.startswith("LAWYGEN_KEY:"):
                    col_index[first_line[len("LAWYGEN_KEY:"):]] = cell.column
            if col_index:
                header_row = r
                break

        if not col_index:
            # Archivo de una version anterior a esta marca (o alguien borro
            # las notas a mano): no hay forma fiable de saber que columna es
            # cada cosa. Mejor no arriesgar un merge mal hecho -- se
            # regenera todo desde cero esta vez, como si no hubiera archivo
            # previo.
            return old_values, old_meta

        n_cols_old = ws_old.max_column
        for row_cells in ws_old.iter_rows(min_row=header_row + 1):
            values = {
                name: (row_cells[idx - 1].value if idx <= n_cols_old else None)
                for name, idx in col_index.items()
            }
            if not values.get("Carpeta"):
                continue
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

# ---------------------------------------------------------------------------
# Diseno visual (colores de la app LawyGen). Todo compatible con Excel 2007:
# rellenos, bordes, formato condicional clasico y formulas COUNTIF/REPT.
# ---------------------------------------------------------------------------
PURPLE_900 = "2E1065"
PURPLE_800 = "4C1D95"
PURPLE_600 = "7C3AED"
PURPLE_100 = "EDE9FE"
PURPLE_50 = "F7F5FE"
TEXT_DARK = "312E5B"
TEXT_MUTED = "6B6180"
GRID = "E4E0F0"
FONT_NAME = "Arial"

TITLE_ROW = 1
SUBTITLE_ROW = 2
HEADER_ROW = 3
FIRST_DATA_ROW = HEADER_ROW + 1

# (relleno, texto) de cada Estado; se aplica con formato condicional para que
# el color cambie solo cuando alguien elige otro valor en el desplegable.
ESTADO_COLORS = {
    "Abierto": ("DBEAFE", "1E40AF"),
    "En trámite": ("FEF3C7", "92400E"),
    "En apelación": ("FFEDD5", "9A3412"),
    "Suspendido": ("E5E7EB", "374151"),
    "Archivado": ("E2E8F0", "475569"),
    "Resuelto": ("DCFCE7", "166534"),
}
MATERIAS = ["Civil", "Penal", "Laboral", "Mercantil", "Contencioso-Administrativo"]

# Columnas con varios valores separados por " | ": se muestran uno por linea.
BULLET_COLUMNS = {"Argumentos", "FundamentosDerecho", "DocumentosDemandado", "DocumentosDemandante"}
MULTILINE_COLUMNS = {"Partes", "Representantes", "CuantiaCostas"}
MUTED_COLUMNS = {"Carpeta", "Adjuntos", "DocumentosDemandado", "DocumentosDemandante"}

COLUMN_WIDTHS = {
    "Asunto": 28, "Resumen": 40, "Procedimiento": 24, "Materia": 16, "Juzgado": 28,
    "Partes": 34, "Representantes": 34, "Argumentos": 55, "Objeto": 42,
    "FundamentosDerecho": 50, "CuantiaCostas": 30, "Estado": 15, "Carpeta": 26,
    "Documento": 30, "Adjuntos": 32, "DocumentosDemandado": 36, "DocumentosDemandante": 36,
}


def _pretty(col, value):
    """Da formato legible a los campos con varios valores. Se aplica ANTES
    del merge, asi el valor guardado en _LawyGenMeta coincide con la celda y
    la deteccion de ediciones manuales sigue funcionando."""
    if not isinstance(value, str):
        return value
    if col == "Adjuntos" and "; " in value:
        return "\n".join("• " + p.strip() for p in value.split(";") if p.strip())
    if " | " not in value:
        return value
    parts = [p.strip() for p in value.split(" | ") if p.strip()]
    if col in BULLET_COLUMNS:
        return "\n".join("• " + p for p in parts)
    if col in MULTILINE_COLUMNS:
        return "\n".join(parts)
    return value


def _estimate_height(values_by_col):
    """Excel no ajusta el alto de fila al abrir un archivo generado; se
    estima a partir del texto de cada columna (aprox. 1,4 car. por unidad
    de ancho), con un tope para que ninguna fila sea gigantesca."""
    max_lines = 1
    for key, text in values_by_col.items():
        if not text:
            continue
        width = COLUMN_WIDTHS.get(key, 20)
        per_line = max(int(width * 1.4), 8)
        lines = 0
        for part in str(text).split("\n"):
            lines += max(1, -(-len(part) // per_line))
        max_lines = max(max_lines, lines)
    return min(8 + 12.5 * max_lines, 260)


def _font(**kw):
    kw.setdefault("name", FONT_NAME)
    kw.setdefault("size", 10)
    kw.setdefault("color", TEXT_DARK)
    return Font(**kw)


def _build_panel(wb, n_asuntos, estado_letter, materia_letter, last_row):
    """Hoja 'Panel' con recuento por Estado y por Materia (formulas vivas)."""
    ws = wb.create_sheet("Panel")
    ws.sheet_view.showGridLines = False
    title_fill = PatternFill("solid", start_color=PURPLE_900, end_color=PURPLE_900)
    head_fill = PatternFill("solid", start_color=PURPLE_800, end_color=PURPLE_800)
    for col, w in zip("ABCD", [30, 12, 34, 4]):
        ws.column_dimensions[col].width = w
    for c in range(1, 5):
        ws.cell(row=1, column=c).fill = title_fill
        ws.cell(row=2, column=c).fill = title_fill
    ws["A1"] = "LawyGen · Panel del seguimiento"
    ws["A1"].font = _font(size=16, bold=True, color="FFFFFF")
    ws["A2"] = "Se actualiza solo al cambiar el Estado o la Materia en la hoja Seguimiento Legal"
    ws["A2"].font = _font(size=9, color=PURPLE_100)
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 18

    data_ref = f"'Seguimiento Legal'!${{col}}${FIRST_DATA_ROW}:${{col}}${last_row + 500}"
    thin = Side(style="thin", color=GRID)

    def block(start_row, title, labels, col_letter, colors=None):
        for c, h in enumerate([title, "Asuntos", ""], start=1):
            cell = ws.cell(row=start_row, column=c, value=h)
            cell.fill = head_fill
            cell.font = _font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(vertical="center")
        ws.row_dimensions[start_row].height = 22
        rng = data_ref.format(col=col_letter)
        for i, label in enumerate(labels, start=1):
            r = start_row + i
            a = ws.cell(row=r, column=1, value=label)
            b = ws.cell(row=r, column=2, value=f'=COUNTIF({rng},A{r})')
            c = ws.cell(row=r, column=3, value=f'=REPT("■",B{r})')
            fg = colors[label][1] if colors else PURPLE_600
            if colors:
                a.fill = PatternFill("solid", start_color=colors[label][0], end_color=colors[label][0])
            a.font = _font(bold=True, color=fg if colors else TEXT_DARK)
            b.font = _font(bold=True)
            b.alignment = Alignment(horizontal="center")
            c.font = _font(color=fg, size=11)
            for cell in (a, b, c):
                cell.border = Border(bottom=thin)
            ws.row_dimensions[r].height = 19
        return start_row + len(labels) + 2

    r = 4
    ws.cell(row=r, column=1, value="Total de asuntos").font = _font(bold=True, size=11)
    tot = ws.cell(row=r, column=2, value=f"=COUNTA({data_ref.format(col=get_column_letter(CARPETA_COL))})")
    tot.font = _font(bold=True, size=14, color=PURPLE_900)
    tot.alignment = Alignment(horizontal="center")
    ws.cell(row=r + 1, column=1, value='Celdas con "revisar"').font = _font(bold=True, size=11)
    rev = ws.cell(row=r + 1, column=2,
                  value=f"=COUNTIF('Seguimiento Legal'!$A${FIRST_DATA_ROW}:$Q${last_row + 500},\"*revisar*\")")
    rev.font = _font(bold=True, size=14, color="B91C1C")
    rev.alignment = Alignment(horizontal="center")
    ws.row_dimensions[r].height = 22
    ws.row_dimensions[r + 1].height = 22

    r = block(r + 3, "Estado", ESTADOS, estado_letter, ESTADO_COLORS)
    block(r, "Materia", MATERIAS, materia_letter)
    return ws


def build_workbook(rows, output_path):
    """rows: lista de dicts con keys = FIELD_KEYS + 'Carpeta' + 'Documento'
    (tupla (texto_visible, ruta_absoluta) o None) + 'Adjuntos' (texto)."""
    old_values, old_meta = _load_previous(output_path)

    wb = Workbook()
    ws = wb.active
    ws.title = "Seguimiento Legal"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = PURPLE_600

    total_cols = len(HEADERS)
    last_col_letter = get_column_letter(total_cols)
    title_fill = PatternFill("solid", start_color=PURPLE_900, end_color=PURPLE_900)
    header_fill = PatternFill("solid", start_color=PURPLE_800, end_color=PURPLE_800)
    band_fill = PatternFill("solid", start_color=PURPLE_50, end_color=PURPLE_50)
    white_fill = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
    grid = Side(style="thin", color=GRID)
    row_border = Border(bottom=grid, left=grid, right=grid)
    header_border = Border(bottom=Side(style="medium", color=PURPLE_600),
                           left=Side(style="thin", color=PURPLE_900),
                           right=Side(style="thin", color=PURPLE_900))

    # ---------- Banda de titulo ----------
    for col in range(1, total_cols + 1):
        ws.cell(row=TITLE_ROW, column=col).fill = title_fill
        ws.cell(row=SUBTITLE_ROW, column=col).fill = title_fill
    ws.cell(row=TITLE_ROW, column=1, value="LawyGen · Seguimiento legal").font = \
        _font(size=16, bold=True, color="FFFFFF")
    ws.cell(row=SUBTITLE_ROW, column=1, value=(
        f"Generado el {datetime.datetime.now():%d/%m/%Y %H:%M}  ·  {len(rows)} "
        f"procedimiento{'s' if len(rows) != 1 else ''}  ·  Las celdas marcadas "
        "\"revisar\" requieren comprobación manual  ·  Resumen en la hoja Panel"
    )).font = _font(size=9, color=PURPLE_100)
    ws.row_dimensions[TITLE_ROW].height = 30
    ws.row_dimensions[SUBTITLE_ROW].height = 18

    # ---------- Encabezados ----------
    for col, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=HEADER_ROW, column=col, value=h)
        c.fill = header_fill
        c.font = _font(color="FFFFFF", bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = header_border
        c.comment = Comment(
            f"LAWYGEN_KEY:{COLUMN_KEYS[col - 1]}\n"
            "(No borrar esta nota: permite que LawyGen reconozca esta "
            "columna aunque renombres el título o la muevas de sitio.)",
            "LawyGen",
        )
    ws.row_dimensions[HEADER_ROW].height = 36

    # ---------- Filas ----------
    meta_rows = []
    for i, row in enumerate(rows):
        r_idx = FIRST_DATA_ROW + i
        row = dict(row)
        for key in MERGE_COLUMNS:
            row[key] = _pretty(key, row.get(key, ""))
        merged_row, row_meta = _merge_row(row, old_values, old_meta)
        meta_rows.append((_row_key(row), row_meta))
        fill = band_fill if i % 2 else white_fill

        cell_values = {}
        for col, key in enumerate(FIELD_KEYS, start=1):
            value = merged_row.get(key, "")
            if key == "Estado":
                value = value or "Abierto"
            ws.cell(row=r_idx, column=col, value=value)
            cell_values[key] = value

        ws.cell(row=r_idx, column=CARPETA_COL, value=row.get("Carpeta", ""))
        cell_values["Carpeta"] = row.get("Carpeta", "")

        doc_cell = ws.cell(row=r_idx, column=DOCUMENTO_COL)
        doc_info = row.get("Documento")
        if doc_info:
            display_name, abs_path = doc_info
            doc_cell.value = display_name
            doc_cell.hyperlink = _file_uri(abs_path)
        else:
            doc_cell.value = "revisar"
        cell_values["Documento"] = doc_cell.value

        for col, key in ((ADJUNTOS_COL, "Adjuntos"), (DEMANDADO_COL, "DocumentosDemandado"),
                         (DEMANDANTE_COL, "DocumentosDemandante")):
            v = merged_row.get(key, "—")
            ws.cell(row=r_idx, column=col, value=v)
            cell_values[key] = v

        # estilo de toda la fila
        for col, key in enumerate(COLUMN_KEYS, start=1):
            c = ws.cell(row=r_idx, column=col)
            c.fill = fill
            c.border = row_border
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if key == "Asunto":
                c.font = _font(bold=True, color=PURPLE_900)
            elif key in ("Estado", "Materia"):
                c.font = _font(bold=True, color=PURPLE_800 if key == "Materia" else TEXT_DARK)
                c.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
            elif key == "Documento" and doc_info:
                c.font = _font(color=PURPLE_600, underline="single")
            elif key in MUTED_COLUMNS:
                c.font = _font(size=9, color=TEXT_MUTED)
            else:
                c.font = _font()
        ws.row_dimensions[r_idx].height = _estimate_height(cell_values)

    last_row = max(FIRST_DATA_ROW + len(rows) - 1, FIRST_DATA_ROW)
    cf_last = last_row + 200  # margen para filas anadidas a mano

    # ---------- Desplegable de Estado ----------
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
    dv.add(f"{estado_letter}{FIRST_DATA_ROW}:{estado_letter}{cf_last}")
    ws.add_data_validation(dv)

    # ---------- Formato condicional (colores de Estado y "revisar") ----------
    estado_rng = f"{estado_letter}{FIRST_DATA_ROW}:{estado_letter}{cf_last}"
    for estado, (bg, fg) in ESTADO_COLORS.items():
        ws.conditional_formatting.add(estado_rng, CellIsRule(
            operator="equal", formula=[f'"{estado}"'],
            fill=PatternFill("solid", start_color=bg, end_color=bg, bgColor=bg),
            font=Font(color=fg, bold=True)))
    data_rng = f"A{FIRST_DATA_ROW}:{last_col_letter}{cf_last}"
    ws.conditional_formatting.add(data_rng, FormulaRule(
        formula=[f'ISNUMBER(SEARCH("revisar",A{FIRST_DATA_ROW}))'],
        fill=PatternFill("solid", start_color="FEF2F2", end_color="FEF2F2", bgColor="FEF2F2"),
        font=Font(color="B91C1C", italic=True)))

    # ---------- Anchos, paneles, filtro, impresion ----------
    for i, key in enumerate(COLUMN_KEYS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = COLUMN_WIDTHS.get(key, 20)
    ws.freeze_panes = f"B{FIRST_DATA_ROW}"  # fija titulo, encabezados y columna Asunto
    ws.auto_filter.ref = f"A{HEADER_ROW}:{last_col_letter}{last_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"
    ws.oddFooter.center.text = "LawyGen · Página &P de &N"

    _build_panel(wb, len(rows), estado_letter, get_column_letter(FIELD_KEYS.index("Materia") + 1), last_row)

    ws_meta = wb.create_sheet(META_SHEET_NAME)
    ws_meta.append(["_row_key"] + MERGE_COLUMNS)
    for key, row_meta in meta_rows:
        ws_meta.append([key] + [row_meta.get(c, "") for c in MERGE_COLUMNS])
    ws_meta.sheet_state = "hidden"
    wb.active = 0
    wb.save(output_path)