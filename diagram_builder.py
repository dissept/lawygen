"""
diagram_builder.py
Genera un diagrama de flujo visual (imagen PNG) a partir de una lista
ordenada de pasos/fases procesales, para insertarlo en el Word del Esquema.
Usa matplotlib (offline, sin dependencias de Node/Mermaid/Graphviz).

Paleta seria/profesional (azul marino y grises) porque estos documentos
pueden enviarse a juzgados, contrapartes o clientes — no llevan el morado
de la interfaz de la app, que es solo uso interno.

El alto de cada caja y el alto total de la figura se calculan a partir del
propio contenido (longitud de texto, numero de pasos) en vez de usar un
tamaño fijo — un caso con 3 pasos y otro con 15 pasos se ven bien sin
tocar codigo ni truncar texto en silencio.
"""
import textwrap
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Circle

# Cambiar aqui para recolorear todo el diagrama.
PALETTE = {
    "box_edge":    "#1E293B",
    "box_mid":     "#334155",
    "arrow":       "#94A3B8",
    "title":       "#1E293B",
    "step_number": "#1E293B",
    "step_badge":  "#E2E8F0",
    "text":        "#FFFFFF",
}

BOX_WIDTH = 8.6
WRAP_CHARS = 62
FONT_SIZE = 9.5
LINE_HEIGHT_IN = 0.21           # alto aproximado de una linea de texto, en pulgadas
BOX_PAD_V = 0.38                # relleno vertical (arriba+abajo) dentro de cada caja
GAP = 0.55
MAX_LINES_PER_BOX = 5           # a partir de aqui se corta con "…" en vez de desbordar
MAX_FIG_H = 16.0                # tope de alto de la imagen; si se supera, se escala todo


def _wrap_step(text):
    lines = textwrap.wrap(text, width=WRAP_CHARS) or [""]
    if len(lines) > MAX_LINES_PER_BOX:
        lines = lines[:MAX_LINES_PER_BOX]
        lines[-1] = lines[-1].rstrip() + "…"
    return lines


def build_flowchart(steps, out_path, title=None):
    """steps: lista ordenada de strings (cada uno = una fase/clausula del
    procedimiento). Genera un diagrama vertical tipo flujo, con flechas
    conectando cada caja con la siguiente, y lo guarda como PNG.

    El alto de cada caja sale del texto que contiene (mas lineas -> caja
    mas alta) en vez de ser fijo. Si el diagrama completo superaria
    MAX_FIG_H, todo se escala hacia abajo -cajas, huecos y letra- para
    que la imagen siga siendo razonable."""
    if not steps:
        return None

    n = len(steps)
    title_h = 0.9 if title else 0.0

    wrapped_steps = [_wrap_step(s) for s in steps]
    box_heights = [BOX_PAD_V + len(lines) * LINE_HEIGHT_IN for lines in wrapped_steps]

    raw_fig_h = title_h + sum(box_heights) + n * GAP
    scale = min(1.0, MAX_FIG_H / raw_fig_h) if raw_fig_h > 0 else 1.0

    box_heights = [h * scale for h in box_heights]
    gap = GAP * scale
    fig_h = title_h + sum(box_heights) + n * gap
    fig_w = 9.5

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.set_xlim(0, fig_w)
    ax.set_ylim(0, fig_h)
    ax.axis("off")
    fig.patch.set_alpha(0)

    if title:
        ax.text(fig_w / 2, fig_h - title_h / 2, title, ha="center", va="center",
                fontsize=15 * scale, fontweight="bold", color=PALETTE["title"])

    x0 = (fig_w - BOX_WIDTH) / 2
    y = fig_h - title_h
    centers = []

    for i, (step, lines, box_h) in enumerate(zip(steps, wrapped_steps, box_heights)):
        y -= box_h
        is_first = i == 0
        is_last = i == n - 1
        face = PALETTE["box_edge"] if is_first or is_last else PALETTE["box_mid"]

        box = FancyBboxPatch(
            (x0, y), BOX_WIDTH, box_h,
            boxstyle="round,pad=0.05,rounding_size=0.16",
            linewidth=0, facecolor=face,
        )
        ax.add_patch(box)

        ax.text(x0 + BOX_WIDTH / 2, y + box_h / 2, "\n".join(lines),
                ha="center", va="center", fontsize=FONT_SIZE * scale,
                color=PALETTE["text"],
                fontweight="bold" if (is_first or is_last) else "normal",
                linespacing=1.4)

               # numero de paso a la izquierda de la caja, en una insignia circular
        badge_cx, badge_cy = x0 - 0.35, y + box_h / 2
        ax.add_patch(Circle((badge_cx, badge_cy), 0.20 * max(scale, 0.6),
                             facecolor=PALETTE["step_badge"], linewidth=0,
                             clip_on=False, zorder=3))
        ax.text(badge_cx, badge_cy, str(i + 1), ha="center", va="center",
                fontsize=11 * scale, color=PALETTE["step_number"], fontweight="bold",
                zorder=4)

        centers.append((x0 + BOX_WIDTH / 2, y, box_h))
        y -= gap

    # flechas entre cajas consecutivas
    for i in range(len(centers) - 1):
        x1, y1, h1 = centers[i]
        x2, y2, h2 = centers[i + 1]
        ax.annotate(
            "", xy=(x2, y2 + h2), xytext=(x1, y1),
            arrowprops=dict(arrowstyle="-|>", color=PALETTE["arrow"], lw=2.2 * scale,
                             shrinkA=0, shrinkB=0, mutation_scale=20 * scale),
        )

    fig.tight_layout(pad=0.4)
    fig.savefig(out_path, dpi=180, transparent=True, bbox_inches="tight")
    plt.close(fig)
    return out_path