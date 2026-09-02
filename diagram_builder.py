"""
diagram_builder.py
Genera un diagrama de flujo visual (imagen PNG) a partir de una lista
ordenada de pasos/fases procesales, para insertarlo en el Word del Esquema.
Usa matplotlib (offline, sin dependencias de Node/Mermaid/Graphviz).

Paleta seria/profesional (azul marino y grises) porque estos documentos
pueden enviarse a juzgados, contrapartes o clientes — no llevan el morado
de la interfaz de la app, que es solo uso interno.
"""
import textwrap
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

NAVY_DARK = "#1E293B"     # cajas de inicio/fin
NAVY = "#334155"          # cajas intermedias
NAVY_LIGHT = "#94A3B8"    # flechas
TITLE_COLOR = "#1E293B"
STEP_NUMBER_COLOR = "#1E293B"
TEXT_WHITE = "#FFFFFF"

BOX_WIDTH = 8.6
BOX_HEIGHT = 1.15
GAP = 0.55
WRAP_CHARS = 62


def build_flowchart(steps, out_path, title=None):
    """steps: lista ordenada de strings (cada uno = una fase/clausula del
    procedimiento). Genera un diagrama vertical tipo flujo, con flechas
    conectando cada caja con la siguiente, y lo guarda como PNG."""
    if not steps:
        return None

    n = len(steps)
    title_h = 0.9 if title else 0.0
    fig_h = title_h + n * (BOX_HEIGHT + GAP)
    fig_w = 9.5

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.set_xlim(0, fig_w)
    ax.set_ylim(0, fig_h)
    ax.axis("off")
    fig.patch.set_alpha(0)

    if title:
        ax.text(fig_w / 2, fig_h - title_h / 2, title, ha="center", va="center",
                fontsize=15, fontweight="bold", color=TITLE_COLOR)

    x0 = (fig_w - BOX_WIDTH) / 2
    y = fig_h - title_h - BOX_HEIGHT
    centers = []

    for i, step in enumerate(steps):
        is_first = i == 0
        is_last = i == n - 1
        face = NAVY_DARK if is_first or is_last else NAVY

        box = FancyBboxPatch(
            (x0, y), BOX_WIDTH, BOX_HEIGHT,
            boxstyle="round,pad=0.05,rounding_size=0.10",
            linewidth=0, facecolor=face,
        )
        ax.add_patch(box)

        wrapped = "\n".join(textwrap.wrap(step, width=WRAP_CHARS)[:3])
        ax.text(x0 + BOX_WIDTH / 2, y + BOX_HEIGHT / 2, wrapped,
                ha="center", va="center", fontsize=9.5, color=TEXT_WHITE,
                fontweight="bold" if (is_first or is_last) else "normal",
                linespacing=1.4)

        # numero de paso a la izquierda de la caja
        ax.text(x0 - 0.35, y + BOX_HEIGHT / 2, str(i + 1), ha="center", va="center",
                fontsize=11, color=STEP_NUMBER_COLOR, fontweight="bold")

        centers.append((x0 + BOX_WIDTH / 2, y))
        y -= (BOX_HEIGHT + GAP)

    # flechas entre cajas consecutivas
    for i in range(len(centers) - 1):
        x1, y1 = centers[i]
        x2, y2 = centers[i + 1]
        ax.annotate(
            "", xy=(x2, y2 + BOX_HEIGHT), xytext=(x1, y1),
            arrowprops=dict(arrowstyle="-|>", color=NAVY_LIGHT, lw=2.2,
                             shrinkA=0, shrinkB=0, mutation_scale=20),
        )

    fig.tight_layout(pad=0.4)
    fig.savefig(out_path, dpi=180, transparent=True, bbox_inches="tight")
    plt.close(fig)
    return out_path
