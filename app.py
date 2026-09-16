"""
app.py — LawyGen
Punto de entrada. Ejecutar con:  python app.py
Interfaz grafica de escritorio (Tkinter, incluido en Python estandar).
"""
import os
import sys
import threading
import queue
import traceback
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import processor
import excel_builder
import config_store

LINKEDIN_URL = "https://www.linkedin.com/in/ana-berjano-86290728a/"
_LINKEDIN_URL = "https://www.linkedin.com/in/anastasiia-pertsova/"

def _asset_path(*parts):
    """Resuelve rutas a /assets tanto en desarrollo (python app.py) como
    empaquetado como .exe con PyInstaller (que extrae a una carpeta temporal)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "assets", *parts)

# ---------------------------------------------------------------------------
# Paleta "morado LawyGen"
# ---------------------------------------------------------------------------
PURPLE_900 = "#2E1065"
PURPLE_700 = "#6D28D9"
PURPLE_600 = "#7C3AED"
PURPLE_500 = "#8B5CF6"
PURPLE_400 = "#A78BFA"
PURPLE_100 = "#EDE9FE"
PURPLE_50 = "#F5F3FF"
WHITE = "#FFFFFF"
TEXT_DARK = "#312E5B"
TEXT_MUTED = "#6B6180"
GREEN = "#15803D"
ORANGE = "#B45309"
RED = "#B91C1C"


class HoverButton(tk.Button):
    """Boton con efecto hover dinamico (cambia de tono al pasar el mouse)."""
    def __init__(self, master, bg, hover_bg, **kwargs):
        super().__init__(
            master, bg=bg, activebackground=hover_bg, fg="white",
            activeforeground="white", bd=0, relief="flat", cursor="hand2",
            font=("Segoe UI", 10, "bold"), **kwargs
        )
        self._bg = bg
        self._hover_bg = hover_bg
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, _e):
        if self["state"] != "disabled":
            self.config(bg=self._hover_bg)

    def _on_leave(self, _e):
        if self["state"] != "disabled":
            self.config(bg=self._bg)

    def set_disabled(self, disabled):
        if disabled:
            self.config(state="disabled", bg="#C4B5FD", cursor="arrow")
        else:
            self.config(state="normal", bg=self._bg, cursor="hand2")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("LawyGen")
        self.geometry("820x640")
        self.minsize(720, 560)
        self.configure(bg=PURPLE_50)
        self.root_folder = tk.StringVar()
        self.status_text = tk.StringVar(value="Listo")
        self._buttons = []
        self._event_queue = queue.Queue()
        self._build_style()
        self._build_ui()
        self.after(80, self._poll_queue)

    # ------------------------------------------------------------------
    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Purple.Horizontal.TProgressbar",
                         troughcolor=PURPLE_100, background=PURPLE_600,
                         bordercolor=PURPLE_100, lightcolor=PURPLE_500, darkcolor=PURPLE_600)

    def _build_ui(self):
        # ---------- Encabezado ----------
        header = tk.Frame(self, bg=PURPLE_900)
        header.pack(fill="x")

        top_row = tk.Frame(header, bg=PURPLE_900)
        top_row.pack(fill="x", padx=20, pady=(14, 0))

        left_block = tk.Frame(top_row, bg=PURPLE_900)
        left_block.pack(side="left", anchor="w")

        self.logo_img = None
        logo_path = _asset_path("lawygen_logo_header.png")
        try:
            self.logo_img = tk.PhotoImage(file=logo_path)
            tk.Label(left_block, image=self.logo_img, bg=PURPLE_900).pack(side="left", padx=(0, 10))
        except Exception:
            pass  # si falta el asset, seguimos solo con el texto

        title_block = tk.Frame(left_block, bg=PURPLE_900)
        title_block.pack(side="left")
        tk.Label(title_block, text="LawyGen", font=("Segoe UI", 20, "bold"),
                 bg=PURPLE_900, fg="white").pack(anchor="w")

        credits_block = tk.Frame(top_row, bg=PURPLE_900)
        credits_block.pack(side="right", anchor="n", pady=(4, 0))

        link = tk.Label(credits_block, text="Made by Ana", font=("Segoe UI", 9, "underline"),
                         bg=PURPLE_900, fg=PURPLE_100, cursor="hand2")
        link.pack(side="top", anchor="e")
        link.bind("<Button-1>", lambda e: webbrowser.open(LINKEDIN_URL))
        link.bind("<Enter>", lambda e: link.config(fg="white"))
        link.bind("<Leave>", lambda e: link.config(fg=PURPLE_100))

        link2 = tk.Label(credits_block, text="& Asiia", font=("Segoe UI", 9, "underline"),
                          bg=PURPLE_900, fg=PURPLE_100, cursor="hand2")
        link2.pack(side="top", anchor="e")
        link2.bind("<Button-1>", lambda e: webbrowser.open(_LINKEDIN_URL))
        link2.bind("<Enter>", lambda e: link2.config(fg="white"))
        link2.bind("<Leave>", lambda e: link2.config(fg=PURPLE_100))

        tk.Label(header, text="Generación automática de seguimiento legal a partir de expedientes",
                 font=("Segoe UI", 10), bg=PURPLE_900, fg=PURPLE_100).pack(anchor="w", padx=20, pady=(2, 16))

        # icono de la ventana (si el asset esta disponible)
        try:
            self.iconphoto(True, tk.PhotoImage(file=_asset_path("lawygen_icon.png")))
        except Exception:
            pass

        body = tk.Frame(self, bg=PURPLE_50)
        body.pack(fill="both", expand=True, padx=20, pady=16)

        # ---------- Selector de carpeta ----------
        card1 = self._card(body)
        card1.pack(fill="x", pady=(0, 14))
        tk.Label(card1, text="1. Carpeta de casos", font=("Segoe UI", 11, "bold"),
                 bg=WHITE, fg=TEXT_DARK).pack(anchor="w", padx=16, pady=(14, 4))
        tk.Label(card1, text="Selecciona la carpeta que contiene una subcarpeta por cada asunto/cliente.",
                 font=("Segoe UI", 9), bg=WHITE, fg=TEXT_MUTED, wraplength=740, justify="left").pack(
            anchor="w", padx=16, pady=(0, 10))

        folder_row = tk.Frame(card1, bg=WHITE)
        folder_row.pack(fill="x", padx=16, pady=(0, 16))
        entry = tk.Entry(folder_row, textvariable=self.root_folder, font=("Segoe UI", 10),
                          relief="flat", highlightthickness=1, highlightbackground=PURPLE_400,
                          highlightcolor=PURPLE_600)
        entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))
        HoverButton(folder_row, bg=PURPLE_600, hover_bg=PURPLE_700, text="Elegir carpeta...",
                    command=self._choose_folder).pack(side="left", ipadx=10, ipady=6)

        # ---------- Botones de accion (los 3 pedidos) ----------
        card2 = self._card(body)
        card2.pack(fill="x", pady=(0, 14))
        tk.Label(card2, text="2. Generar", font=("Segoe UI", 11, "bold"),
                 bg=WHITE, fg=TEXT_DARK).pack(anchor="w", padx=16, pady=(14, 10))

        btn_row = tk.Frame(card2, bg=WHITE)
        btn_row.pack(fill="x", padx=16, pady=(0, 16))
        btn_row.columnconfigure((0, 1, 2), weight=1, uniform="btns")

        self.btn_excel = HoverButton(btn_row, bg=PURPLE_600, hover_bg=PURPLE_700,
                                      text="📊  Generar Excel",
                                      command=lambda: self._start_processing(processor.MODE_EXCEL))
        self.btn_resumen = HoverButton(btn_row, bg=PURPLE_500, hover_bg=PURPLE_600,
                                        text="📝  Generar Resumen",
                                        command=lambda: self._start_processing(processor.MODE_RESUMEN))
        self.btn_esquema = HoverButton(btn_row, bg=PURPLE_500, hover_bg=PURPLE_600,
                                        text="🗂️  Generar Esquema",
                                        command=lambda: self._start_processing(processor.MODE_ESQUEMA))
        for i, b in enumerate([self.btn_excel, self.btn_resumen, self.btn_esquema]):
            b.grid(row=0, column=i, sticky="ew", padx=6, ipady=10)
            self._buttons.append(b)

        # ---------- Estado dinamico ----------
        status_row = tk.Frame(body, bg=PURPLE_50)
        status_row.pack(fill="x", pady=(0, 6))
        self.status_dot = tk.Canvas(status_row, width=12, height=12, bg=PURPLE_50, highlightthickness=0)
        self.status_dot.pack(side="left", padx=(2, 6))
        self._dot_id = self.status_dot.create_oval(2, 2, 10, 10, fill=PURPLE_400, outline="")
        tk.Label(status_row, textvariable=self.status_text, font=("Segoe UI", 9, "bold"),
                 bg=PURPLE_50, fg=TEXT_DARK).pack(side="left")

        self.progress = ttk.Progressbar(body, mode="determinate", style="Purple.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 12))

        # ---------- Registro de actividad ----------
        card3 = self._card(body)
        card3.pack(fill="both", expand=True)
        tk.Label(card3, text="Registro de actividad", font=("Segoe UI", 11, "bold"),
                 bg=WHITE, fg=TEXT_DARK).pack(anchor="w", padx=16, pady=(14, 6))
        self.log_box = scrolledtext.ScrolledText(card3, height=14, state="disabled",
                                                   font=("Consolas", 9), relief="flat",
                                                   bg="#FAFAFF", fg=TEXT_DARK, bd=0)
        self.log_box.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self.log_box.tag_config("info", foreground=TEXT_DARK)
        self.log_box.tag_config("aviso", foreground=ORANGE)
        self.log_box.tag_config("error", foreground=RED)
        self.log_box.tag_config("exito", foreground=GREEN)

    def _card(self, parent):
        return tk.Frame(parent, bg=WHITE, highlightbackground=PURPLE_100, highlightthickness=1)

    # ------------------------------------------------------------------
    def _choose_folder(self):
        path = filedialog.askdirectory(title="Selecciona la carpeta raíz de casos")
        if path:
            self.root_folder.set(path)

    def _set_status(self, text, color=PURPLE_400):
        self.status_text.set(text)
        self.status_dot.itemconfig(self._dot_id, fill=color)
        self.update_idletasks()

    def _log(self, msg, tag="info"):
        prefix = {"aviso": "⚠ ", "error": "✖ ", "exito": "✔ "}.get(tag, "")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", prefix + str(msg) + "\n", tag)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def _log_router(self, msg):
        text = str(msg)
        if "[ERROR" in text:
            self._event_queue.put(("log", (text.replace("[ERROR]", "").replace("[ERROR CRÍTICO]", "").strip(), "error")))
        elif "[AVISO]" in text:
            self._event_queue.put(("log", (text.replace("[AVISO]", "").strip(), "aviso")))
        else:
            self._event_queue.put(("log", (text, "info")))

    def _set_buttons_enabled(self, enabled):
        for b in self._buttons:
            b.set_disabled(not enabled)

    def _start_processing(self, mode):
        root_path = self.root_folder.get().strip()
        if not root_path or not os.path.isdir(root_path):
            messagebox.showerror("Error", "Selecciona primero una carpeta válida.")
            return
        self._set_buttons_enabled(False)
        self.progress["value"] = 0
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        labels = {
            processor.MODE_EXCEL: "Generando Excel...",
            processor.MODE_RESUMEN: "Generando Resúmenes...",
            processor.MODE_ESQUEMA: "Generando Esquemas...",
        }
        self._set_status(labels[mode], PURPLE_600)

        t = threading.Thread(target=self._run_pipeline, args=(root_path, mode), daemon=True)
        t.start()

    def _run_pipeline(self, root_path, mode):
        q = self._event_queue
        try:
            subfolders = [
                n for n in os.listdir(root_path)
                if os.path.isdir(os.path.join(root_path, n)) and not n.startswith(".")
            ]
            total = len(subfolders)
            if total == 0:
                q.put(("log", ("No se encontraron subcarpetas dentro de la carpeta seleccionada.", "aviso")))
                q.put(("finish", (PURPLE_400, "Listo")))
                return

            q.put(("progress_max", total))

            def on_progress(i, tot):
                q.put(("progress", i))
                q.put(("status", (f"Procesando {i}/{tot} carpetas...", PURPLE_600)))

            q.put(("log", (f"Encontradas {total} subcarpetas. Iniciando procesamiento...", "info")))
            if config_store.is_ai_enabled():
                q.put(("log", (f"IA activada (modelo: {config_store.get_model()}).\n", "info")))
            else:
                q.put(("log", ("Modo gratuito: usando solo reglas de texto (sin IA).\n", "info")))

            rows, generated_files = processor.process_root_folder(
                root_path, mode, progress_callback=on_progress, log=self._log_router
            )

            if mode == processor.MODE_EXCEL:
                out_path = os.path.join(root_path, "Seguimiento_Legal.xlsx")
                excel_builder.build_workbook(rows, out_path)
                q.put(("log", (f"\nExcel generado: {out_path}", "exito")))
                q.put(("log", (f"Filas totales: {len(rows)}", "exito")))
                msg = f"Proceso completado.\n\nExcel generado en:\n{out_path}"
            else:
                if mode == processor.MODE_RESUMEN:
                    q.put(("log", ("\nResumen General consolidado generado.", "exito")))
                    for p in generated_files:
                        q.put(("log", (f"  {p}", "info")))
                    msg = f"Proceso completado.\n\nResumen General guardado en:\n{generated_files[0] if generated_files else '(sin datos)'}"
                else:
                    q.put(("log", (f"\nEsquemas generados: {len(generated_files)} archivo(s) .docx (con diagrama visual)", "exito")))
                    for p in generated_files:
                        q.put(("log", (f"  {p}", "info")))
                    msg = "Proceso completado.\n\nEsquemas con diagrama generados dentro de cada subcarpeta."

            q.put(("log", ("\nRevisa los campos marcados como 'revisar' — el dato no pudo "
                            "determinarse con certeza a partir del contenido.", "aviso")))
            q.put(("status", ("Completado", GREEN)))
            q.put(("info", msg))
        except Exception as e:
            q.put(("log", (f"\n[ERROR CRÍTICO] {e}", "error")))
            q.put(("log", (traceback.format_exc(), "info")))
            q.put(("status", ("Error", RED)))
            q.put(("error", f"Ocurrió un error:\n{e}"))
        finally:
            q.put(("finish", (None, None)))

    def _finish(self, dot_color=None, status_label=None):
        self._set_buttons_enabled(True)
        if status_label:
            self._set_status(status_label, dot_color or PURPLE_400)

    def _poll_queue(self):
        """Corre en el hilo principal via .after(). Es el UNICO lugar donde
        se tocan widgets con datos que vienen del hilo de fondo -- Tkinter
        no es thread-safe, así que el hilo de trabajo nunca debe llamar
        directamente a self._log/self._set_status/messagebox/etc."""
        try:
            while True:
                kind, payload = self._event_queue.get_nowait()
                if kind == "log":
                    self._log(*payload)
                elif kind == "status":
                    self._set_status(*payload)
                elif kind == "progress":
                    self.progress["value"] = payload
                elif kind == "progress_max":
                    self.progress.config(maximum=payload)
                elif kind == "info":
                    messagebox.showinfo("LawyGen", payload)
                elif kind == "error":
                    messagebox.showerror("Error", payload)
                elif kind == "finish":
                    self._finish(*payload)
        except queue.Empty:
            pass
        self.after(80, self._poll_queue)

if __name__ == "__main__":
    try:
        app = App()
        app.mainloop()
    except Exception:
        traceback.print_exc()
        input("Ocurrió un error. Presiona Enter para salir...")
