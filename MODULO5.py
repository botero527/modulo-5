"""
MODULO5.py — Aplicacion principal unificada AGP Glass
Modulo 5: Generador de Combinaciones + Bloqueos + Automatizacion SAP

Flujo:
  Pestaña 1 (COMBINACIONES) → ejecuta COMBINADOR.py
  Pestaña 2 (BLOQUEOS)      → gestiona bloqueos por color
  Pestaña 3 (SAP)           → ejecuta cambio de color masivo en SAP
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import queue
import os
import sys
import datetime
import subprocess

try:
    from PIL import Image, ImageTk
    PIL_OK = True
except ImportError:
    PIL_OK = False

# ── Rutas ─────────────────────────────────────────────────────────────────────
BASE_DIR           = r"C:\Users\abotero\OneDrive - AGP GROUP\Documentos\MODULO_5"
RUTA_COMBINACIONES = os.path.join(BASE_DIR, "combinaciones5.xlsx")
LOGO_PATH          = os.path.join(BASE_DIR, "agp-america-s-a-logo-vector-removebg-preview (1).png")

# ── Paleta AGP ────────────────────────────────────────────────────────────────
C = {
    "bg_header":       "#2B3A47",
    "agp_blue":        "#7DBFD4",
    "agp_blue_dark":   "#4E9CB5",
    "bg_app":          "#EFF3F7",
    "bg_card":         "#FFFFFF",
    "txt_primary":     "#1A2634",
    "txt_secondary":   "#637282",
    "txt_white":       "#FFFFFF",
    "txt_header":      "#C8D8E4",
    "activa_border":   "#27AE60",
    "bloqueada_border":"#E74C3C",
    "pendiente_border":"#F39C12",
    "btn_ok":          "#27AE60",
    "btn_ok_h":        "#1E8449",
    "btn_err":         "#E74C3C",
    "btn_err_h":       "#C0392B",
    "btn_neutral":     "#7F8C8D",
    "btn_neutral_h":   "#626D6E",
    "border_light":    "#D5E0EA",
    "tbl_hdr_border":  "#3D5166",
    "tab_active":      "#7DBFD4",
    "tab_bg":          "#243140",
}

FONT_TITLE  = ("Segoe UI", 15, "bold")
FONT_BOLD   = ("Segoe UI", 10, "bold")
FONT_LABEL  = ("Segoe UI", 10)
FONT_SMALL  = ("Segoe UI", 9)
FONT_MICRO  = ("Segoe UI", 8)
FONT_CAPS   = ("Segoe UI", 8, "bold")
FONT_METRIC = ("Segoe UI", 22, "bold")


# ── Boton moderno ─────────────────────────────────────────────────────────────

class Boton(tk.Label):
    def __init__(self, parent, texto, comando, bg, hover,
                 fg="#FFFFFF", ancho=13, **kwargs):
        super().__init__(
            parent, text=texto, bg=bg, fg=fg,
            font=FONT_BOLD, cursor="hand2",
            padx=16, pady=12, width=ancho,
            relief="flat", **kwargs
        )
        self._bg    = bg
        self._hover = hover
        self._cmd   = comando
        self.bind("<Enter>",    lambda e: self.config(bg=self._hover))
        self.bind("<Leave>",    lambda e: self.config(bg=self._bg))
        self.bind("<Button-1>", lambda e: self._run())

    def _run(self):
        self.config(bg=self._hover)
        self._cmd()
        self.after(120, lambda: self.config(bg=self._bg) if self.winfo_exists() else None)

    def set_estado(self, activo: bool):
        if activo:
            self.config(bg=self._bg, cursor="hand2")
            self.bind("<Button-1>", lambda e: self._run())
        else:
            self.config(bg=C["btn_neutral"], cursor="arrow")
            self.unbind("<Button-1>")


# ── Header compartido ─────────────────────────────────────────────────────────

def build_header(parent, subtitulo: str, logo_img_ref: list) -> tk.Frame:
    """Construye el header navy AGP. logo_img_ref = lista de 1 elemento para evitar GC."""
    outer = tk.Frame(parent, bg=C["bg_header"])
    tk.Frame(outer, bg=C["agp_blue"], height=4).pack(fill="x")

    inner = tk.Frame(outer, bg=C["bg_header"], padx=24, pady=12)
    inner.pack(fill="x")

    # Logo
    logo_frame = tk.Frame(inner, bg=C["bg_header"])
    logo_frame.pack(side="left")

    if PIL_OK and os.path.exists(LOGO_PATH):
        try:
            img   = Image.open(LOGO_PATH).convert("RGBA")
            fondo = Image.new("RGBA", img.size, color=(43, 58, 71, 255))
            fondo.paste(img, mask=img.split()[3])
            img   = fondo.convert("RGB")
            w, h  = img.size
            nuevo_h = 60
            img   = img.resize((int(w * nuevo_h / h), nuevo_h), Image.LANCZOS)
            tk_img = ImageTk.PhotoImage(img)
            logo_img_ref.append(tk_img)
            tk.Label(logo_frame, image=tk_img,
                     bg=C["bg_header"]).pack(padx=(0, 20))
        except Exception:
            _logo_texto(logo_frame)
    else:
        _logo_texto(logo_frame)

    tk.Frame(inner, bg=C["tbl_hdr_border"], width=1).pack(
        side="left", fill="y", padx=(0, 20), pady=4)

    titulo_frame = tk.Frame(inner, bg=C["bg_header"])
    titulo_frame.pack(side="left")
    tk.Label(titulo_frame, text="AGP GLASS  —  MODULO 5",
             bg=C["bg_header"], fg=C["agp_blue"], font=FONT_CAPS).pack(anchor="w")
    tk.Label(titulo_frame, text=subtitulo,
             bg=C["bg_header"], fg=C["txt_white"], font=FONT_TITLE).pack(anchor="w")

    tk.Frame(outer, bg=C["tbl_hdr_border"], height=1).pack(fill="x")
    return outer


def _logo_texto(parent):
    tk.Label(parent, text="AGP", bg=C["bg_header"],
             fg=C["agp_blue"], font=("Segoe UI", 24, "bold"), padx=8).pack()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — COMBINACIONES
# ══════════════════════════════════════════════════════════════════════════════

class TabCombinaciones(tk.Frame):

    def __init__(self, parent, on_combinaciones_listas=None):
        super().__init__(parent, bg=C["bg_app"])
        self._on_listas = on_combinaciones_listas
        self._logo      = []
        self._proceso   = None
        self._build()

    def _build(self):
        build_header(self, "Generador de Combinaciones", self._logo).pack(fill="x")

        # ── Cuerpo ────────────────────────────────────────────────────────────
        cuerpo = tk.Frame(self, bg=C["bg_app"], padx=32, pady=20)
        cuerpo.pack(fill="both", expand=True)

        # Instrucciones
        tk.Label(cuerpo,
                 text="Genera todas las combinaciones de formula × acero × color "
                      "para el ZFER base ingresado.",
                 bg=C["bg_app"], fg=C["txt_secondary"], font=FONT_SMALL,
                 wraplength=700, justify="left").pack(anchor="w", pady=(0, 20))

        # Card de configuracion
        card = tk.Frame(cuerpo, bg=C["bg_card"],
                        highlightthickness=1,
                        highlightbackground=C["border_light"],
                        padx=24, pady=20)
        card.pack(fill="x", pady=(0, 16))

        tk.Frame(card, bg=C["agp_blue"], width=4).pack(side="left", fill="y", padx=(0, 16))

        form = tk.Frame(card, bg=C["bg_card"])
        form.pack(side="left", fill="both", expand=True)

        tk.Label(form, text="CONFIGURACION", bg=C["bg_card"],
                 fg=C["txt_secondary"], font=FONT_CAPS).pack(anchor="w", pady=(0, 12))

        # ZFER
        fila1 = tk.Frame(form, bg=C["bg_card"])
        fila1.pack(fill="x", pady=4)
        tk.Label(fila1, text="ZFER base:", bg=C["bg_card"],
                 fg=C["txt_primary"], font=FONT_BOLD, width=14, anchor="w").pack(side="left")
        self._zfer_var = tk.StringVar()
        entry = tk.Entry(fila1, textvariable=self._zfer_var,
                         bg=C["bg_app"], fg=C["txt_primary"],
                         font=FONT_LABEL, width=20, relief="flat",
                         highlightthickness=1,
                         highlightbackground=C["border_light"])
        entry.pack(side="left", padx=(0, 12))
        tk.Label(fila1, text="Ej: 700179044",
                 bg=C["bg_card"], fg=C["txt_secondary"], font=FONT_MICRO).pack(side="left")
        tk.Label(fila1, text="  El mercado se detecta automaticamente desde la BD.",
                 bg=C["bg_card"], fg=C["txt_secondary"], font=FONT_MICRO).pack(side="left", padx=(16, 0))

        # Boton generar
        btn_frame = tk.Frame(form, bg=C["bg_card"])
        btn_frame.pack(anchor="w", pady=(16, 0))
        self._btn_generar = Boton(btn_frame, "Generar combinaciones",
                                   self._generar,
                                   C["agp_blue_dark"], C["agp_blue"],
                                   ancho=22)
        self._btn_generar.pack(side="left")

        # ── Log de salida ─────────────────────────────────────────────────────
        tk.Label(cuerpo, text="SALIDA", bg=C["bg_app"],
                 fg=C["txt_secondary"], font=FONT_CAPS).pack(anchor="w", pady=(8, 4))

        log_frame = tk.Frame(cuerpo, bg=C["bg_header"],
                             highlightthickness=1,
                             highlightbackground=C["border_light"])
        log_frame.pack(fill="both", expand=True)

        self._log = scrolledtext.ScrolledText(
            log_frame, bg="#1A2634", fg="#7DBFD4",
            font=("Consolas", 9), relief="flat",
            insertbackground="#7DBFD4", state="disabled",
            wrap="word", padx=12, pady=10,
        )
        self._log.pack(fill="both", expand=True)

        # ── Footer ────────────────────────────────────────────────────────────
        tk.Frame(self, bg=C["tbl_hdr_border"], height=1).pack(side="bottom", fill="x")
        footer = tk.Frame(self, bg=C["bg_header"], padx=20, pady=10)
        footer.pack(side="bottom", fill="x")
        self._btn_ir_bloqueos = Boton(footer, "Ir a Bloqueos →",
                                       self._ir_bloqueos,
                                       C["btn_ok"], C["btn_ok_h"],
                                       ancho=18)
        self._btn_ir_bloqueos.pack(side="right")
        self._btn_ir_bloqueos.set_estado(False)
        tk.Frame(self, bg=C["agp_blue"], height=4).pack(side="bottom", fill="x")

    def _log_write(self, texto: str):
        self._log.config(state="normal")
        self._log.insert("end", texto + "\n")
        self._log.see("end")
        self._log.config(state="disabled")

    def _generar(self):
        zfer = self._zfer_var.get().strip()
        if not zfer:
            messagebox.showwarning("ZFER requerido",
                                   "Ingresa el numero de ZFER base.", parent=self)
            return

        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")
        self._btn_generar.set_estado(False)
        self._btn_ir_bloqueos.set_estado(False)

        def _run():
            script = os.path.join(BASE_DIR, "COMBINADOR.py")
            python = sys.executable
            try:
                env = os.environ.copy()
                env["M5_ZFER_BASE"]            = zfer
                env["PYTHONIOENCODING"]        = "utf-8"
                env["PYTHONLEGACYWINDOWSSTDIO"] = "0"
                proc = subprocess.Popen(
                    [python, script],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", env=env,
                    cwd=BASE_DIR,
                )
                for linea in proc.stdout:
                    self.after(0, self._log_write, linea.rstrip())
                proc.wait()
                if proc.returncode == 0:
                    self.after(0, self._on_generar_ok)
                else:
                    self.after(0, self._log_write, "\n  [ERROR] COMBINADOR termino con error.")
                    self.after(0, lambda: self._btn_generar.set_estado(True))
            except Exception as e:
                self.after(0, self._log_write, f"  [ERROR] {e}")
                self.after(0, lambda: self._btn_generar.set_estado(True))

        threading.Thread(target=_run, daemon=True).start()

    def _on_generar_ok(self):
        self._log_write("\n  Combinaciones generadas correctamente.")
        self._btn_generar.set_estado(True)
        self._btn_ir_bloqueos.set_estado(True)
        if self._on_listas:
            self._on_listas()

    def _ir_bloqueos(self):
        if self._on_listas:
            self._on_listas()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — BLOQUEOS (embeds VistaPreviaBloqueos)
# ══════════════════════════════════════════════════════════════════════════════

class TabBloqueos(tk.Frame):

    def __init__(self, parent, on_listo_para_sap=None):
        super().__init__(parent, bg=C["bg_app"])
        self._on_listo   = on_listo_para_sap
        self._vista      = None
        self._build()

    def _build(self):
        # Footer primero (side=bottom), luego el cuerpo (expand=True)
        tk.Frame(self, bg=C["agp_blue"], height=4).pack(side="bottom", fill="x")
        tk.Frame(self, bg=C["tbl_hdr_border"], height=1).pack(side="bottom", fill="x")
        footer = tk.Frame(self, bg=C["bg_header"], padx=20, pady=10)
        footer.pack(side="bottom", fill="x")
        self._btn_sap = Boton(footer, "Ir a Automatizacion SAP →",
                               self._ir_sap,
                               C["btn_ok"], C["btn_ok_h"], ancho=26)
        self._btn_sap.pack(side="right")
        self._btn_sap.set_estado(False)

        # Placeholder ocupa el resto
        self._placeholder = tk.Frame(self, bg=C["bg_app"])
        self._placeholder.pack(fill="both", expand=True)

        tk.Label(self._placeholder,
                 text="Ejecuta primero la generacion de combinaciones\nen la pestana anterior.",
                 bg=C["bg_app"], fg=C["txt_secondary"],
                 font=FONT_LABEL, justify="center").pack(expand=True)

    def cargar(self, ruta_excel: str = RUTA_COMBINACIONES):
        """Carga o recarga las combinaciones desde el Excel."""
        if not os.path.exists(ruta_excel):
            messagebox.showwarning("Archivo no encontrado",
                                   f"No se encontro:\n{ruta_excel}\n\n"
                                   "Ejecuta primero la generacion de combinaciones.",
                                   parent=self)
            return

        # Destruir vista anterior si existe
        if self._vista:
            self._vista.destroy()
            self._vista = None

        self._placeholder.pack_forget()

        from VISTAAAA import VistaPreviaBloqueos
        self._vista = VistaPreviaBloqueos(
            self, ruta_excel,
            on_close=self._on_vista_cerrar,
        )
        self._vista.pack(fill="both", expand=True)
        self._btn_sap.set_estado(True)

    def _on_vista_cerrar(self):
        """Callback 'Cerrar' cuando VistaPreviaBloqueos esta embebida en este tab."""
        # Solo limpiar la vista y volver al placeholder — no cerrar la app
        if self._vista:
            self._vista.destroy()
            self._vista = None
        self._btn_sap.set_estado(False)
        self._placeholder.pack(fill="both", expand=True)

    def get_items_activos(self) -> list:
        """Retorna lista de ItemColor activos (no bloqueados) para el SAP."""
        if not self._vista:
            return []
        return [
            it for it in self._vista.items.values()
            if not it.bloqueado and not it.pendiente
        ]

    def _ir_sap(self):
        if self._on_listo:
            self._on_listo()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — AUTOMATIZACION SAP
# ══════════════════════════════════════════════════════════════════════════════

class TabSAP(tk.Frame):

    def __init__(self, parent, get_items_fn=None):
        super().__init__(parent, bg=C["bg_app"])
        self._get_items = get_items_fn   # funcion que retorna lista de ItemColor
        self._logo      = []
        self._queue     = queue.Queue()
        self._corriendo = False
        self._build()

    def _build(self):
        build_header(self, "Automatizacion SAP — Cambio de Color", self._logo).pack(fill="x")

        # ── Panel superior: metricas ──────────────────────────────────────────
        metricas_frame = tk.Frame(self, bg=C["bg_app"], padx=20, pady=14)
        metricas_frame.pack(fill="x")

        self._cards = {}
        for titulo, key, color in [
            ("A PROCESAR",  "total",  C["agp_blue_dark"]),
            ("EXITOSAS",    "ok",     C["activa_border"]),
            ("CON ERROR",   "error",  C["bloqueada_border"]),
            ("EN PROGRESO", "activo", C["pendiente_border"]),
        ]:
            card = tk.Frame(metricas_frame, bg=C["bg_card"],
                            highlightthickness=1,
                            highlightbackground=C["border_light"])
            card.pack(side="left", padx=(0, 10))
            tk.Frame(card, bg=color, width=4).pack(side="left", fill="y")
            inner = tk.Frame(card, bg=C["bg_card"], padx=14, pady=8)
            inner.pack(side="left")
            tk.Label(inner, text=titulo.upper(), bg=C["bg_card"],
                     fg=C["txt_secondary"], font=FONT_CAPS).pack(anchor="w")
            lbl = tk.Label(inner, text="0", bg=C["bg_card"],
                           fg=color, font=FONT_METRIC)
            lbl.pack(anchor="w")
            self._cards[key] = lbl

        # Barra de progreso
        prog_frame = tk.Frame(self, bg=C["bg_app"], padx=20)
        prog_frame.pack(fill="x", pady=(0, 8))
        self._progreso = ttk.Progressbar(prog_frame, mode="determinate",
                                          length=400)
        self._progreso.pack(side="left", padx=(0, 12))
        self._lbl_prog = tk.Label(prog_frame, text="",
                                   bg=C["bg_app"], fg=C["txt_secondary"],
                                   font=FONT_SMALL)
        self._lbl_prog.pack(side="left")

        # ── Log en vivo ───────────────────────────────────────────────────────
        tk.Label(self, text="  LOG EN VIVO", bg=C["bg_app"],
                 fg=C["txt_secondary"], font=FONT_CAPS).pack(anchor="w", padx=20)

        log_outer = tk.Frame(self, bg=C["bg_app"], padx=20)
        log_outer.pack(fill="both", expand=True, pady=(4, 0))

        log_frame = tk.Frame(log_outer, bg=C["bg_header"],
                             highlightthickness=1,
                             highlightbackground=C["border_light"])
        log_frame.pack(fill="both", expand=True)

        self._log = scrolledtext.ScrolledText(
            log_frame, bg="#1A2634", fg="#C8D8E4",
            font=("Consolas", 9), relief="flat",
            insertbackground="#7DBFD4", state="disabled",
            wrap="word", padx=12, pady=10,
        )
        self._log.pack(fill="both", expand=True)

        # Colores de lineas de log
        self._log.tag_configure("ok",    foreground="#52BE80")
        self._log.tag_configure("error", foreground="#E74C3C")
        self._log.tag_configure("warn",  foreground="#F39C12")
        self._log.tag_configure("info",  foreground="#7DBFD4")
        self._log.tag_configure("dim",   foreground="#637282")

        # ── Footer ────────────────────────────────────────────────────────────
        tk.Frame(self, bg=C["tbl_hdr_border"], height=1).pack(side="bottom", fill="x")
        footer = tk.Frame(self, bg=C["bg_header"], padx=20, pady=12)
        footer.pack(side="bottom", fill="x")

        left_f = tk.Frame(footer, bg=C["bg_header"])
        left_f.pack(side="left")
        tk.Label(left_f,
                 text="SAP GUI debe estar abierto y con sesion activa antes de iniciar.",
                 bg=C["bg_header"], fg=C["txt_header"], font=FONT_MICRO).pack(anchor="w")
        tk.Label(left_f,
                 text="Scripting habilitado: tuerca → Options → Accessibility & Scripting.",
                 bg=C["bg_header"], fg=C["txt_header"], font=FONT_MICRO).pack(anchor="w")

        right_f = tk.Frame(footer, bg=C["bg_header"])
        right_f.pack(side="right")

        self._btn_reporte = Boton(right_f, "Abrir reporte",
                                   self._abrir_reporte,
                                   C["btn_neutral"], C["btn_neutral_h"],
                                   ancho=14)
        self._btn_reporte.pack(side="left", padx=(0, 8))
        self._btn_reporte.set_estado(False)

        self._btn_iniciar = Boton(right_f, "Iniciar automatizacion",
                                   self._iniciar,
                                   C["agp_blue_dark"], C["agp_blue"],
                                   ancho=22)
        self._btn_iniciar.pack(side="left")

        tk.Frame(self, bg=C["agp_blue"], height=4).pack(side="bottom", fill="x")

        self._ultimo_reporte = ""

        # Arrancar poll de la queue
        self._poll_queue()

    # ── Log helper ────────────────────────────────────────────────────────────

    def _log_write(self, texto: str, tag: str = ""):
        self._log.config(state="normal")
        if tag:
            self._log.insert("end", texto + "\n", tag)
        else:
            self._log.insert("end", texto + "\n")
        self._log.see("end")
        self._log.config(state="disabled")

    def _card_set(self, key: str, valor):
        self._cards[key].config(text=str(valor))

    # ── Inicio del proceso ────────────────────────────────────────────────────

    def _obtener_formula_base(self, zfer: str) -> str:
        """
        Consulta rapida a BD de produccion para obtener la formula del ZFER base.
        Busca en dos tablas en orden:
          1. ZFER_Characteristics_Genesis  (SpecID = ZFER, columna FormulaCode)
          2. TCAL_CALENDARIO_COLOMBIA_DIRECT (ZFER = ZFER, columna Formula)
        Retorna "" si no encuentra en ninguna.
        """
        try:
            from SAP_AUTOMATIZADOR import DB_PROD
            import pyodbc as _pyodbc
            conn = _pyodbc.connect(
                f"DRIVER={{{DB_PROD['driver']}}};"
                f"SERVER={DB_PROD['server']};"
                f"DATABASE={DB_PROD['database']};"
                f"UID={DB_PROD['user']};"
                f"PWD={DB_PROD['password']};",
                autocommit=True, timeout=10,
            )
            cur = conn.cursor()

            # 1. ZFER_Characteristics_Genesis
            cur.execute(
                "SELECT TOP 1 FormulaCode FROM dbo.ZFER_Characteristics_Genesis "
                "WHERE SpecID = ?",
                (zfer,)
            )
            row = cur.fetchone()
            if row and row[0]:
                conn.close()
                return str(row[0]).strip()

            # 2. TCAL_CALENDARIO_COLOMBIA_DIRECT
            cur.execute(
                "SELECT TOP 1 Formula FROM dbo.TCAL_CALENDARIO_COLOMBIA_DIRECT "
                "WHERE ZFER = ?",
                (zfer,)
            )
            row = cur.fetchone()
            conn.close()
            if row and row[0]:
                return str(row[0]).strip()

        except Exception:
            pass
        return ""
#VEHICULO	VERSION DE DISEÑO	FORMULA	COLOR	PIEZA
#Z0409_	000_	LL40-8_	01_	001
    def _iniciar(self):
        if self._corriendo:
            return

        items = self._get_items() if self._get_items else []
        if not items:
            messagebox.showwarning(
                "Sin combinaciones",
                "No hay combinaciones activas para procesar.\n\n"
                "Ve a la pestana Bloqueos y asegurate de que existan\n"
                "combinaciones activas (no bloqueadas).",
                parent=self,
            )
            return

        # ── Filtrar por formula ANTES del dialogo ─────────────────────────────
        zfer_base    = items[0].zfer_origen if items else ""
        formula_base = self._obtener_formula_base(zfer_base)

        if not formula_base:
            # No se encontro en ninguna tabla — bloquear y alertar
            messagebox.showerror(
                "Formula no encontrada",
                f"No se encontro la formula del ZFER base '{zfer_base}' en ninguna "
                f"de las tablas de referencia:\n\n"
                f"  • ZFER_Characteristics_Genesis\n"
                f"  • TCAL_CALENDARIO_COLOMBIA_DIRECT\n\n"
                f"Verifica que el ZFER este registrado en la BD de produccion "
                f"antes de continuar.",
                parent=self,
            )
            return

        fb         = formula_base.strip().upper()
        items_sap  = [it for it in items if it.formula.strip().upper() == fb]
        items_solo = [it for it in items if it.formula.strip().upper() != fb]

        # ── Dialogo con numeros reales ────────────────────────────────────────
        lineas = []
        lineas.append(f"Formula base detectada: {formula_base}")
        lineas.append("")
        lineas.append(f"  Cambio de color (van a SAP) : {len(items_sap)}")
        if items_solo:
            lineas.append(f"  Solo reporte (otra formula)  : {len(items_solo)}")
        lineas.append("")
        lineas.append("SAP GUI debe estar abierto con sesion activa.")
        lineas.append("Continuar?")

        resp = messagebox.askyesno(
            "Confirmar inicio",
            "\n".join(lineas),
            parent=self,
        )
        if not resp:
            return

        # ── Limpiar estado UI ─────────────────────────────────────────────────
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")
        for k in self._cards:
            self._card_set(k, 0)
        self._card_set("total", len(items_sap))
        self._progreso["maximum"] = max(len(items_sap), 1)
        self._progreso["value"]   = 0
        self._lbl_prog.config(text=f"0 / {len(items_sap)}")
        self._corriendo = True
        self._btn_iniciar.set_estado(False)
        self._btn_reporte.set_estado(False)
        self._ultimo_reporte = ""

        threading.Thread(
            target=self._hilo_sap,
            args=(items_sap, items_solo, formula_base),
            daemon=True,
        ).start()

    def _hilo_sap(self, items_sap: list, items_solo: list, formula_base: str):
        """
        Corre en un hilo — recibe los items ya filtrados desde _iniciar.
        items_sap   : solo los de misma formula (van a SAP)
        items_solo  : fórmula diferente (solo reporte)
        formula_base: formula del ZFER base (ya detectada)
        """
        from SAP_AUTOMATIZADOR import AutomatizadorSAP

        def emit(texto, tag=""):
            self._queue.put(("log", texto, tag))

        def progreso(procesados, ok, errores):
            self._queue.put(("progreso", procesados, ok, errores))

        emit(f"  Batch iniciado — {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "info")
        emit(f"  Formula base  : {formula_base or '(no detectada)'}", "info")
        emit(f"  A procesar SAP: {len(items_sap)}", "ok")
        if items_solo:
            emit(f"  Solo reporte  : {len(items_solo)}", "warn")
        emit("  " + "─" * 52, "dim")

        auto = AutomatizadorSAP()
        auto.formula_base = formula_base

        try:
            if not auto.conectar():
                emit("  [ERROR] No se pudo conectar a SAP GUI.", "error")
                emit("  Asegurate de que SAP este abierto con sesion activa.", "warn")
                self._queue.put(("fin", False, ""))
                return
        except Exception as e:
            emit(f"  [ERROR] Conexion SAP: {e}", "error")
            self._queue.put(("fin", False, ""))
            return

        # ── MM02 del ZFER base: solo necesitamos la FRANJA (P_FRANJ) ─────────
        zfer_base = items_sap[0].zfer_origen if items_sap else (
            items_solo[0].zfer_origen if items_solo else ""
        )
        p_franj = "00"
        emit(f"  Leyendo MM02 ({zfer_base}) para franja...", "info")
        try:
            clasif  = auto.leer_clasificacion_zfer(zfer_base)
            p_franj = clasif.franja if clasif.franja else "00"
            emit(f"  Franja: {clasif.franja or 'Sin Franja'}  →  P_FRANJ={p_franj}", "dim")
            emit(f"  PARTNUMBER: {clasif.partnumber}", "dim")
        except Exception as e:
            emit(f"  [WARN] MM02 base: {e} — se asume P_FRANJ=00", "warn")

        # ── Poblar items_solo_reporte en el automatizador ─────────────────────
        auto.items_solo_reporte = [
            {
                "zfer_base":  it.zfer_origen,
                "formula":    it.formula,
                "acero":      it.acero,
                "color":      it.color,
                "cod_pieza":  getattr(it, "cod_pieza",  ""),
                "tipo_pieza": getattr(it, "tipo_pieza", ""),
                "motivo":     (f"Formula '{it.formula}' difiere de la base "
                               f"'{formula_base}' — requiere cambio de formula previo"),
            }
            for it in items_solo
        ]

        if not items_sap:
            emit("  Sin items para SAP — generando reporte de solo-reporte...", "warn")
            try:
                ruta = auto._generar_reporte()
                emit(f"  Reporte: {ruta}", "ok")
                self._queue.put(("fin", True, ruta))
            except Exception as e:
                emit(f"  [ERROR] Reporte: {e}", "error")
                self._queue.put(("fin", False, ""))
            return

        auto.batch_id   = auto.batch_id
        auto.resultados = []
        ok_n  = 0
        err_n = 0

        for idx, item in enumerate(items_sap, 1):
            emit(f"\n  [{idx}/{len(items_sap)}] {item.formula} / {item.acero} / {item.color[:35]}", "info")

            p_color = auto._extraer_numero_color(item.color)
            if not p_color:
                emit(f"    [WARN] No se pudo extraer número de color de: {item.color}", "warn")

            res = auto.procesar_combinacion(
                zfer_base  = item.zfer_origen,
                formula    = item.formula,
                acero      = item.acero,
                color      = item.color,
                cod_pieza  = getattr(item, "cod_pieza",  ""),
                tipo_pieza = getattr(item, "tipo_pieza", ""),
                p_color    = p_color,
                p_franj    = p_franj,
            )
            auto.resultados.append(res)
            auto._log_bd(res)
            auto._guardar_progreso_json()   # checkpoint JSON después de cada item

            if res.estado == "OK":
                ok_n += 1
                emit(f"    ZFER nuevo : {res.zfer_nuevo}  |  ZFOR: {res.zfor_nuevo}", "ok")
                emit(f"    Posiciones : {', '.join(res.posiciones_bom)}", "dim")
                emit(f"    Duración   : {res.duracion_seg}s", "ok")
            else:
                err_n += 1
                emit(f"    ERROR: {res.error}", "error")

            progreso(idx, ok_n, err_n)

        # Reporte final
        emit("\n  " + "─" * 52, "dim")
        if items_solo:
            emit(f"  {len(items_solo)} items omitidos (cambio formula pendiente) incluidos en reporte.", "warn")
        emit("  Generando reporte Excel...", "info")
        try:
            ruta = auto._generar_reporte()
            emit(f"  Reporte guardado: {ruta}", "ok")
            emit(f"  JSON progreso  : {auto._ruta_json}", "dim")
            self._queue.put(("fin", True, ruta))
        except Exception as e:
            emit(f"  [WARN] No se pudo generar reporte: {e}", "warn")
            self._queue.put(("fin", True, ""))

    # ── Poll de la queue (actualiza UI desde hilo) ────────────────────────────

    def _poll_queue(self):
        try:
            while True:
                msg = self._queue.get_nowait()
                tipo = msg[0]

                if tipo == "log":
                    _, texto, tag = msg
                    self._log_write(texto, tag)

                elif tipo == "progreso":
                    _, procesados, ok, errores = msg
                    total = int(self._progreso["maximum"])
                    self._progreso["value"] = procesados
                    self._lbl_prog.config(text=f"{procesados} / {total}")
                    self._card_set("ok",    ok)
                    self._card_set("error", errores)
                    self._card_set("activo", procesados)

                elif tipo == "total_real":
                    _, n_sap, n_solo = msg
                    self._card_set("total", n_sap)
                    # reusar card "activo" para mostrar los de solo reporte
                    if hasattr(self, "_lbl_solo_rep"):
                        self._lbl_solo_rep.config(text=f"Solo reporte: {n_solo}")

                elif tipo == "fin":
                    _, exito, ruta = msg
                    self._corriendo = False
                    self._btn_iniciar.set_estado(True)
                    self._card_set("activo", 0)
                    self._ultimo_reporte = ruta
                    if ruta:
                        self._btn_reporte.set_estado(True)

        except queue.Empty:
            pass
        self.after(150, self._poll_queue)

    def _abrir_reporte(self):
        if self._ultimo_reporte and os.path.exists(self._ultimo_reporte):
            os.startfile(self._ultimo_reporte)


# ══════════════════════════════════════════════════════════════════════════════
# APP PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

class AppModulo5(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("MODULO 5  —  AGP Glass  |  Generador de Combinaciones")
        self.geometry("1480x860")
        self.minsize(1100, 680)
        self.configure(bg=C["bg_header"])

        self._aplicar_estilos()
        self._build()

    def _aplicar_estilos(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("AGP.TNotebook",
            background=C["bg_header"],
            borderwidth=0,
            tabmargins=[0, 0, 0, 0],
        )
        style.configure("AGP.TNotebook.Tab",
            background=C["tab_bg"],
            foreground=C["txt_header"],
            font=FONT_BOLD,
            padding=[24, 10],
            borderwidth=0,
        )
        style.map("AGP.TNotebook.Tab",
            background=[("selected", C["bg_app"])],
            foreground=[("selected", C["txt_primary"])],
        )
        style.configure("M5.Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=C["txt_primary"],
            rowheight=38,
            font=("Segoe UI", 10),
            borderwidth=0,
        )
        style.configure("M5.Treeview.Heading",
            background=C["bg_header"],
            foreground=C["txt_white"],
            font=FONT_BOLD,
            relief="flat",
            padding=[10, 10],
        )
        style.map("M5.Treeview",
            background=[("selected", C["agp_blue_dark"])],
            foreground=[("selected", C["txt_white"])],
        )
        style.map("M5.Treeview.Heading",
            background=[("active", C["tbl_hdr_border"])],
        )
        style.configure("M5.Vertical.TScrollbar",
            background=C["border_light"],
            troughcolor=C["bg_app"],
            arrowcolor=C["txt_secondary"],
            borderwidth=0, relief="flat",
        )
        style.configure("M5.TCombobox",
            fieldbackground=C["bg_card"],
            background=C["bg_card"],
            foreground=C["txt_primary"],
            arrowcolor=C["txt_secondary"],
            borderwidth=0, relief="flat",
            padding=[6, 4],
        )
        style.map("M5.TCombobox",
            fieldbackground=[("readonly", C["bg_card"])],
            selectbackground=[("readonly", C["bg_card"])],
            selectforeground=[("readonly", C["txt_primary"])],
        )

    def _build(self):
        self._notebook = ttk.Notebook(self, style="AGP.TNotebook")
        self._notebook.pack(fill="both", expand=True)

        # Instanciar tabs
        self._tab_combinaciones = TabCombinaciones(
            self._notebook,
            on_combinaciones_listas=self._on_combinaciones_listas,
        )
        self._tab_bloqueos = TabBloqueos(
            self._notebook,
            on_listo_para_sap=self._on_listo_para_sap,
        )
        self._tab_sap = TabSAP(
            self._notebook,
            get_items_fn=self._tab_bloqueos.get_items_activos,
        )

        self._notebook.add(self._tab_combinaciones, text="  1   Combinaciones  ")
        self._notebook.add(self._tab_bloqueos,      text="  2   Bloqueos       ")
        self._notebook.add(self._tab_sap,           text="  3   SAP            ")

        # Deshabilitar tabs 2 y 3 hasta que estén listos
        # (no hay disable en ttk.Notebook nativo — usamos lógica en callbacks)

    def _on_combinaciones_listas(self):
        """Callback cuando COMBINADOR termina — ir a tab Bloqueos y cargar."""
        self._tab_bloqueos.cargar(RUTA_COMBINACIONES)
        self._notebook.select(1)

    def _on_listo_para_sap(self):
        """Callback cuando el tecnico confirma bloqueos — ir a tab SAP."""
        self._notebook.select(2)


# ── Punto de entrada ──────────────────────────────────────────────────────────

def main():
    app = AppModulo5()
    app.mainloop()


if __name__ == "__main__":
    main()
