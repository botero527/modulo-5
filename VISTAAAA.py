
import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
import pyodbc
import os
import getpass
from dataclasses import dataclass
from collections import defaultdict

try:
    from PIL import Image, ImageTk
    PIL_OK = True
except ImportError:
    PIL_OK = False


# ── Paleta AGP corporativa ────────────────────────────────────────────────────

C = {
    # Estructura
    "bg_header":       "#2B3A47",   # navy oscuro — barra superior/inferior
    "agp_blue":        "#7DBFD4",   # azul AGP (arco del logo)
    "agp_blue_dark":   "#4E9CB5",   # azul AGP hover
    "bg_app":          "#EFF3F7",   # fondo general
    "bg_card":         "#FFFFFF",   # tarjetas métricas
    "bg_filter":       "#FFFFFF",   # barra filtros
    "bg_row_odd":      "#F5F9FC",
    "bg_row_even":     "#FFFFFF",
    # Texto
    "txt_primary":     "#1A2634",
    "txt_secondary":   "#637282",
    "txt_white":       "#FFFFFF",
    "txt_header":      "#C8D8E4",   # texto suave en fondos navy
    # Estados
    "activa_txt":      "#1A7340",
    "activa_bg":       "#D6F5E3",
    "activa_border":   "#27AE60",
    "bloqueada_txt":   "#922B21",
    "bloqueada_bg":    "#FADBD8",
    "bloqueada_border":"#E74C3C",
    "pendiente_txt":   "#9A5B00",
    "pendiente_bg":    "#FEF3CD",
    "pendiente_border":"#F39C12",
    # Grupos padre
    "grupo_act_bg":    "#D6E4F0",
    "grupo_bloq_bg":   "#F5B7B1",
    "grupo_pend_bg":   "#FAE5A0",
    "grupo_mix_bg":    "#AED6F1",
    # Botones
    "btn_bloquear":    "#C0392B",
    "btn_bloquear_h":  "#A93226",
    "btn_reactivar":   "#1A7340",
    "btn_reactivar_h": "#155930",
    "btn_confirm":     "#7DBFD4",
    "btn_confirm_h":   "#4E9CB5",
    "btn_neutral":     "#7F8C8D",
    "btn_neutral_h":   "#626D6E",
    # Bordes
    "border_light":    "#D5E0EA",
    "border_dark":     "#2B3A47",
    "tbl_hdr_border":  "#3D5166",
}

LOGO_PATH = (
    r"C:\Users\abotero\OneDrive - AGP GROUP\Documentos\MODULO_5"
    r"\agp-america-s-a-logo-vector-removebg-preview (1).png"
)

FONT_TITLE   = ("Segoe UI", 16, "bold")
FONT_LABEL   = ("Segoe UI", 10)
FONT_BOLD    = ("Segoe UI", 10, "bold")
FONT_SMALL   = ("Segoe UI", 9)
FONT_MICRO   = ("Segoe UI", 8)
FONT_METRIC  = ("Segoe UI", 24, "bold")
FONT_META    = ("Segoe UI", 8)
FONT_CAPS    = ("Segoe UI", 8, "bold")


# ── Configuración BD ──────────────────────────────────────────────────────────

DB_LOCAL = {
    "server":   r"localhost\SQLEXPRESS",
    "database": "MODULO_5",
    "driver":   "ODBC Driver 17 for SQL Server",
}

RUTA_COMBINACIONES = (
    r"C:\Users\abotero\OneDrive - AGP GROUP"
    r"\Documentos\MODULO_5\combinaciones5.xlsx"
)


def get_conn_local() -> pyodbc.Connection:
    s = (
        f"DRIVER={{{DB_LOCAL['driver']}}};"
        f"SERVER={DB_LOCAL['server']};"
        f"DATABASE={DB_LOCAL['database']};"
        "Trusted_Connection=yes;"
    )
    return pyodbc.connect(s, autocommit=False)


# ── Estructuras de datos ──────────────────────────────────────────────────────

@dataclass
class Combinacion:
    formula:    str
    color:      str
    acero:      str
    zfer_origen:str
    mercado:    str
    cod_pieza:  str
    tipo_pieza: str


@dataclass
class ItemColor:
    """
    Una combinación individual: formula × acero × color = 1 ZFER en SAP.
    El bloqueo opera a este nivel.
    """
    formula:          str
    acero:            str
    color:            str
    zfer_origen:      str  = ""
    mercado:          str  = ""
    cod_pieza:        str  = ""
    tipo_pieza:       str  = ""
    bloqueado:        bool = False
    motivo:           str  = ""
    bloqueado_por:    str  = ""
    pendiente:        bool = False
    motivo_pendiente: str  = ""
    accion_pendiente: str  = ""   # "BLOQUEAR" | "REACTIVAR"

    @property
    def estado_label(self) -> str:
        if self.pendiente:
            accion = "BLOQUEAR" if self.accion_pendiente == "BLOQUEAR" else "REACTIVAR"
            return f"[ PENDIENTE — {accion} ]"
        return "BLOQUEADA" if self.bloqueado else "ACTIVA"

    @property
    def estado_tag(self) -> str:
        if self.pendiente:
            return "color_pendiente"
        return "color_bloqueada" if self.bloqueado else "color_activa"


# ── Lectura del Excel ─────────────────────────────────────────────────────────

def leer_combinaciones(ruta: str) -> list:
    if not os.path.exists(ruta):
        raise FileNotFoundError(
            f"Excel no encontrado:\n{ruta}\n\n"
            "Ejecuta primero COMBINADOR.py para generar las combinaciones."
        )
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    ws = wb.active
    resultado = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        try:
            _, zfer, mercado, cod_pieza, tipo_pieza, formula, color, acero, *_ = row
            resultado.append(Combinacion(
                formula     = str(formula    or "").strip(),
                color       = str(color      or "").strip(),
                acero       = str(acero      or "").strip(),
                zfer_origen = str(zfer       or "").strip(),
                mercado     = str(mercado    or "").strip(),
                cod_pieza   = str(cod_pieza  or "").strip(),
                tipo_pieza  = str(tipo_pieza or "").strip(),
            ))
        except ValueError:
            continue
    wb.close()
    return resultado


def agrupar(combinaciones: list) -> tuple:
    items  = {}
    grupos = defaultdict(list)
    for c in combinaciones:
        key = (c.formula, c.acero, c.color)
        if key not in items:
            it = ItemColor(
                formula     = c.formula,
                acero       = c.acero,
                color       = c.color,
                zfer_origen = c.zfer_origen,
                mercado     = c.mercado,
                cod_pieza   = c.cod_pieza,
                tipo_pieza  = c.tipo_pieza,
            )
            items[key] = it
            grupos[(c.formula, c.acero)].append(it)
    return items, dict(grupos)


# ── Base de datos ─────────────────────────────────────────────────────────────

def cargar_bloqueos(zfer: str) -> dict:
    try:
        conn   = get_conn_local()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT formula, acero_variante, color_codigo, motivo, bloqueado_por
            FROM   dbo.M5_Bloqueos
            WHERE  pedido_origen = ? AND activo = 1
        """, (zfer,))
        resultado = {}
        for formula, acero, color, motivo, por in cursor.fetchall():
            resultado[(formula, acero, color or "")] = {
                "motivo":        str(motivo or ""),
                "bloqueado_por": str(por    or ""),
            }
        cursor.close()
        conn.close()
        return resultado
    except pyodbc.Error as e:
        print(f"  [WARN] No se pudo conectar a BD local: {e}")
        return {}


def guardar_bloqueos_batch(
    zfer:       str,
    cod_pieza:  str,
    bloquear:   list,
    reactivar:  list,
) -> tuple:
    usuario = getpass.getuser()
    try:
        conn   = get_conn_local()
        cursor = conn.cursor()

        for formula, acero, color, motivo in bloquear:
            cursor.execute("""
                INSERT INTO dbo.M5_Bloqueos
                    (pedido_origen, tipo_pieza, formula, acero_variante, color_codigo,
                     motivo, bloqueado_por, activo)
                VALUES (?, ?, ?, ?, ?, ?, ?, 1)
            """, (zfer, cod_pieza or "N/A", formula, acero, color, motivo, usuario))

        for formula, acero, color in reactivar:
            cursor.execute("""
                UPDATE dbo.M5_Bloqueos
                SET    activo = 0
                WHERE  pedido_origen  = ?
                  AND  formula        = ?
                  AND  acero_variante = ?
                  AND  color_codigo   = ?
                  AND  activo         = 1
            """, (zfer, formula, acero, color))

        conn.commit()
        cursor.close()
        conn.close()
        return True, "OK"
    except pyodbc.Error as e:
        return False, str(e)


# ── Componentes UI ────────────────────────────────────────────────────────────

class Boton(tk.Label):
    """Botón moderno sin relieve, con efecto hover."""
    def __init__(self, parent, texto, comando, bg, hover,
                 fg="#FFFFFF", ancho=13, **kwargs):
        super().__init__(
            parent, text=texto, bg=bg, fg=fg,
            font=FONT_BOLD, cursor="hand2",
            padx=16, pady=12,
            width=ancho,
            relief="flat", **kwargs
        )
        self._bg    = bg
        self._hover = hover
        self._cmd   = comando
        self.bind("<Enter>",    lambda e: self.config(bg=self._hover))
        self.bind("<Leave>",    lambda e: self.config(bg=self._bg))
        self.bind("<Button-1>", lambda e: self._ejecutar())

    def _ejecutar(self):
        self.config(bg=self._hover)
        self._cmd()
        self.after(120, lambda: self.config(bg=self._bg) if self.winfo_exists() else None)


class TarjetaMetrica(tk.Frame):
    """Tarjeta de métrica con borde izquierdo de color."""
    def __init__(self, parent, titulo: str, color_acento: str, **kwargs):
        super().__init__(parent, bg=C["bg_card"], relief="flat", **kwargs)
        # Borde izquierdo coloreado
        tk.Frame(self, bg=color_acento, width=4).pack(side="left", fill="y")
        inner = tk.Frame(self, bg=C["bg_card"], padx=16, pady=10)
        inner.pack(side="left", fill="both", expand=True)
        tk.Label(inner, text=titulo.upper(), bg=C["bg_card"],
                 fg=C["txt_secondary"], font=FONT_CAPS).pack(anchor="w")
        self.lbl_valor = tk.Label(inner, text="—", bg=C["bg_card"],
                                  fg=color_acento, font=FONT_METRIC)
        self.lbl_valor.pack(anchor="w")

    def set(self, valor):
        self.lbl_valor.config(text=str(valor))


# ── Diálogo de bloqueo ────────────────────────────────────────────────────────

class DialogoBloqueo(tk.Toplevel):
    def __init__(self, parent, formula: str, acero: str, colores: list):
        super().__init__(parent)
        self.title("Bloquear combinacion")
        self.configure(bg=C["bg_app"])
        self.resizable(False, False)
        self.grab_set()
        self.motivo_resultado = None

        px = parent.winfo_rootx() + parent.winfo_width()  // 2
        py = parent.winfo_rooty() + parent.winfo_height() // 2
        self.geometry(f"520x490+{px-260}+{py-245}")
        self._build(formula, acero, colores)
        self.wait_window()

    def _build(self, formula, acero, colores):
        # Barra superior roja
        tk.Frame(self, bg=C["bloqueada_border"], height=4).pack(fill="x")

        # Encabezado
        hdr = tk.Frame(self, bg=C["bg_card"], padx=24, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="BLOQUEAR COMBINACION", bg=C["bg_card"],
                 fg=C["bloqueada_txt"], font=FONT_BOLD).pack(anchor="w")
        tk.Label(hdr, text="Esta accion queda registrada en la BD de control.",
                 bg=C["bg_card"], fg=C["txt_secondary"], font=FONT_SMALL).pack(anchor="w")

        tk.Frame(self, bg=C["border_light"], height=1).pack(fill="x")

        # Info combinacion
        info = tk.Frame(self, bg=C["bg_app"], padx=24, pady=12)
        info.pack(fill="x")
        for label, valor in [
            ("Formula",  formula),
            ("Acero",    acero if acero != "NO" else "Sin acero"),
            ("Colores",  f"{len(colores)} color(es) seleccionado(s)"),
        ]:
            fila = tk.Frame(info, bg=C["bg_app"])
            fila.pack(fill="x", pady=2)
            tk.Label(fila, text=f"{label}:", bg=C["bg_app"],
                     fg=C["txt_secondary"], font=FONT_SMALL, width=10, anchor="w").pack(side="left")
            tk.Label(fila, text=valor, bg=C["bg_app"],
                     fg=C["txt_primary"], font=FONT_BOLD, anchor="w").pack(side="left")

        sample = ", ".join(colores[:4])
        if len(colores) > 4:
            sample += f" (+{len(colores)-4} mas)"
        tk.Label(info, text=f"   {sample}", bg=C["bg_app"],
                 fg=C["txt_secondary"], font=FONT_SMALL,
                 anchor="w", wraplength=460).pack(fill="x", pady=(2, 0))

        tk.Frame(self, bg=C["border_light"], height=1).pack(fill="x")

        # Motivo
        cuerpo = tk.Frame(self, bg=C["bg_app"], padx=24, pady=12)
        cuerpo.pack(fill="x")
        tk.Label(cuerpo, text="Motivo del bloqueo (obligatorio):",
                 bg=C["bg_app"], fg=C["txt_primary"], font=FONT_BOLD).pack(anchor="w", pady=(0, 6))

        frame_txt = tk.Frame(cuerpo, bg=C["bloqueada_border"], padx=1, pady=1)
        frame_txt.pack(fill="x")
        self.txt_motivo = tk.Text(
            frame_txt, height=4, bg=C["bg_card"], fg=C["txt_primary"],
            font=FONT_LABEL, insertbackground=C["txt_primary"],
            relief="flat", padx=10, pady=8, wrap="word",
        )
        self.txt_motivo.pack(fill="x")
        self.txt_motivo.focus_set()

        self.lbl_error = tk.Label(cuerpo, text="", bg=C["bg_app"],
                                  fg=C["bloqueada_txt"], font=FONT_SMALL)
        self.lbl_error.pack(anchor="w", pady=(4, 0))

        # Botones
        tk.Frame(self, bg=C["border_light"], height=1).pack(fill="x")
        btns = tk.Frame(self, bg=C["bg_card"], padx=24, pady=12)
        btns.pack(fill="x")
        Boton(btns, "Cancelar", self._cancelar,
              C["btn_neutral"], C["btn_neutral_h"], ancho=10).pack(side="right", padx=(6, 0))
        Boton(btns, "Confirmar bloqueo", self._confirmar,
              C["btn_bloquear"], C["btn_bloquear_h"], ancho=18).pack(side="right")

        self.bind("<Return>", lambda e: self._confirmar())
        self.bind("<Escape>", lambda e: self._cancelar())

    def _confirmar(self):
        motivo = self.txt_motivo.get("1.0", "end").strip()
        if not motivo:
            self.lbl_error.config(text="  El motivo es obligatorio.")
            self.txt_motivo.focus_set()
            return
        self.motivo_resultado = motivo
        self.destroy()

    def _cancelar(self):
        self.motivo_resultado = None
        self.destroy()


# ── Ventana principal ─────────────────────────────────────────────────────────

class VistaPreviaBloqueos(tk.Frame):

    def __init__(self, parent, archivo_excel: str, on_close=None):
        super().__init__(parent, bg=C["bg_app"])
        self.archivo_excel    = archivo_excel
        self._on_close        = on_close
        self.zfer_origen      = ""
        self.cod_pieza        = ""
        self.mercado          = ""
        self.items: dict      = {}
        self.grupos_display: dict = {}
        self._logo_img        = None   # referencia PIL para evitar GC

        self._filtro_formula = tk.StringVar()
        self._filtro_acero   = tk.StringVar(value="Todos")
        self._filtro_estado  = tk.StringVar(value="Todos")
        self._sort_col       = None
        self._sort_asc       = True
        self._todos_expandidos = True

        self._aplicar_estilos()
        self._cargar_datos()
        self._build_header()
        self._build_metricas()
        self._build_filtros()
        self._build_footer()   # primero footer (side=bottom), luego tabla (expand=True)
        self._build_tabla()
        self._sincronizar_bloqueos_bd()
        self._refrescar_tabla()

        self.bind("<F5>",     lambda e: self._sincronizar_bloqueos_bd())
        self.bind("<Escape>", lambda e: self._cerrar())

    # ── Estilos ttk ──────────────────────────────────────────────────────────

    def _aplicar_estilos(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("M5.Treeview",
            background=C["bg_row_even"],
            fieldbackground=C["bg_row_even"],
            foreground=C["txt_primary"],
            rowheight=38,
            font=FONT_LABEL,
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

    # ── Header ────────────────────────────────────────────────────────────────

    def _build_header(self):
        outer = tk.Frame(self, bg=C["bg_header"])
        outer.pack(fill="x")

        # Franja AGP azul en la parte superior
        tk.Frame(outer, bg=C["agp_blue"], height=4).pack(fill="x")

        inner = tk.Frame(outer, bg=C["bg_header"], padx=28, pady=16)
        inner.pack(fill="x")

        # ─ Logo ──────────────────────────────────────────────────────────────
        logo_frame = tk.Frame(inner, bg=C["bg_header"])
        logo_frame.pack(side="left")

        logo_cargado = False
        if PIL_OK and os.path.exists(LOGO_PATH):
            try:
                img = Image.open(LOGO_PATH).convert("RGBA")
                # Componer sobre fondo navy para eliminar transparencia
                fondo = Image.new("RGBA", img.size, color=(43, 58, 71, 255))
                fondo.paste(img, mask=img.split()[3])
                img = fondo.convert("RGB")
                # Redimensionar proporcionalmente a alto=72
                w, h  = img.size
                nuevo_h = 72
                nuevo_w = int(w * nuevo_h / h)
                img = img.resize((nuevo_w, nuevo_h), Image.LANCZOS)
                self._logo_img = ImageTk.PhotoImage(img)
                tk.Label(logo_frame, image=self._logo_img,
                         bg=C["bg_header"]).pack(padx=(0, 24))
                logo_cargado = True
            except Exception:
                pass

        if not logo_cargado:
            tk.Label(logo_frame, text="AGP", bg=C["bg_header"],
                     fg=C["agp_blue"], font=("Segoe UI", 28, "bold"),
                     padx=12).pack()

        # Separador vertical
        tk.Frame(inner, bg=C["tbl_hdr_border"], width=1).pack(
            side="left", fill="y", padx=(0, 24), pady=4)

        # ─ Titulo y subtitulo ─────────────────────────────────────────────────
        titulo_frame = tk.Frame(inner, bg=C["bg_header"])
        titulo_frame.pack(side="left", fill="y")

        tk.Label(titulo_frame,
                 text="AGP GLASS  —  MODULO 5",
                 bg=C["bg_header"], fg=C["agp_blue"],
                 font=FONT_CAPS).pack(anchor="w")
        tk.Label(titulo_frame,
                 text="Vista Previa de Combinaciones",
                 bg=C["bg_header"], fg=C["txt_white"],
                 font=FONT_TITLE).pack(anchor="w")
        tk.Label(titulo_frame,
                 text=(f"ZFER: {self.zfer_origen}   |   "
                       f"Pieza: {self.cod_pieza}   |   "
                       f"Mercado: {self.mercado}"),
                 bg=C["bg_header"], fg=C["txt_header"],
                 font=FONT_META).pack(anchor="w", pady=(3, 0))

        # ─ Badges derecha ─────────────────────────────────────────────────────
        right = tk.Frame(inner, bg=C["bg_header"])
        right.pack(side="right", anchor="ne")

        self.lbl_pendientes = tk.Label(right, text="",
                                       bg=C["bg_header"],
                                       fg=C["pendiente_border"],
                                       font=FONT_BOLD)
        self.lbl_pendientes.pack(anchor="e")

        badge = tk.Frame(right, bg=C["agp_blue_dark"],
                         padx=12, pady=4)
        badge.pack(anchor="e", pady=(6, 0))
        tk.Label(badge, text="Motor Combinaciones AGP",
                 bg=C["agp_blue_dark"], fg=C["txt_white"],
                 font=FONT_MICRO).pack()

        # Franja azul inferior
        tk.Frame(outer, bg=C["tbl_hdr_border"], height=1).pack(fill="x")

    # ── Metricas ──────────────────────────────────────────────────────────────

    def _build_metricas(self):
        outer = tk.Frame(self, bg=C["bg_app"])
        outer.pack(fill="x", padx=20, pady=(16, 4))

        metricas = [
            ("Combinaciones", C["agp_blue_dark"]),
            ("Activas",       C["activa_border"]),
            ("Bloqueadas",    C["bloqueada_border"]),
            ("Pendientes",    C["pendiente_border"]),
            ("Colores unicos",C["agp_blue"]),
        ]
        self._metric_cards = {}
        keys = ["total", "activas", "bloqueadas", "pendientes", "colores"]

        for (titulo, color_acento), key in zip(metricas, keys):
            card = TarjetaMetrica(outer, titulo, color_acento,
                                  highlightthickness=1,
                                  highlightbackground=C["border_light"])
            card.pack(side="left", padx=(0, 10), ipadx=0, ipady=0)
            self._metric_cards[key] = card

        # Panel informativo a la derecha
        info = tk.Frame(outer, bg=C["bg_app"])
        info.pack(side="left", padx=(10, 0))
        for linea in [
            "Generadas por COMBINADOR.py a partir del ZFER base.",
            "Bloqueo granular: formula x acero x color = 1 ZFER en SAP.",
            "Seleccionar fila padre bloquea/reactiva todos sus colores.",
        ]:
            tk.Label(info, text=f"  {linea}",
                     bg=C["bg_app"], fg=C["txt_secondary"],
                     font=FONT_MICRO, anchor="w").pack(fill="x")

    def _actualizar_metricas(self):
        total   = len(self.items)
        bloq    = sum(1 for it in self.items.values() if it.bloqueado and not it.pendiente)
        pend    = sum(1 for it in self.items.values() if it.pendiente)
        activas = total - bloq - pend
        colores = len({it.color for it in self.items.values()})

        self._metric_cards["total"].set(total)
        self._metric_cards["activas"].set(activas)
        self._metric_cards["bloqueadas"].set(bloq)
        self._metric_cards["pendientes"].set(pend)
        self._metric_cards["colores"].set(colores)

        if pend > 0:
            self.lbl_pendientes.config(
                text=f"  {pend} cambio(s) pendiente(s) — sin guardar en BD"
            )
        else:
            self.lbl_pendientes.config(text="")

    # ── Filtros ───────────────────────────────────────────────────────────────

    def _build_filtros(self):
        # Tarjeta blanca con borde
        outer = tk.Frame(self, bg=C["bg_app"], padx=20, pady=4)
        outer.pack(fill="x")

        card = tk.Frame(outer, bg=C["bg_filter"],
                        highlightthickness=1,
                        highlightbackground=C["border_light"],
                        padx=16, pady=10)
        card.pack(fill="x")

        # Franja AGP azul izquierda
        tk.Frame(card, bg=C["agp_blue"], width=3).pack(side="left", fill="y", padx=(0, 14))

        # Etiqueta FILTROS
        tk.Label(card, text="FILTROS", bg=C["bg_filter"],
                 fg=C["txt_secondary"], font=FONT_CAPS).pack(side="left", padx=(0, 16))

        # Formula
        tk.Label(card, text="Formula:", bg=C["bg_filter"],
                 fg=C["txt_secondary"], font=FONT_SMALL).pack(side="left")
        entry_f = tk.Entry(card, textvariable=self._filtro_formula,
                           bg=C["bg_card"], fg=C["txt_primary"],
                           insertbackground=C["txt_primary"],
                           font=FONT_LABEL, width=12, relief="flat",
                           highlightthickness=1,
                           highlightbackground=C["border_light"])
        entry_f.pack(side="left", padx=(4, 18))
        entry_f.bind("<KeyRelease>", lambda e: self._refrescar_tabla())

        # Acero
        tk.Label(card, text="Acero:", bg=C["bg_filter"],
                 fg=C["txt_secondary"], font=FONT_SMALL).pack(side="left")
        cb_a = ttk.Combobox(card, textvariable=self._filtro_acero,
                            values=["Todos", "NO", "SN", "SP"],
                            state="readonly", width=8, style="M5.TCombobox")
        cb_a.pack(side="left", padx=(4, 18))
        cb_a.bind("<<ComboboxSelected>>", lambda e: self._refrescar_tabla())

        # Estado
        tk.Label(card, text="Estado:", bg=C["bg_filter"],
                 fg=C["txt_secondary"], font=FONT_SMALL).pack(side="left")
        cb_e = ttk.Combobox(card, textvariable=self._filtro_estado,
                            values=["Todos", "ACTIVA", "BLOQUEADA", "PENDIENTE"],
                            state="readonly", width=12, style="M5.TCombobox")
        cb_e.pack(side="left", padx=(4, 18))
        cb_e.bind("<<ComboboxSelected>>", lambda e: self._refrescar_tabla())

        Boton(card, "Limpiar", self._limpiar_filtros,
              C["btn_neutral"], C["btn_neutral_h"],
              fg=C["txt_white"], ancho=8).pack(side="left", padx=(0, 20))

        # Separador visual
        tk.Frame(card, bg=C["border_light"], width=1).pack(side="left", fill="y", padx=8)

        # Acciones
        Boton(card, "Bloquear", self._accion_bloquear,
              C["btn_bloquear"], C["btn_bloquear_h"],
              fg=C["txt_white"], ancho=11).pack(side="left", padx=4)
        Boton(card, "Reactivar", self._accion_reactivar,
              C["btn_reactivar"], C["btn_reactivar_h"],
              fg=C["txt_white"], ancho=11).pack(side="left", padx=4)

        tk.Label(card, text="seleccion", bg=C["bg_filter"],
                 fg=C["txt_secondary"], font=FONT_SMALL).pack(side="left", padx=(2, 12))

        tk.Frame(card, bg=C["border_light"], width=1).pack(side="left", fill="y", padx=8)

        self.btn_expand = Boton(card, "Colapsar todo", self._toggle_expand,
                                C["agp_blue_dark"], C["agp_blue"],
                                fg=C["txt_white"], ancho=13)
        self.btn_expand.pack(side="left", padx=4)

    def _limpiar_filtros(self):
        self._filtro_formula.set("")
        self._filtro_acero.set("Todos")
        self._filtro_estado.set("Todos")
        self._refrescar_tabla()

    # ── Tabla ─────────────────────────────────────────────────────────────────

    def _build_tabla(self):
        outer = tk.Frame(self, bg=C["bg_app"])
        outer.pack(fill="both", expand=True, padx=20, pady=(8, 0))

        # Marco con borde para la tabla
        frame = tk.Frame(outer, bg=C["bg_header"],
                         highlightthickness=1,
                         highlightbackground=C["border_light"])
        frame.pack(fill="both", expand=True)

        cols = ("zfer_origen", "formula", "acero", "color", "estado", "motivo", "bloqueado_por")
        self.tree = ttk.Treeview(frame, columns=cols, show="tree headings",
                                 style="M5.Treeview", selectmode="extended")

        self.tree.column("#0", width=22, minwidth=22, stretch=False)
        self.tree.heading("#0", text="")

        defs = [
            ("zfer_origen",   "ZFER Base",         110, "center"),
            ("formula",       "Formula",            115, "center"),
            ("acero",         "Acero",               90, "center"),
            ("color",         "Color / Codigo",     320, "w"),
            ("estado",        "Estado",             160, "center"),
            ("motivo",        "Motivo de bloqueo",  240, "w"),
            ("bloqueado_por", "Bloqueado por",      140, "center"),
        ]
        for col, heading, width, anchor in defs:
            self.tree.heading(col, text=heading,
                              command=lambda c=col: self._ordenar(c))
            self.tree.column(col, width=width, anchor=anchor, minwidth=60)

        # Tags filas padre (grupo formula x acero)
        self.tree.tag_configure("grupo_activa",
            background=C["grupo_act_bg"],
            foreground=C["txt_primary"],
            font=FONT_BOLD)
        self.tree.tag_configure("grupo_bloqueada",
            background=C["grupo_bloq_bg"],
            foreground=C["bloqueada_txt"],
            font=FONT_BOLD)
        self.tree.tag_configure("grupo_pendiente",
            background=C["grupo_pend_bg"],
            foreground=C["pendiente_txt"],
            font=FONT_BOLD)
        self.tree.tag_configure("grupo_mixto",
            background=C["grupo_mix_bg"],
            foreground=C["txt_primary"],
            font=FONT_BOLD)

        # Tags filas hijo (color individual)
        self.tree.tag_configure("color_activa",
            background=C["bg_row_odd"],
            foreground=C["txt_primary"])
        self.tree.tag_configure("color_activa_alt",
            background=C["bg_row_even"],
            foreground=C["txt_primary"])
        self.tree.tag_configure("color_bloqueada",
            background=C["bloqueada_bg"],
            foreground=C["bloqueada_txt"])
        self.tree.tag_configure("color_pendiente",
            background=C["pendiente_bg"],
            foreground=C["pendiente_txt"])

        sb_v = ttk.Scrollbar(frame, orient="vertical",
                              command=self.tree.yview,
                              style="M5.Vertical.TScrollbar")
        sb_h = ttk.Scrollbar(frame, orient="horizontal",
                              command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb_v.set, xscrollcommand=sb_h.set)
        sb_v.pack(side="right", fill="y")
        sb_h.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Double-Button-1>", self._doble_clic)

        # Etiqueta de conteo debajo de la tabla
        self.lbl_conteo = tk.Label(self, text="",
                                   bg=C["bg_app"], fg=C["txt_secondary"],
                                   font=FONT_MICRO)
        self.lbl_conteo.pack(side="bottom", anchor="e", padx=24, pady=(2, 4))

    def _grupos_filtrados(self) -> list:
        filtro_f = self._filtro_formula.get().strip().upper()
        filtro_a = self._filtro_acero.get()
        filtro_e = self._filtro_estado.get()

        resultado = []
        for (formula, acero), items_lista in self.grupos_display.items():
            if filtro_f and filtro_f not in formula.upper():
                continue
            if filtro_a != "Todos" and acero != filtro_a:
                continue

            if filtro_e == "ACTIVA":
                items_f = [it for it in items_lista if not it.bloqueado and not it.pendiente]
            elif filtro_e == "BLOQUEADA":
                items_f = [it for it in items_lista if it.bloqueado and not it.pendiente]
            elif filtro_e == "PENDIENTE":
                items_f = [it for it in items_lista if it.pendiente]
            else:
                items_f = items_lista

            if items_f:
                resultado.append(((formula, acero), items_f))

        if self._sort_col == "formula":
            resultado.sort(key=lambda x: x[0][0], reverse=not self._sort_asc)
        elif self._sort_col == "acero":
            resultado.sort(key=lambda x: x[0][1], reverse=not self._sort_asc)
        elif self._sort_col == "estado":
            resultado.sort(
                key=lambda x: sum(1 for it in x[1] if it.bloqueado),
                reverse=not self._sort_asc,
            )
        return resultado

    def _refrescar_tabla(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)

        grupos = self._grupos_filtrados()

        ACERO_LABEL = {"NO": "Sin acero", "SN": "Acero SN", "SP": "Acero SP"}

        for (formula, acero), items_lista in grupos:
            n      = len(items_lista)
            n_bloq = sum(1 for it in items_lista if it.bloqueado and not it.pendiente)
            n_pend = sum(1 for it in items_lista if it.pendiente)
            n_act  = n - n_bloq - n_pend

            acero_txt = ACERO_LABEL.get(acero, acero)

            parts = [f"{n_act} activa{'s' if n_act != 1 else ''}"]
            if n_bloq:
                parts.append(f"{n_bloq} bloqueada{'s' if n_bloq != 1 else ''}")
            if n_pend:
                parts.append(f"{n_pend} pendiente{'s' if n_pend != 1 else ''}")
            resumen = "  (" + "  /  ".join(parts) + f"  —  {n} total)"

            if n_pend > 0:
                group_tag = "grupo_pendiente"
            elif n_bloq == n:
                group_tag = "grupo_bloqueada"
            elif n_bloq > 0:
                group_tag = "grupo_mixto"
            else:
                group_tag = "grupo_activa"

            parent_iid = f"G__{formula}__{acero}"
            zfer_grupo = items_lista[0].zfer_origen if items_lista else ""
            self.tree.insert("", "end",
                iid    = parent_iid,
                text   = "",
                open   = self._todos_expandidos,
                tags   = (group_tag,),
                values = (zfer_grupo, formula, acero_txt, resumen, "", "", ""),
            )

            for ci, item in enumerate(items_lista):
                if item.pendiente:
                    ctag = "color_pendiente"
                elif item.bloqueado:
                    ctag = "color_bloqueada"
                else:
                    ctag = "color_activa" if ci % 2 == 0 else "color_activa_alt"

                motivo_txt = (item.motivo_pendiente if item.pendiente else item.motivo)
                motivo_txt = (motivo_txt[:58] + "...") if len(motivo_txt) > 60 else motivo_txt
                motivo_txt = motivo_txt or "—"

                self.tree.insert(parent_iid, "end",
                    iid    = f"C__{formula}__{acero}__{ci}",
                    text   = "",
                    tags   = (ctag,),
                    values = (
                        item.zfer_origen,
                        "",
                        "",
                        item.color,
                        item.estado_label,
                        motivo_txt,
                        item.bloqueado_por or "—",
                    ),
                )

        self._actualizar_metricas()
        n_grupos = len(grupos)
        total_g  = len(self.grupos_display)
        total_it = len(self.items)
        self.lbl_conteo.config(
            text=(f"Mostrando {n_grupos} de {total_g} grupos  |  "
                  f"{total_it} combinaciones totales  |  "
                  f"Doble clic en color para bloquear/reactivar  |  F5 para sincronizar BD")
        )

    def _ordenar(self, col: str):
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        self._refrescar_tabla()

    # ── Footer ────────────────────────────────────────────────────────────────

    def _build_footer(self):
        # side="bottom" garantiza que el footer siempre sea visible
        # aunque la tabla tenga expand=True
        tk.Frame(self, bg=C["agp_blue"], height=4).pack(side="bottom", fill="x")
        tk.Frame(self, bg=C["tbl_hdr_border"], height=1).pack(side="bottom", fill="x")

        footer = tk.Frame(self, bg=C["bg_header"], padx=20, pady=14)
        footer.pack(side="bottom", fill="x")

        # Texto izquierda
        left = tk.Frame(footer, bg=C["bg_header"])
        left.pack(side="left")
        for linea in [
            "Los bloqueos se persisten en BD local al confirmar — no fila por fila.",
            "Un bloqueo guardado no se puede eliminar, solo reactivar (trazabilidad auditada).",
        ]:
            tk.Label(left, text=linea, bg=C["bg_header"],
                     fg=C["txt_header"], font=FONT_MICRO).pack(anchor="w")

        # Botones derecha
        right = tk.Frame(footer, bg=C["bg_header"])
        right.pack(side="right")

        Boton(right, "Cerrar", self._cerrar,
              C["btn_neutral"], C["btn_neutral_h"],
              fg=C["txt_white"], ancho=9).pack(side="left", padx=(0, 6))

        Boton(right, "Sincronizar BD", self._sincronizar_bloqueos_bd,
              C["agp_blue_dark"], C["agp_blue"],
              fg=C["txt_white"], ancho=16).pack(side="left", padx=(0, 6))

        self.btn_confirmar = Boton(right, "Confirmar y guardar",
                                   self._confirmar_y_guardar,
                                   C["btn_confirm"], C["btn_confirm_h"],
                                   fg=C["txt_white"], ancho=20)
        self.btn_confirmar.pack(side="left")


    # ── Logica de acciones ────────────────────────────────────────────────────

    def _toggle_expand(self):
        self._todos_expandidos = not self._todos_expandidos
        for iid in self.tree.get_children():
            self.tree.item(iid, open=self._todos_expandidos)
        self.btn_expand.config(
            text="Expandir todo" if not self._todos_expandidos else "Colapsar todo"
        )

    def _resolver_items(self, iids) -> list:
        vistos = {}
        for iid in iids:
            if iid.startswith("G__"):
                parts = iid.split("__", 2)
                if len(parts) == 3:
                    formula, acero = parts[1], parts[2]
                    for it in (self.grupos_display.get((formula, acero)) or []):
                        key = (it.formula, it.acero, it.color)
                        if key not in vistos:
                            vistos[key] = it
            elif iid.startswith("C__"):
                parent_iid = self.tree.parent(iid)
                if parent_iid and parent_iid.startswith("G__"):
                    g_parts = parent_iid.split("__", 2)
                    if len(g_parts) == 3:
                        formula, acero = g_parts[1], g_parts[2]
                        vals  = self.tree.item(iid, "values")
                        color = vals[3] if len(vals) > 3 else ""  # idx 3: zfer, formula, acero, COLOR
                        key   = (formula, acero, color)
                        it    = self.items.get(key)
                        if it and key not in vistos:
                            vistos[key] = it
        return list(vistos.values())

    def _doble_clic(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid or iid.startswith("G__"):
            return
        self.tree.selection_set(iid)
        items = self._resolver_items([iid])
        if not items:
            return
        it = items[0]
        if it.bloqueado or (it.pendiente and it.accion_pendiente == "BLOQUEAR"):
            self._accion_reactivar()
        else:
            self._accion_bloquear()

    def _accion_bloquear(self):
        iids = self.tree.selection()
        if not iids:
            messagebox.showwarning("Sin seleccion",
                                   "Selecciona al menos un color para bloquear.",
                                   parent=self)
            return
        bloqueables = [
            it for it in self._resolver_items(iids)
            if not it.bloqueado and not (it.pendiente and it.accion_pendiente == "BLOQUEAR")
        ]
        if not bloqueables:
            messagebox.showinfo("Info",
                                "Las filas seleccionadas ya estan bloqueadas o pendientes.",
                                parent=self)
            return
        primer = bloqueables[0]
        dlg = DialogoBloqueo(self, primer.formula, primer.acero,
                             [it.color for it in bloqueables])
        if dlg.motivo_resultado is None:
            return
        motivo = dlg.motivo_resultado
        for it in bloqueables:
            it.pendiente        = True
            it.accion_pendiente = "BLOQUEAR"
            it.motivo_pendiente = motivo
        self._refrescar_tabla()

    def _accion_reactivar(self):
        iids = self.tree.selection()
        if not iids:
            messagebox.showwarning("Sin seleccion",
                                   "Selecciona al menos un color para reactivar.",
                                   parent=self)
            return
        reactivables = [
            it for it in self._resolver_items(iids)
            if it.bloqueado or (it.pendiente and it.accion_pendiente == "BLOQUEAR")
        ]
        if not reactivables:
            messagebox.showinfo("Info", "Las filas seleccionadas ya estan activas.",
                                parent=self)
            return
        nombres = "\n".join(
            f"  - {it.formula} / {it.acero} / {it.color[:35]}"
            for it in reactivables[:5]
        )
        if len(reactivables) > 5:
            nombres += f"\n  ... y {len(reactivables)-5} mas"
        resp = messagebox.askyesno("Reactivar combinaciones",
                                   f"Reactivar las siguientes combinaciones?\n\n{nombres}",
                                   parent=self)
        if not resp:
            return
        for it in reactivables:
            if it.pendiente and it.accion_pendiente == "BLOQUEAR":
                it.pendiente        = False
                it.accion_pendiente = ""
                it.motivo_pendiente = ""
            else:
                it.pendiente        = True
                it.accion_pendiente = "REACTIVAR"
                it.motivo_pendiente = ""
        self._refrescar_tabla()

    def _confirmar_y_guardar(self):
        pendientes = [it for it in self.items.values() if it.pendiente]
        if not pendientes:
            messagebox.showinfo("Sin cambios",
                                "No hay cambios pendientes para guardar.\n\n"
                                "Usa Bloquear / Reactivar para marcar filas primero.",
                                parent=self)
            return

        a_bloquear  = [it for it in pendientes if it.accion_pendiente == "BLOQUEAR"]
        a_reactivar = [it for it in pendientes if it.accion_pendiente == "REACTIVAR"]

        resumen = "Se guardaran los siguientes cambios en la BD:\n\n"
        if a_bloquear:
            resumen += f"  Bloquear   {len(a_bloquear)} combinacion(es)\n"
        if a_reactivar:
            resumen += f"  Reactivar  {len(a_reactivar)} combinacion(es)\n"
        resumen += "\nConfirmar? Esta accion no se puede deshacer."

        resp = messagebox.askyesno("Confirmar cambios", resumen, parent=self)
        if not resp:
            return

        bloquear_list  = [(it.formula, it.acero, it.color, it.motivo_pendiente) for it in a_bloquear]
        reactivar_list = [(it.formula, it.acero, it.color) for it in a_reactivar]

        exito, msg = guardar_bloqueos_batch(
            zfer=self.zfer_origen, cod_pieza=self.cod_pieza,
            bloquear=bloquear_list, reactivar=reactivar_list,
        )
        if not exito:
            messagebox.showerror("Error al guardar",
                                 f"No se pudo guardar en la BD:\n\n{msg}\n\n"
                                 "Los cambios NO fueron aplicados.",
                                 parent=self)
            return

        usuario = getpass.getuser()
        for it in a_bloquear:
            it.bloqueado = True; it.motivo = it.motivo_pendiente
            it.bloqueado_por = usuario; it.pendiente = False
            it.accion_pendiente = ""; it.motivo_pendiente = ""

        for it in a_reactivar:
            it.bloqueado = False; it.motivo = ""
            it.bloqueado_por = ""; it.pendiente = False
            it.accion_pendiente = ""; it.motivo_pendiente = ""

        self._refrescar_tabla()
        messagebox.showinfo("Guardado exitoso",
                            f"Cambios guardados correctamente.\n\n"
                            f"  Bloqueadas : {len(a_bloquear)}\n"
                            f"  Reactivadas: {len(a_reactivar)}",
                            parent=self)

    def _sincronizar_bloqueos_bd(self):
        bloqueos = cargar_bloqueos(self.zfer_origen)
        for (formula, acero, color), datos in bloqueos.items():
            key = (formula, acero, color)
            if key in self.items:
                it = self.items[key]
                if not it.pendiente:
                    it.bloqueado     = True
                    it.motivo        = datos["motivo"]
                    it.bloqueado_por = datos["bloqueado_por"]
        self._refrescar_tabla()

    def _cerrar(self):
        pendientes = [it for it in self.items.values() if it.pendiente]
        if pendientes:
            resp = messagebox.askyesno(
                "Cambios sin guardar",
                f"Hay {len(pendientes)} cambio(s) pendiente(s) sin guardar.\n\n"
                "Cerrar de todas formas? Los cambios se perderan.",
                parent=self
            )
            if not resp:
                return
        if self._on_close:
            self._on_close()
        else:
            self.winfo_toplevel().destroy()

    # ── Carga de datos ────────────────────────────────────────────────────────

    def _cargar_datos(self):
        self.combinaciones              = leer_combinaciones(self.archivo_excel)
        self.items, self.grupos_display = agrupar(self.combinaciones)
        if self.combinaciones:
            c = self.combinaciones[0]
            self.zfer_origen = c.zfer_origen
            self.cod_pieza   = c.cod_pieza
            self.mercado     = c.mercado

    def ejecutar(self):
        """Para uso standalone — configura protocolo de cierre y arranca mainloop."""
        top = self.winfo_toplevel()
        top.protocol("WM_DELETE_WINDOW", self._cerrar)
        self.pack(fill="both", expand=True)
        top.mainloop()


# ── Punto de entrada ──────────────────────────────────────────────────────────

def main():
    print("  MODULO 5 — Vista Previa con Bloqueos")

    if not os.path.exists(RUTA_COMBINACIONES):
        print(f"\n  [ERROR] No encontrado: {RUTA_COMBINACIONES}")
        print("  Ejecuta COMBINADOR.py primero para generar el Excel.")
        input("  Presiona Enter para salir...")
        return

    root = tk.Tk()
    root.title("MODULO 5  —  Vista Previa de Combinaciones  |  AGP Glass")
    root.geometry("1440x820")
    root.minsize(1100, 640)
    app = VistaPreviaBloqueos(root, RUTA_COMBINACIONES)
    app.ejecutar()


if __name__ == "__main__":
    main()
