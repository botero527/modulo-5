"""
sap_auto.py — Automatizador SAP MODULO 5 AGP Glass (adaptado para Flask)
Flujo por combinación:
  1. ZPPP0042 → validar que ZFER exista y utilización = "1"
  2. ZMME0001 → Homologar → Cambio de Color → Ejecutar → ZFER_NUEVO
  3. ZPPR0020 → sesión auxiliar → polling fases (> 7 con S)
  4. ZMME0001 → cambiar material a ZFER_NUEVO → Comp BOM → llenar tabla → COPY_ITEM
  5. MM02 → actualizar PARTNUMBER del ZFER_NUEVO (y ZFOR_NUEVO si existe)

Parámetros que vienen de la BD (ya resueltos por app.py):
  - franja    : shade_band del ZFER base (ej: "00", "01")
  - pn_base   : Z_AGP_PARTNUMBER del ZFER base
  - zpla      : ZPLA sugerido de la matriz de combinaciones
  - color_codigo: código del color SAP (ej: "19")
"""

import os
import re
import win32com.client
import time
import pyodbc
import datetime
import uuid
from dataclasses import dataclass, field
from typing import Optional

# ── Tiempos de espera (máximos por categoría) ────────────────────────────────
T_RAPIDO = 0.8   # máx para clicks / campos simples
T_MEDIO  = 2.5   # máx para navegación entre pantallas
T_LENTO  = 8.0   # máx para ejecutar transacciones pesadas

# Mínimos garantizados antes de empezar a hacer poll
_T_MIN_RAPIDO = 0.05
_T_MIN_MEDIO  = 0.10
_T_MIN_LENTO  = 0.20

_SAP_USER = os.environ.get("SAP_USER", "FESPITIA") #PROGRAING

# ── BD Local ──────────────────────────────────────────────────────────────────
_DB_LOCAL_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    r"SERVER=localhost\SQLEXPRESS;"
    "DATABASE=MODULO_5;"
    "Trusted_Connection=yes;"
)

# ── BD SAP (Azure) — solo lectura, para buscar planos ─────────────────────────
_DB_SAP_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=agpcolsap.database.windows.net;"
    "DATABASE=DB_COL_SAP;"
    "UID=Viewer;"
    "PWD=AgpconsCol2023;"
    "Encrypt=yes;TrustServerCertificate=no;Connection Timeout=20;"
)


# ── Estructura de resultado ───────────────────────────────────────────────────

@dataclass
class ResultadoItem:
    batch_id:       str   = ""
    zfer_base:      str   = ""
    color_codigo:   str   = ""
    zfer_nuevo:     str   = ""
    zfor_nuevo:     str   = ""
    zpla:           str   = ""
    posiciones_bom: list  = field(default_factory=list)
    bom_detalle:    list  = field(default_factory=list)  # [{posnr, clase_destino}]
    advertencias:   list  = field(default_factory=list)  # advertencias no-fatales
    estado:         str   = "PENDIENTE"   # EN_PROCESO | OK | ERROR
    error:          str   = ""
    fecha_inicio:   Optional[datetime.datetime] = None
    fecha_fin:      Optional[datetime.datetime] = None
    log:            list  = field(default_factory=list)
    # Campos extra para log BD
    formula:        str   = ""
    tipo_pieza:     str   = ""
    acero:          str   = ""
    color_nombre:   str   = ""
    tipo:           str   = ""   # "color" | "formula"

    @property
    def duracion_seg(self) -> float:
        if self.fecha_inicio and self.fecha_fin:
            return round((self.fecha_fin - self.fecha_inicio).total_seconds(), 1)
        return 0.0

    def _log(self, msg: str):
        print(f"  [SAP] {msg}")
        self.log.append(msg)

    def _advertir(self, msg: str):
        print(f"  [SAP][ADV] {msg}")
        self.advertencias.append(msg)
        self.log.append(f"[ADV] {msg}")


# ── Automatizador ─────────────────────────────────────────────────────────────

class AutomatizadorSAP:

    # ── IDs SAP GUI (confirmados por grabaciones VBS) ─────────────────────────
    _ID_TCODE_BOX   = "wnd[0]/tbar[0]/okcd"
    _ID_STATUSBAR   = "wnd[0]/sbar"

    # ZMME0001
    _ID_MATER_LOW   = "wnd[0]/usr/ctxtP_MATER-LOW"
    _ID_CTX_CENTER  = "wnd[0]/usr/ctxtP_CENTER"
    _ID_RAD_HOMOLOG = "wnd[0]/usr/radRB5"
    _ID_RAD_COLOR   = "wnd[0]/usr/radRB3_A1"
    _ID_CTX_P_COLOR = "wnd[0]/usr/ctxtP_COLOR"
    _ID_CTX_P_FRANJ = "wnd[0]/usr/ctxtP_FRANJ"
    _ID_CTX_P_ZPLA  = "wnd[0]/usr/ctxtP_ZPLA"
    _ID_BTN_EXEC    = "wnd[0]/tbar[1]/btn[8]"
    _ID_BTN_BACK    = "wnd[0]/tbar[0]/btn[3]"
    _ID_GRID_RESULT = "wnd[0]/usr/cntlGRID1/shellcont/shell"
    _ID_BTN_COMP    = "wnd[0]/usr/btnBUTTON1"

    # Tabla inferior ZMME0001
    _TBL_BASE        = ("wnd[0]/usr/tabsTABSTRIP_MAX/tabpPUSH1"
                        "/ssub%_SUBSCREEN_MAX:ZMME0001:0200")
    _ID_BTN_INSERT    = _TBL_BASE + "/btnT_LISTA_MATERIA_INSERT"
    _ID_TBL_LISTA     = _TBL_BASE + "/tblZMME0001T_LISTA_MATERIA"
    _ID_BTN_COPY_ITEM = _TBL_BASE + "/btnCOPY_ITEM"

    # ZPPR0020
    _ID_ZPPR_USER   = "wnd[0]/usr/txtS_USER-LOW"
    _ID_ZPPR_CENTRO = "wnd[0]/usr/ctxtS_WERKS-LOW"

    # MM02
    _ID_MM02_MATNR = "wnd[0]/usr/ctxtRMMG1-MATNR"
    _ID_MM02_TAB03 = "wnd[0]/usr/tabsTABSPR1/tabpSP03"
    _ID_MM02_TAB4  = ("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000"
                      "/tabsTABSTRIP_CHAR/tabpTAB4")
    _ID_MM02_TABLA = ("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000"
                      "/tabsTABSTRIP_CHAR/tabpTAB4"
                      "/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100"
                      "/tblSAPLCTMSCHARS_S")

    # ── Init ──────────────────────────────────────────────────────────────────

    def __init__(self):
        self.app      = None
        self.conn_sap = None
        self.session  = None
#excepcion de tiempo 
    def conectar(self) -> bool:
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            self.app     = sap_gui_auto.GetScriptingEngine
            if self.app.Children.Count == 0:
                return False
            self.conn_sap = self.app.Children(0)
            self.session  = self.conn_sap.Children(0)
            self.session.findById("wnd[0]").maximize()
            return True
        except Exception as e:
            print(f"  [ERROR] Conexión SAP: {e}")
            return False

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _esperar(self, max_seg: float = T_RAPIDO):
        """
        Espera inteligente: duerme el mínimo garantizado, luego hace poll de
        application.Busy cada 100ms hasta que SAP esté listo o se alcance max_seg.
        """
        # mínimo garantizado según categoría
        if max_seg <= T_RAPIDO:
            time.sleep(_T_MIN_RAPIDO)
        elif max_seg <= T_MEDIO:
            time.sleep(_T_MIN_MEDIO)
        else:
            time.sleep(_T_MIN_LENTO)

        t0 = time.time()
        while time.time() - t0 < max_seg:
            try:
                if not self.app.Busy:
                    return
            except Exception:
                pass
            time.sleep(0.1)
        # si llegó al tope: continúa de todas formas (no falla)

    def _navegar(self, tcode: str):
        self.session.findById(self._ID_TCODE_BOX).text = f"/N{tcode}"
        self.session.findById("wnd[0]").sendVKey(0)
        self._esperar(T_MEDIO)

    def _estado_sap(self) -> str:
        try:
            return self.session.findById(self._ID_STATUSBAR).text.strip()
        except Exception:
            return ""

    def _sbar(self) -> tuple:
        """Retorna (tipo, texto) del statusbar. tipo: 'E','W','S','I','' """
        try:
            sbar = self.session.findById(self._ID_STATUSBAR)
            txt  = (sbar.text or "").strip()
            try:
                mtype = (sbar.messageType or "").strip().upper()
            except Exception:
                # messageType no disponible en todas las versiones — inferir del texto
                mtype = ""
            if not mtype and txt:
                tup = txt.upper()
                if any(w in tup for w in ("ERROR", "NO EXISTE", "NO SE PUEDE", "INCORRECTO", "FALTA")):
                    mtype = "E"
                elif any(w in tup for w in ("ADVERTENCIA", "WARNING", "ATENCIÓN")):
                    mtype = "W"
            return mtype, txt
        except Exception:
            return "", ""

    def _aceptar_dialogo(self):
        try:
            self.session.findById("wnd[1]").sendVKey(0)
            self._esperar(T_RAPIDO)
        except Exception:
            pass

    def _cerrar_dialogs_abiertos(self):
        for wnd in ("wnd[2]", "wnd[1]"):
            try:
                self.session.findById(wnd)
                try:
                    self.session.findById(wnd).sendVKey(12)
                except Exception:
                    try:
                        self.session.findById(wnd).sendVKey(0)
                    except Exception:
                        pass
                self._esperar(T_RAPIDO)
            except Exception:
                pass

    # ── ZPPP0042 — Validar versión ────────────────────────────────────────────

    def zppp0042_validar(self, zfer: str) -> dict:
        self._navegar("ZPPP0042")
        self.session.findById("wnd[0]/usr/ctxtP_WERKS").text = "CO01"
        self.session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = zfer
        self.session.findById("wnd[0]").sendVKey(8)
        self._esperar(T_LENTO)

        grid_id = "wnd[0]/usr/cntlCC_ALV/shellcont/shell"
        try:
            grid = self.session.findById(grid_id)
            row_count = grid.RowCount
        except Exception as e:
            return {"ok": False, "error": f"Grid ZPPP0042 no encontrado: {e}", "verid": ""}

        if row_count == 0:
            return {"ok": False, "error": f"ZFER {zfer} no encontrado en ZPPP0042", "verid": ""}

        for row in range(row_count):
            mat = ""
            for col in ("MATNR", "MATERIAL"):
                try:
                    mat = grid.getCellValue(row, col).strip()
                    if mat:
                        break
                except Exception:
                    pass
            if mat != zfer:
                continue
            try:
                verid = grid.getCellValue(row, "VERID").strip()
            except Exception:
                verid = ""
            # Verificar utilización lista material
            util = "1"
            for col in ("STLAN", "STLAL"):
                try:
                    v = grid.getCellValue(row, col).strip()
                    if v:
                        util = v
                        break
                except Exception:
                    pass
            if util and util != "1":
                return {"ok": False,
                        "error": f"Utilización lista material = '{util}' (se esperaba '1')",
                        "verid": verid}
            return {"ok": True, "error": "", "verid": verid}

        return {"ok": False, "error": f"ZFER {zfer} no encontrado en grid ZPPP0042", "verid": ""}

    # ── ZMME0001 — ejecutar ───────────────────────────────────────────────────

    def zmme0001_ejecutar(self, zfer_base: str, p_color: str, p_franj: str,
                          zplas_validos: list, forzar_be: bool = False) -> tuple:
        """
        ZMME0001 → Homologar → Cambio de Color → F4 en ZPLA (SAP sugiere) →
        valida contra zplas_validos del combinador → doble clic → F8.
        forzar_be=True: en el popup F4 selecciona la fila cuya descripción
        contenga 'BE' (caso especial nivel 02/03 + tipo pieza 009/090 + clase 0100 termina en 800).
        Retorna (zfer_nuevo, zfor_nuevo, zpla_seleccionado).
        """
        self._cerrar_dialogs_abiertos()
        self._navegar("ZMME0001")
        self.session.findById("wnd[0]").maximize()
        self._esperar(T_RAPIDO)

        # 1. Homologar
        self.session.findById(self._ID_RAD_HOMOLOG).setFocus()
        self.session.findById(self._ID_RAD_HOMOLOG).select()
        self._esperar(T_RAPIDO)

        # 2. Material
        self.session.findById(self._ID_MATER_LOW).text = zfer_base
        self._esperar(T_RAPIDO)

        # 3. Centro
        self.session.findById(self._ID_CTX_CENTER).text = "CO01"

        # 4. Cambio de color
        self.session.findById(self._ID_RAD_COLOR).setFocus()
        self.session.findById(self._ID_RAD_COLOR).select()
        self._esperar(T_RAPIDO)

        # 5. Color y Franja (después del select para que SAP no los limpie)
        self.session.findById(self._ID_CTX_P_COLOR).text = p_color
        self.session.findById(self._ID_CTX_P_FRANJ).text = p_franj

        # 6. F4 en ZPLA → SAP abre popup con sugerencias
        self.session.findById(self._ID_CTX_P_ZPLA).setFocus()
        self.session.findById(self._ID_CTX_P_ZPLA).caretPosition = 0
        self.session.findById("wnd[0]").sendVKey(4)   # F4
        self._esperar(T_LENTO)   # máquina lenta: esperar carga completa del popup

        # 7. Leer popup y seleccionar ZPLA que coincida con los del combinador
        _ID_POPUP_GRID = "wnd[1]/usr/cntlLO_CONTAINER0500/shellcont/shell"
        zpla_seleccionado = ""
        fila_seleccionada = 0

        # Retry: si el grid no cargó todavía, esperar y reintentar una vez
        grid_popup = None
        for _intento_f4 in range(3):
            try:
                grid_popup = self.session.findById(_ID_POPUP_GRID)
                break   # encontrado
            except Exception:
                if _intento_f4 < 2:
                    print(f"    [F4] popup no listo, esperando... (intento {_intento_f4+1})")
                    self._esperar(T_LENTO)
        if grid_popup is None:
            raise RuntimeError(f"F4 ZPLA popup falló: grid no apareció tras 3 intentos")

        try:
            n_filas    = grid_popup.RowCount
            print(f"    F4 ZPLA popup: {n_filas} sugerencias SAP")

            # — Dump columnas reales del grid (solo para debug, una vez) ——————
            try:
                cols_obj = grid_popup.Columns
                col_names_real = [cols_obj.Item(ci).Name for ci in range(cols_obj.Count)]
                print(f"    [DEBUG] Columnas popup: {col_names_real}")
            except Exception as _ce:
                col_names_real = []
                print(f"    [DEBUG] No se pudieron leer columnas: {_ce}")

            # Leer todos los ZPLAs que sugiere SAP (número + descripción)
            zplas_sap  = []
            descs_sap  = []
            for i in range(n_filas):
                val  = ""
                desc = ""
                for col in ("MATNR", "ZPLA", "ZPLARF", "MATERIAL"):
                    try:
                        val = str(grid_popup.GetCellValue(i, col) or "").strip()
                        if val:
                            break
                    except Exception:
                        pass
                # Intentar columnas conocidas + las columnas reales detectadas
                desc_cols = list(dict.fromkeys(
                    ["MAKTX", "DESCR", "TEXT", "MAKTG", "MAKTX_K", "BEZEICHNUNG", "BEZEICH"]
                    + [c for c in col_names_real if c not in ("MATNR", "ZPLA", "ZPLARF", "MATERIAL")]
                ))
                for col in desc_cols:
                    try:
                        desc = str(grid_popup.GetCellValue(i, col) or "").strip()
                        if desc:
                            break
                    except Exception:
                        pass
                zplas_sap.append(val)
                descs_sap.append(desc)
                print(f"      Sugerencia SAP fila {i}: {val}  desc='{desc}'")

            # Selección de fila según contexto
            zplas_set = {z.strip() for z in zplas_validos if z.strip()}
            fila_seleccionada = 0   # default: primera sugerencia

            if forzar_be:
                # Caso especial: buscar fila cuya descripción contenga "BE"
                fila_be = next(
                    (i for i, d in enumerate(descs_sap) if "BE" in d.upper()),
                    None
                )
                if fila_be is not None:
                    fila_seleccionada = fila_be
                    print(f"    [BE] fila seleccionada: {fila_be} = {zplas_sap[fila_be]}  desc='{descs_sap[fila_be]}'")
                else:
                    print(f"    [BE][WARN] No se encontró ZPLA con 'BE' en desc — tomando fila 0")
            else:
                # Flujo normal: buscar coincidencia con zplas_validos del combinador
                for i, z in enumerate(zplas_sap):
                    if z in zplas_set:
                        fila_seleccionada = i
                        print(f"    ZPLA validado: fila {i} = {z} (coincide con combinador)")
                        break
                else:
                    print(f"    [WARN] Ningún ZPLA del popup coincide con el combinador {zplas_validos} — tomando fila 0: {zplas_sap[0] if zplas_sap else '?'}")

            zpla_seleccionado = zplas_sap[fila_seleccionada] if zplas_sap else ""

            # Seleccionar y doble clic (igual que VBS grabado)
            grid_popup.selectedRows = str(fila_seleccionada)
            grid_popup.doubleClickCurrentCell()
            self._esperar(T_RAPIDO)

        except Exception as e:
            raise RuntimeError(f"F4 ZPLA popup falló: {e}")

        # 8. Foco en FRANJ + caretPosition antes de F8 (confirmado por VBS)
        self.session.findById(self._ID_CTX_P_FRANJ).setFocus()
        self.session.findById(self._ID_CTX_P_FRANJ).caretPosition = 2
        self._esperar(T_RAPIDO)

        # 9. F8 Ejecutar
        self.session.findById(self._ID_BTN_EXEC).press()
        self._esperar(T_LENTO)

        # 10. Leer grid resultado
        msg_sap = self._estado_sap()
        try:
            grid       = self.session.findById(self._ID_GRID_RESULT)
            zfer_nuevo = grid.GetCellValue(0, "ZFER").strip()
            zfor_nuevo = grid.GetCellValue(0, "ZFOR").strip()
        except Exception as e:
            raise RuntimeError(
                f"No se pudo leer resultado del grid ZMME0001. SAP: '{msg_sap}'. Detalle: {e}"
            )

        if not zfer_nuevo:
            raise RuntimeError(f"ZFER_NUEVO vacío tras ejecutar ZMME0001. SAP: '{msg_sap}'")

        print(f"    ZMME0001 OK: ZFER_NUEVO={zfer_nuevo} | ZFOR={zfor_nuevo} | ZPLA={zpla_seleccionado}")

        # F3 para volver a pantalla de selección (necesario para el paso 4)
        self.session.findById("wnd[0]").sendVKey(3)
        self._esperar(T_RAPIDO)

        return zfer_nuevo, zfor_nuevo, zpla_seleccionado
        
    # ── ZPPR0020 — Esperar fases en sesión auxiliar ───────────────────────────

    def zppr0020_esperar_fases(self, zfer_nuevo: str,
                                intervalo_seg: int = 5,
                                max_espera_seg: int = 600) -> dict:
        """
        Abre sesión auxiliar SAP para ZPPR0020 (deja ZMME0001 intacta en sesión
        principal). Polling hasta > 7 fases con 'S', o error 'E', o timeout.
        Cierra sesión auxiliar al terminar y re-adquiere sesión principal.
        """
        print("     Abriendo sesión auxiliar para ZPPR0020...")
        self.session.createSession()
        self._esperar(T_MEDIO)

        idx_nueva = self.conn_sap.Children.Count - 1
        ses2 = self.conn_sap.Children(idx_nueva)
        self._esperar(T_MEDIO)   # dar tiempo a que la sesión arranque
        ses2.findById("wnd[0]").maximize()

        # cerrar popup de bienvenida / avisos si existe
        try:
            ses2.findById("wnd[1]").sendVKey(0)
            self._esperar(T_RAPIDO)
        except Exception:
            pass

        try:
            print("     ZPPR0020: navegando...")
            ses2.findById(self._ID_TCODE_BOX).text = "ZPPR0020"
            ses2.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)   # esperar carga ZPPR0020

            # cerrar cualquier popup post-navegación
            for _ in range(3):
                try:
                    ses2.findById("wnd[1]").sendVKey(0)
                    self._esperar(T_RAPIDO)
                except Exception:
                    break

            print("     ZPPR0020: llenando filtros...")
            ses2.findById(self._ID_ZPPR_USER).text   = _SAP_USER
            ses2.findById(self._ID_ZPPR_CENTRO).text = "CO01"
            print("     ZPPR0020: ejecutando F8...")
            ses2.findById(self._ID_BTN_EXEC).press()
            self._esperar(T_LENTO)

            iteraciones       = max(1, max_espera_seg // intervalo_seg)
            sin_datos_counter = 0

            for intento in range(iteraciones):
                resultado = self._leer_zppr0020_grid(zfer_nuevo, ses2)

                if resultado.get("encontrado"):
                    fases = resultado.get("fases", {})
                    zpla  = resultado.get("zpla", "")

                    for nombre_fase, valor in fases.items():
                        if str(valor).strip().upper() == "E":
                            return {"ok": False, "zpla": zpla,
                                    "fase_error": nombre_fase,
                                    "detalle": f"{nombre_fase} en estado 'E' (Error)",
                                    "fases": fases}

                    n_s = sum(1 for v in fases.values() if str(v).strip().upper() == "S")
                    if n_s > 7:
                        print(f"    ZPPR0020 OK: {n_s} fases S | ZPLA={zpla}")
                        return {"ok": True, "zpla": zpla, "fase_error": "",
                                "detalle": f"{n_s} fases completadas", "fases": fases}

                    print(f"    ZPPR0020: {n_s} fases S (esperando >7)... intento {intento+1}/{iteraciones}")
                else:
                    sin_datos_counter += 1
                    print(f"    ZPPR0020: sin datos intento {intento+1}/{iteraciones}")
                    if sin_datos_counter >= 10:
                        return {"ok": False, "zpla": "", "fase_error": "SIN_DATOS",
                                "detalle": f"ZPPR0020 no mostró datos de {zfer_nuevo}", "fases": {}}

                if intento < iteraciones - 1:
                    time.sleep(intervalo_seg)
                    ses2.findById("wnd[0]").sendVKey(9)   # F9 refresh
                    self._esperar(T_MEDIO)

            return {"ok": False, "zpla": "", "fase_error": "TIMEOUT",
                    "detalle": f"ZPPR0020 no completó en {max_espera_seg//60} min.", "fases": {}}

        finally:
            # Cerrar sesión auxiliar con /i
            try:
                ses2.findById(self._ID_TCODE_BOX).text = "/i"
                ses2.findById("wnd[0]").sendVKey(0)
            except Exception:
                try:
                    ses2.findById("wnd[0]").close()
                except Exception:
                    pass
            self._esperar(T_MEDIO)
            # Re-adquirir sesión principal
            try:
                self.session = self.conn_sap.Children(0)
                self.session.findById("wnd[0]").maximize()
            except Exception as e:
                print(f"     [WARN] Re-adquirir sesión: {e}")
            print("     Sesión auxiliar ZPPR0020 cerrada.")

    # ── ZPPR0020 — Leer grid ──────────────────────────────────────────────────

    def _leer_zppr0020_grid(self, zfer_nuevo: str, ses) -> dict:
        resultado = {"encontrado": False, "zpla": "", "fases": {}}

        grid = None
        for ruta in (
            "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell",
            "wnd[0]/usr/cntlGRID1/shellcont/shell",
            "wnd[0]/usr/cntlGRID/shellcont/shell",
            "wnd[0]/shellcont/shell",
        ):
            try:
                obj = ses.findById(ruta)
                if obj.RowCount is not None:
                    grid = obj
                    break
            except Exception:
                pass

        if grid is None:
            try:
                grid = self._buscar_grid_recursivo(ses.findById("wnd[0]"))
            except Exception:
                pass

        if grid is None:
            print("    [WARN] ZPPR0020: grid no encontrado.")
            return resultado

        # Nombres confirmados por log real: MAT_ZFER, MAT_ZPLA, PHASE1..PHASE18
        _COLS_ZFER = ("MAT_ZFER", "ZFER", "MATNR_ZFER", "ZFER_NEW", "MATNR", "MATERIAL")
        _COLS_ZPLA = ("MAT_ZPLA", "ZPLA", "MATNR_ZPLA", "ZPLA_NEW")

        def _leer(fila, candidatas):
            for col in candidatas:
                try:
                    v = str(grid.GetCellValue(fila, col) or "").strip()
                    if v:
                        return v, col
                except Exception:
                    pass
            return "", ""
        
        try:
            n_filas = grid.RowCount
            if not n_filas:
                return resultado

            # Obtener TODAS las columnas reales del grid
            all_cols = []
            try:
                co = grid.ColumnOrder
                if isinstance(co, str):
                    all_cols = co.split()
                else:
                    all_cols = [str(c) for c in co]
            except Exception:
                pass
            print(f"    [DEBUG] ZPPR0020 columnas reales: {all_cols}")

            # Detectar columnas de fase dinámicamente: cualquier columna cuyo nombre
            # contenga "FASE", "PHASE" o "F0" y empiece por F/P
            _COLS_FASE = tuple(
                c for c in all_cols
                if any(k in c.upper() for k in ("FASE", "PHASE"))
                or (c.upper().startswith("F") and c[1:].isdigit())
            )
            if not _COLS_FASE:
                # Fallback: nombres confirmados por log real (PHASE1..18) + variantes
                _COLS_FASE = (
                    tuple(f"PHASE{i}"    for i in range(1, 19)) +
                    tuple(f"FASE{i}"     for i in range(1, 16)) +
                    tuple(f"FASE_{i:02}" for i in range(1, 16)) +
                    tuple(f"F{i:02}"     for i in range(1, 16))
                )
            print(f"    [DEBUG] ZPPR0020 cols fase detectadas: {_COLS_FASE}")

            # Debug fila 0: mostrar valor de CADA columna real
            if all_cols:
                vals_debug = {}
                for col in all_cols:
                    try:
                        v = str(grid.GetCellValue(0, col) or "").strip()
                        if v:
                            vals_debug[col] = v
                    except Exception:
                        pass
                print(f"    [DEBUG] ZPPR0020 fila0 valores: {vals_debug}")

            for i in range(n_filas):
                zfer_fila, col_zfer = _leer(i, _COLS_ZFER)
                if zfer_fila != zfer_nuevo:
                    continue

                resultado["encontrado"] = True
                zpla_val, _ = _leer(i, _COLS_ZPLA)
                resultado["zpla"] = zpla_val
                fases = {}
                for col in _COLS_FASE:
                    try:
                        v = str(grid.GetCellValue(i, col) or "").strip()
                        if v:
                            fases[col] = v
                    except Exception:
                        pass
                resultado["fases"] = fases
                n_s = sum(1 for v in fases.values() if v.upper() == "S")
                n_e = sum(1 for v in fases.values() if v.upper() == "E")
                print(f"    ZPPR0020 fila {i} (col={col_zfer}): ZPLA={zpla_val} | S={n_s} | E={n_e} | fases={fases}")
                break

        except Exception as e:
            print(f"    [WARN] _leer_zppr0020_grid: {e}")

        return resultado

    def _buscar_grid_recursivo(self, contenedor, profundidad: int = 0):
        if profundidad > 8:
            return None
        try:
            n = contenedor.Children.Count
        except Exception:
            return None
        for i in range(n):
            try:
                hijo = contenedor.Children(i)
            except Exception:
                continue
            try:
                tipo = hijo.Type
            except Exception:
                tipo = ""
            if tipo in ("GuiGridView", "GuiTableControl"):
                return hijo
            encontrado = self._buscar_grid_recursivo(hijo, profundidad + 1)
            if encontrado is not None:
                return encontrado
        return None

    # ── ZMME0001 — Comparar BOM, llenar tabla, COPY_ITEM ─────────────────────

    def zmme0001_leer_posiciones_popup(self) -> list:
        """
        Retorna lista de dicts: {pos, tipo, msg}
          tipo 5 → modificar clase
          tipo 6 → agregar posición (el ZPLA tiene la pos, el ZFER no)
          tipo 7 → eliminar posición (el ZFER tiene la pos, el ZPLA no)
        Lee la columna Nº directamente del grid (más fiable que parsear texto).
        """
        self.session.findById(self._ID_BTN_COMP).press()
        self._esperar(T_MEDIO)

        filas = []

        def _tipo_de_num(num_str: str, msg: str) -> int:
            """Intenta leer el número directamente; fallback a texto."""
            try:
                n = int(num_str.strip())
                if n in (5, 6, 7):
                    return n
            except Exception:
                pass
            # fallback por texto
            m = msg.upper()
            if "NO INCLUIDA EN EL ZPLA" in m or "NO ESTA INCLUIDO EN ZPLA" in m or "BOM NO ESTA" in m:
                return 7
            if "ZPLA MODELO TIENE" in m or "NO TIENE EL B" in m or "AGREGAR" in m:
                return 6
            return 5

        # GuiGridView (es lo que muestra la imagen)
        try:
            grid = self.session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell")
            for i in range(grid.RowCount):
                try:
                    pos = grid.GetCellValue(i, "POSNR").strip()
                    if not pos:
                        continue
                    # Columna Nº — probar nombres posibles
                    num_str = ""
                    for col in ("MSGNO", "NR", "NUMERO", "NUM", "MSG_NR", "MNUM", "NO"):
                        try:
                            v = grid.GetCellValue(i, col).strip()
                            if v:
                                num_str = v
                                break
                        except Exception:
                            pass
                    # Mensaje de texto
                    msg = ""
                    for col in ("VARIABLE_MENSAJE", "MSG", "MESSAGE", "TEXT", "MELDUNG"):
                        try:
                            v = grid.GetCellValue(i, col).strip()
                            if v:
                                msg = v
                                break
                        except Exception:
                            pass
                    tipo = _tipo_de_num(num_str, msg)
                    filas.append({"pos": pos, "tipo": tipo, "num": num_str, "msg": msg})
                except Exception:
                    pass
        except Exception:
            pass

        # Fallback GuiTableControl clásico
        if not filas:
            try:
                tabla = self.session.findById("wnd[1]/usr/tblZMME0001T_COMP")
                for i in range(tabla.RowCount):
                    try:
                        pos = tabla.GetCell(i, 0).text.strip()
                        if not pos:
                            continue
                        num_str = ""
                        msg     = ""
                        try:
                            num_str = tabla.GetCell(i, 2).text.strip()  # columna Nº
                        except Exception:
                            pass
                        try:
                            msg = tabla.GetCell(i, 3).text.strip()
                        except Exception:
                            pass
                        tipo = _tipo_de_num(num_str, msg)
                        filas.append({"pos": pos, "tipo": tipo, "num": num_str, "msg": msg})
                    except Exception:
                        pass
            except Exception:
                pass

        # Cerrar popup
        try:
            self.session.findById("wnd[1]").close()
        except Exception:
            try:
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except Exception:
                pass
        self._esperar(T_RAPIDO)

        # Deduplicar por posición manteniendo orden
        vistos = set()
        result = []
        for f in filas:
            if f["pos"] not in vistos:
                vistos.add(f["pos"])
                result.append(f)
        return result

    def zmme0001_agregar_filas_bom(self, posiciones: list, zpla: str,
                                    clases_dict: dict = None, row_offset: int = 0):
        """
        posiciones: lista de dicts {pos, tipo, msg} de leer_posiciones_popup().
        tipo 5 → INSERT + POSNR + CLASE_DESTINO
        tipo 7 → INSERT + POSNR + marcar ELIMINAR (sin clase)
        tipo 6 → INSERT + POSNR + NEW_POSNR (múltiplo de 5) + CLASE_DESTINO
        clases_dict: {posicion: clase}
        """
        if clases_dict is None:
            clases_dict = {}
            
        # Soporte retrocompatible: si llega lista de strings en vez de dicts
        posiciones = [
            p if isinstance(p, dict) else {"pos": p, "tipo": 5, "msg": ""}
            for p in posiciones
        ]

        # ── Referencia para tipo 6 ────────────────────────────────────────────────
        # Candidatas para el NÚMERO de referencia: TODAS las posiciones de la lista
        # (incluyendo tipo 7 que se borran) — la 458 puede referenciar a 358.
        # La CLASE siempre viene de clases_dict[referencia] o clases_dict[pos_nueva].
        # Regla último dígito:
        #   termina en 00 → candidatas k%100==0
        #   termina en 0 (no 00) → candidatas k%10==5 (00 son intocables)
        #   otros → candidatas k%10==mismo último dígito (puede repetirse)
        _claves_con_clase = set()   # para clase
        for _k in clases_dict.keys():
            try:
                _claves_con_clase.add(int(str(_k).lstrip("0") or "0"))
            except Exception:
                pass
        _todas_posiciones = set()   # para referencia numérica
        for _p in posiciones:
            try:
                _todas_posiciones.add(int(str(_p["pos"]).lstrip("0") or "0"))
            except Exception:
                pass

        def _calcular_ref(pos_str) -> str:
            try:
                pos_int = int(str(int(pos_str)))
                # Candidatas: TODAS las posiciones menores (incl. tipo 7)
                menores = [k for k in _todas_posiciones if k < pos_int]
                if not menores:
                    return ""
                if pos_int % 100 == 0:
                    cands = [k for k in menores if k % 100 == 0]
                elif pos_int % 10 == 0:
                    # Termina en 0 no 00 → busca en 5
                    cands = [k for k in menores if k % 10 == 5]
                else:
                    ultimo = pos_int % 10
                    cands = [k for k in menores if k % 10 == ultimo]
                return str(max(cands)) if cands else str(max(menores))
            except Exception:
                pass
            return ""

        # ── Pre-validación ────────────────────────────────────────────────────────
        faltantes = []
        for item in posiciones:
            tipo = item.get("tipo", 5)
            pos  = item["pos"]
            if tipo == 5:
                clase = clases_dict.get(pos.zfill(4), clases_dict.get(pos, ""))
                if not clase:
                    faltantes.append(f"{pos}(tipo5: clase propia no encontrada en ZPPR0008)")
            elif tipo == 6:
                ref = _calcular_ref(pos)
                if not ref:
                    faltantes.append(f"{pos}(tipo6: sin posición anterior válida)")
                else:
                    ref_key = ref.zfill(4)
                    clase_ref = clases_dict.get(ref_key, clases_dict.get(ref, ""))
                    clase_propia = clases_dict.get(pos.zfill(4), clases_dict.get(pos, ""))
                    if not clase_ref and not clase_propia:
                        faltantes.append(f"{pos}(tipo6: ni referencia {ref} ni posición propia tienen clase en ZPPR0008)")
        if faltantes:
            raise RuntimeError(
                f"BOM incompleto — faltan clases en ZPPR0008 para: {'; '.join(faltantes)}. "
                f"Verifica que el ZPLA {zpla} tenga esas posiciones con clase asignada."
            )

        # Tipo 7 (ELIMINAR) siempre al final para que las referencias ya existan
        posiciones = sorted(posiciones, key=lambda p: 1 if p.get("tipo", 5) == 7 else 0)

        for idx, item in enumerate(posiciones):
            pos  = item["pos"]
            tipo = item.get("tipo", 5)
            pos_sin_ceros = str(int(pos)) if pos.isdigit() else pos

            self.session.findById(self._ID_BTN_INSERT).press()
            self._esperar(T_RAPIDO)

            # ── Calcular índice visible en la tabla ──────────────────────────────
            # La tabla SAP solo muestra VisibleRowCount filas a la vez.
            # Cuando idx >= VisibleRowCount, hay que scrollar y usar el índice
            # relativo al viewport (vis_idx) para acceder a los controles de celda.
            abs_idx = row_offset + idx   # posición absoluta en la tabla
            vis_idx = abs_idx
            try:
                tbl_obj = self.session.findById(self._ID_TBL_LISTA)
                try:
                    vis_cnt = int(tbl_obj.VisibleRowCount or 0)
                except Exception:
                    vis_cnt = 0
                if vis_cnt <= 0:
                    vis_cnt = 8  # fallback conservador si SAP no expone VisibleRowCount
                print(f"    [SCROLL] abs={abs_idx} vis_cnt={vis_cnt}")
                if abs_idx >= vis_cnt:
                    primer_vis = abs_idx - vis_cnt + 1
                    tbl_obj.VerticalScrollbar.Position = primer_vis
                    self._esperar(T_MEDIO)
                    vis_idx = vis_cnt - 1   # la nueva fila queda en la última posición visible
            except Exception as _se:
                print(f"    [WARN] scroll table: {_se}")

            if tipo == 7:
                # ELIMINAR POSICION: marcar checkbox + POSNR
                self.session.findById(
                    f"{self._ID_TBL_LISTA}/txtWA_LISTA-POSNR[0,{vis_idx}]"
                ).text = pos_sin_ceros
                self._esperar(T_RAPIDO)
                try:
                    self.session.findById(
                        f"{self._ID_TBL_LISTA}/chkWA_LISTA-ELIMINAR[4,{vis_idx}]"
                    ).selected = True
                    self._esperar(T_RAPIDO)
                except Exception as e:
                    print(f"    [WARN] No pudo marcar ELIMINAR fila {vis_idx}: {e}")
                print(f"    Fila {idx}(vis={vis_idx}): POS={pos_sin_ceros} → ELIMINAR (tipo 7)")

            elif tipo == 6:
                referencia = _calcular_ref(pos_sin_ceros)

                # col 0 = referencia, col 1 = posición a agregar
                self.session.findById(
                    f"{self._ID_TBL_LISTA}/txtWA_LISTA-POSNR[0,{vis_idx}]"
                ).text = referencia
                self._esperar(T_RAPIDO)
                try:
                    self.session.findById(
                        f"{self._ID_TBL_LISTA}/txtWA_LISTA-NEW_POSNR[1,{vis_idx}]"
                    ).text = pos_sin_ceros
                    self._esperar(T_RAPIDO)
                except Exception as e:
                    print(f"    [WARN] No pudo escribir NEW_POSNR fila {vis_idx}: {e}")

                # Clase: de la referencia si tiene, sino de la posición nueva misma
                ref_key = referencia.zfill(4)
                clase = (clases_dict.get(ref_key, clases_dict.get(referencia, ""))
                         or clases_dict.get(pos.zfill(4), clases_dict.get(pos, "")))
                if clase:
                    self.session.findById(
                        f"{self._ID_TBL_LISTA}/ctxtWA_LISTA-CLASE_DESTINO[3,{vis_idx}]"
                    ).text = clase
                    self._esperar(T_RAPIDO)
                print(f"    Fila {idx}(vis={vis_idx}): POSNR={referencia} NEW_POSNR={pos_sin_ceros} CLASE={clase or '(sin clase)'} → AGREGAR (tipo 6)")

            else:
                # MODIFICAR CLASE (tipo 5, default)
                self.session.findById(
                    f"{self._ID_TBL_LISTA}/txtWA_LISTA-POSNR[0,{vis_idx}]"
                ).text = pos_sin_ceros
                self._esperar(T_RAPIDO)
                clase = clases_dict.get(pos.zfill(4), clases_dict.get(pos, ""))
                if clase:
                    self.session.findById(
                        f"{self._ID_TBL_LISTA}/ctxtWA_LISTA-CLASE_DESTINO[3,{vis_idx}]"
                    ).text = clase
                    self._esperar(T_RAPIDO)
                print(f"    Fila {idx}(vis={vis_idx}): POS={pos_sin_ceros} CLASE={clase or '(sin clase)'} → MODIFICAR (tipo 5)")

    def zmme0001_segunda_comparar_y_copy(self) -> bool:
        self.session.findById(self._ID_BTN_COMP).press()
        self._esperar(T_MEDIO)

        ok = True
        try:
            try:
                grid_err = self.session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell")
                if grid_err.RowCount > 0:
                    try:
                        tipo = grid_err.GetCellValue(0, "TY").strip()
                        if tipo.upper() == "E":
                            msg = grid_err.GetCellValue(0, "VARIABLE_MENSAJE").strip()
                            ok  = False
                            print(f"    [ERROR] Segunda comparación: {msg}")
                    except Exception:
                        pass
            except Exception:
                pass
            try:
                self.session.findById("wnd[1]").close()
            except Exception:
                try:
                    self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                except Exception:
                    pass
            self._esperar(T_RAPIDO)
        except Exception:
            pass
        if ok:
            try:
                self.session.findById(self._ID_BTN_COPY_ITEM).press()
                self._esperar(T_LENTO)
                msg = self._estado_sap()
                print(f"    COPY_ITEM: {msg}")
                if msg and "error" in msg.lower():
                    ok = False
            except Exception as e:
                print(f"    [WARN] COPY_ITEM: {e}")

        # Cerrar popup de log "Programa ZMME0001" que aparece tras COPY_ITEM
        # El usuario lo cierra con la X roja del popup → wnd[1].close() / F12
        self._esperar(T_RAPIDO)
        for wnd_id in ("wnd[2]", "wnd[1]"):
            try:
                self.session.findById(wnd_id)
                # Intentar primero la X roja (sendVKey 12 = F12/Escape)
                try:
                    self.session.findById(wnd_id).sendVKey(12)
                    self._esperar(T_RAPIDO)
                    print(f"    Popup {wnd_id} cerrado (F12)")
                except Exception:
                    try:
                        self.session.findById(wnd_id).close()
                        self._esperar(T_RAPIDO)
                        print(f"    Popup {wnd_id} cerrado (.close)")
                    except Exception:
                        pass
            except Exception:
                pass

        return ok

    def bom_con_retry(self, zpla_usado: str, clases: dict,
                      max_intentos: int = 3, on_retry=None) -> list:
        """
        Loop BOM en la misma pantalla ZMME0001 sin navegar:
          1. Comparar BOM → leer popup de errores
          2. Si hay errores 5/6/7 → insertar esas posiciones → volver a 1
          3. Cuando popup está limpio → Ejecutar BOM (COPY_ITEM) y cerrar log
        Máx max_intentos ciclos. Lanza RuntimeError si los agota.
        """
        posiciones_acum = []
        filas_ya_insertadas = 0
        for intento in range(1, max_intentos + 1):
            print(f"    [BOM] Comparar BOM — ciclo {intento}/{max_intentos}")
            posiciones = self.zmme0001_leer_posiciones_popup()

            if not posiciones:
                # Sin errores → Ejecutar BOM
                print(f"    [BOM] Sin errores en ciclo {intento} → Ejecutar BOM")
                try:
                    self.session.findById(self._ID_BTN_COPY_ITEM).press()
                    self._esperar(T_LENTO)
                    sbar_tipo, sbar_txt = self._sbar()
                    print(f"    Ejecutar BOM sbar: [{sbar_tipo}] {sbar_txt!r}")
                    if sbar_txt:
                        if sbar_tipo == "E":
                            raise RuntimeError(f"BOM COPY_ITEM error SAP: {sbar_txt}")
                        elif sbar_tipo in ("W", ""):
                            # advertencia o mensaje informativo — no frena pero queda en log
                            if on_retry:
                                on_retry(f"[BOM-SAP] {sbar_txt}")
                            else:
                                print(f"    [ADV][BOM] sbar: {sbar_txt}")
                    for wnd_id in ("wnd[2]", "wnd[1]"):
                        try:
                            self.session.findById(wnd_id)
                            try:
                                self.session.findById(wnd_id).sendVKey(12)
                                self._esperar(T_RAPIDO)
                            except Exception:
                                try:
                                    self.session.findById(wnd_id).close()
                                except Exception:
                                    pass
                        except Exception:
                            pass
                    # Segunda lectura de sbar DESPUÉS de cerrar popups
                    sbar_tipo2, sbar_txt2 = self._sbar()
                    if sbar_tipo2 == "E" and sbar_txt2:
                        raise RuntimeError(f"BOM post-popup error SAP: {sbar_txt2}")
                    if sbar_txt2 and sbar_tipo2 in ("W", "") and sbar_txt2 != sbar_txt:
                        print(f"    [ADV][BOM] sbar post-popup: {sbar_txt2}")
                        if on_retry:
                            on_retry(f"[BOM-SAP-post] {sbar_txt2}")
                except RuntimeError:
                    raise
                except Exception as _e:
                    print(f"    [WARN] Ejecutar BOM: {_e}")
                return posiciones_acum

            posiciones_acum = posiciones
            print(f"    [BOM] {len(posiciones)} errores: {[p['pos'] for p in posiciones]}")

            if intento == max_intentos:
                raise RuntimeError(
                    f"BOM: Comparar BOM sigue con {len(posiciones)} errores tras "
                    f"{max_intentos} ciclos. Posiciones: {[p['pos'] for p in posiciones]}"
                )

            self.zmme0001_agregar_filas_bom(posiciones, zpla_usado, clases,
                                            row_offset=filas_ya_insertadas)
            filas_ya_insertadas += len(posiciones)

        return posiciones_acum

    # ── MM02 — Actualizar PARTNUMBER ─────────────────────────────────────────

    def _primera_opcion_si_popup(self, ses=None):
        """Acepta wnd[1] con OPTION1 si existe (cualquier popup SAP)."""
        s = ses or self.session
        try:
            s.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            self._esperar(T_RAPIDO)
            return True
        except Exception:
            pass
        try:
            s.findById("wnd[1]").sendVKey(0)
            self._esperar(T_RAPIDO)
            return True
        except Exception:
            return False

    def mm02_actualizar_partnumber(self, material: str, nuevo_pn: str):
        """
        MM02 → Clasificación → PIEZA → actualiza PARTNUMBER AGP.
        Flujo confirmado por VBS: Enter tras llenar → F3 → OPTION1 → F3.
        Maneja pantalla de selección de vistas y niveles org. con Enter.
        """
        self._cerrar_dialogs_abiertos()

        # Navegación limpia con /N
        self.session.findById(self._ID_TCODE_BOX).text = "/NMM02"
        self.session.findById("wnd[0]").sendVKey(0)
        self._esperar(T_MEDIO)

        self.session.findById(self._ID_MM02_MATNR).text = material
        self.session.findById("wnd[0]").sendVKey(0)
        self._esperar(T_MEDIO)

        # Pasar pantallas intermedias de MM02:
        # 1) Selección de vistas  2) Niveles organizativos (Centro)
        # Cada una se acepta con Enter — hasta 3 intentos
        for _ in range(3):
            # Si ya vemos el tab de Clasificación, terminamos
            try:
                self.session.findById(self._ID_MM02_TAB03)
                break
            except Exception:
                pass
            # Primero intentar OPTION1, si no existe mandar Enter
            if not self._primera_opcion_si_popup():
                self.session.findById("wnd[0]").sendVKey(0)
                self._esperar(T_MEDIO)

        # Tab Clasificación
        self.session.findById(self._ID_MM02_TAB03).select()
        self._esperar(T_RAPIDO)
        self._primera_opcion_si_popup()

        # Tab PIEZA
        self.session.findById(self._ID_MM02_TAB4).select()
        self._esperar(T_RAPIDO)

        # Escribir nuevo PARTNUMBER en fila 0
        campo_pn = f"{self._ID_MM02_TABLA}/ctxtRCTMS-MWERT[1,0]"
        self.session.findById(campo_pn).text = nuevo_pn
        self.session.findById(campo_pn).caretPosition = len(nuevo_pn)
        self.session.findById("wnd[0]").sendVKey(0)
        self._esperar(T_RAPIDO)

        # F3 → SAP pregunta si guardar → primera opción → F3
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        self._esperar(T_MEDIO)
        self._primera_opcion_si_popup()
        try:
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self._esperar(T_RAPIDO)
        except Exception:
            pass

        print(f"      MM02 {material} PARTNUMBER → {nuevo_pn}")

    # ── SAP: leer posicion→clase del ZPLA en sesión auxiliar ─────────────────

    def _leer_clases_zpla_sap(self, zpla: str) -> dict:
        """
        Abre sesión auxiliar, navega ZPPR0008 (radio LMat alt.), filtra por
        material=ZPLA y centro=CO01, lee el ALV completo.
        Retorna {posicion_zfill4: clase, posicion: clase}.
        IDs y flujo confirmados por VBS grabado (zppr0008.vbs).
        """
        resultado = {}
        print(f"     Leyendo clases ZPLA {zpla} desde ZPPR0008...")
        self.session.createSession()
        self._esperar(T_MEDIO)

        idx_nueva = self.conn_sap.Children.Count - 1
        ses = self.conn_sap.Children(idx_nueva)
        ses.findById("wnd[0]").maximize()

        try:
            ses.findById(self._ID_TCODE_BOX).text = "ZPPR0008"
            ses.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)

            # Radio "por material" — confirmado por VBS
            ses.findById("wnd[0]/usr/radRB_2").setFocus()
            ses.findById("wnd[0]/usr/radRB_2").select()
            self._esperar(T_RAPIDO)

            ses.findById("wnd[0]/usr/ctxtS_MATNR2-LOW").text = zpla
            ses.findById("wnd[0]/usr/ctxtS_WERKS2-LOW").text  = "CO01"
            ses.findById("wnd[0]/usr/ctxtS_WERKS2-LOW").setFocus()
            ses.findById("wnd[0]/usr/ctxtS_WERKS2-LOW").caretPosition = 4

            ses.findById("wnd[0]/tbar[1]/btn[8]").press()   # F8 Ejecutar
            self._esperar(T_LENTO)

            grid = ses.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            n    = grid.RowCount
            print(f"      ZPPR0008 grid: {n} filas")

            # Imprimir columnas disponibles para debug
            try:
                cols = list(grid.ColumnOrder)
                print(f"      Columnas disponibles: {cols}")
            except Exception:
                cols = []

            for i in range(n):
                pos   = ""
                clase = ""
                for col in ("POSNR", "POSICION", "POS"):
                    try:
                        v = str(grid.GetCellValue(i, col) or "").strip()
                        if v:
                            pos = v
                            break
                    except Exception:
                        pass
                for col in ("CLASE", "IDNRK", "MATNR_K", "COMP", "CLASS"):
                    try:
                        v = str(grid.GetCellValue(i, col) or "").strip()
                        if v:
                            clase = v
                            break
                    except Exception:
                        pass
                if pos:
                    resultado[pos.zfill(4)] = clase
                    resultado[pos]          = clase
                    print(f"      ZPPR0008 fila {i}: POS={pos.zfill(4)} CLASE={clase or '(vacía)'}")

            if not resultado:
                print(f"      [WARN] Sin datos en ZPPR0008 para ZPLA {zpla}")

        except Exception as e:
            print(f"      [WARN] _leer_clases_zpla_sap {zpla}: {e}")
        finally:
            try:
                ses.findById(self._ID_TCODE_BOX).text = "/i"
                ses.findById("wnd[0]").sendVKey(0)
            except Exception:
                try:
                    ses.findById("wnd[0]").close()
                except Exception:
                    pass
            self._esperar(T_MEDIO)
            # Re-adquirir sesión principal
            try:
                self.session = self.conn_sap.Children(0)
                self.session.findById("wnd[0]").maximize()
            except Exception as e:
                print(f"     [WARN] Re-adquirir sesión tras clase ZPLA: {e}")
            print("     Sesión clase ZPLA cerrada.")

        return resultado

    # ── BD: log ───────────────────────────────────────────────────────────────

    def _log_bd(self, res: "ResultadoItem"):
        # Extraer plano anterior/nuevo del log del resultado
        plano_ant, plano_nvo = "", ""
        for line in (res.log or []):
            if "[PLANO] Guardado:" in line:
                # "  [PLANO] Guardado: 'X' → 'Y'"
                try:
                    parts = line.split("'")
                    plano_ant = parts[1][:100] if len(parts) > 1 else ""
                    plano_nvo = parts[3][:100] if len(parts) > 3 else ""
                except Exception:
                    pass

        _SQL_INSERT = (
            "INSERT INTO dbo.M5_LogEjecucion "
            "(batch_id, pedido_origen, tipo_pieza, formula, color_codigo, color_nombre, "
            " acero_variante, tipo, zfer_nuevo, zfor_nuevo, zpla, "
            " estado, detalle_error, duracion_seg, plano_anterior, plano_nuevo, "
            " fecha_inicio, fecha_fin) "
            "VALUES (?,?,?,?,?,?, ?,?,?,?,?, ?,?,?,?,?, ?,?)"
        )
        _vals = (
            str(res.batch_id)[:50],
            str(res.zfer_base or "")[:50],          # pedido_origen nvarchar(50)
            str(getattr(res, "tipo_pieza",   "") or "")[:20],   # nvarchar(20)
            str(getattr(res, "formula",      "") or "")[:20],   # nvarchar(20)
            str(res.color_codigo or "")[:50],       # color_codigo nvarchar(50)
            str(getattr(res, "color_nombre", "") or "")[:100],  # varchar(100)
            str(getattr(res, "acero",        "") or "")[:5],    # acero_variante nvarchar(5)
            str(getattr(res, "tipo",         "") or "")[:20],   # varchar(20)
            str(res.zfer_nuevo or "")[:20],
            str(res.zfor_nuevo or "")[:20],
            str(res.zpla or "")[:20],
            str(res.estado or "")[:20],
            str(res.error)[:2000] if res.error else None,
            res.duracion_seg or None,
            plano_ant or None,
            plano_nvo or None,
            res.fecha_inicio,
            res.fecha_fin,
        )
        try:
            cn  = pyodbc.connect(_DB_LOCAL_STR, autocommit=True)
            cur = cn.cursor()
            cur.execute(_SQL_INSERT, _vals)
            cn.close()
        except Exception as e:
            print(f"    [WARN] log_bd: {e}")

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _construir_nuevo_pn(self, pn_base: str, p_color: str) -> str:
        if not pn_base or not p_color:
            return pn_base
        partes = pn_base.split("_")
        if len(partes) >= 4:
            partes[3] = p_color
            return "_".join(partes)
        return pn_base

    def _construir_nuevo_pn_formula(self, pn_base: str, formula_nueva: str, p_color: str) -> str:
        """Reemplaza índice [2] (fórmula) e índice [3] (color) del PARTNUMBER."""
        if not pn_base:
            return pn_base
        partes = pn_base.split("_")
        if len(partes) >= 4:
            if formula_nueva:
                partes[2] = formula_nueva
            if p_color:
                partes[3] = p_color
            return "_".join(partes)
        return pn_base

    def _construir_plano_desde_pn(self, pn_base: str) -> str:
        """Construye nombre de plano desde PARTNUMBER: M{[0]}   {[1]}   {[4]}"""
        partes = pn_base.split("_")
        if len(partes) >= 5:
            return f"M{partes[0]}   {partes[1]}   {partes[4]}"
        elif len(partes) >= 2:
            return f"M{partes[0]}   {partes[1]}   001"
        return ""

    # ── ZPPR0008 — Validar posición acero ────────────────────────────────────

    def zppr0008_leer_bom_completo(self, material: str) -> dict:
        """
        Entra a ZPPR0008 (modo material radRB_2), lee TODAS las posiciones del BOM.
        Retorna {"ok": True, "posiciones": [int,...], "filas": [{pos, nombre},...], "error": ""}
        """
        print(f"    ZPPR0008 BOM: leyendo posiciones para {material}")
        self._navegar("ZPPR0008")
        self.session.findById("wnd[0]/usr/radRB_2").setFocus()
        self.session.findById("wnd[0]/usr/radRB_2").select()
        self._esperar(T_RAPIDO)
        self.session.findById("wnd[0]/usr/ctxtS_MATNR2-LOW").text = material
        self.session.findById("wnd[0]/usr/ctxtS_WERKS2-LOW").text = "CO01"
        self.session.findById("wnd[0]/usr/ctxtS_WERKS2-LOW").caretPosition = 4
        self.session.findById(self._ID_BTN_EXEC).press()
        self._esperar(T_LENTO)

        grid = None
        for _gid in ("wnd[0]/usr/cntlGRID1/shellcont/shell",
                     "wnd[0]/usr/cntlGRID/shellcont/shell",
                     "wnd[0]/usr/cntlALV/shellcont/shell"):
            try:
                grid = self.session.findById(_gid)
                break
            except Exception:
                pass

        if grid is None:
            return {"ok": False, "posiciones": [], "filas": [],
                    "error": f"ZPPR0008: grid no encontrado para {material}"}

        try:
            n = grid.RowCount
            print(f"    ZPPR0008 BOM: {n} filas para {material}")
            # Imprimir columnas disponibles en primera fila para diagnóstico
            if n > 0:
                try:
                    print(f"    [DIAG] ZPPR0008 BOM columnas: {list(grid.ColumnOrder)}")
                except Exception:
                    pass

            posiciones, filas = [], []
            for i in range(n):
                pos_int = None
                for col in ("POSNR", "POSN", "POS"):
                    try:
                        v = str(grid.GetCellValue(i, col) or "").strip()
                        if v:
                            pos_int = int(v.lstrip("0") or "0")
                            break
                    except Exception:
                        pass
                if pos_int is None:
                    continue
                nombre = ""
                for col in ("MATNR1", "MATNR", "KTNAM", "TXT_OBJEK", "OBJECTKEY", "COMPONENT"):
                    try:
                        v = str(grid.GetCellValue(i, col) or "").strip()
                        if v:
                            nombre = v
                            break
                    except Exception:
                        pass
                posiciones.append(pos_int)
                filas.append({"pos": pos_int, "nombre": nombre})

            return {"ok": True, "posiciones": posiciones, "filas": filas, "error": ""}
        except Exception as e:
            return {"ok": False, "posiciones": [], "filas": [], "error": str(e)}

    def zppr0008_validar_posicion_acero(self, zfer_base: str) -> dict:
        """
        Entra a ZPPR0008 con el ZFER base (modo material, radRB_1) y busca posición 0106 ó 0116.
        Retorna {"ok": True/False, "pos": "0106"|"0116"|"", "error": ""}
        """
        print(f"    ZPPR0008: validando posición acero para ZFER={zfer_base}")
        self._navegar("ZPPR0008")

        # Template lista de materiales: radRB_2 + ctxtS_MATNR2-LOW con ZFER base
        self.session.findById("wnd[0]/usr/radRB_2").setFocus()
        self.session.findById("wnd[0]/usr/radRB_2").select()
        self._esperar(T_RAPIDO)

        self.session.findById("wnd[0]/usr/ctxtS_MATNR2-LOW").text = zfer_base
        self.session.findById("wnd[0]/usr/ctxtS_WERKS2-LOW").text = "CO01"
        self.session.findById("wnd[0]/usr/ctxtS_WERKS2-LOW").setFocus()
        self.session.findById("wnd[0]/usr/ctxtS_WERKS2-LOW").caretPosition = 4
        self.session.findById(self._ID_BTN_EXEC).press()
        self._esperar(T_LENTO)

        # Intentar múltiples IDs de grid (el nombre varía según template/versión SAP)
        grid = None
        for _gid in ("wnd[0]/usr/cntlGRID1/shellcont/shell",
                     "wnd[0]/usr/cntlGRID/shellcont/shell",
                     "wnd[0]/usr/cntlALV/shellcont/shell"):
            try:
                grid = self.session.findById(_gid)
                print(f"    ZPPR0008: grid encontrado en {_gid}")
                break
            except Exception:
                pass

        if grid is None:
            # Diagnóstico: listar hijos de wnd[0]/usr
            try:
                _usr = self.session.findById("wnd[0]/usr")
                _ids = [_usr.Children(i).Id for i in range(min(_usr.Children.Count, 20))]
                print(f"    [DIAG] ZPPR0008 usr hijos tras F8: {_ids}")
            except Exception as _de:
                print(f"    [DIAG] ZPPR0008 no pudo listar hijos: {_de}")
            return {"ok": False, "pos": "", "error": "ZPPR0008: grid no encontrado tras F8 — ver [DIAG] en log"}

        try:
            n = grid.RowCount
            print(f"    ZPPR0008: {n} filas")
            # Imprimir columnas disponibles en fila 0 para diagnóstico
            if n > 0:
                try:
                    _cols = grid.ColumnOrder
                    print(f"    [DIAG] ZPPR0008 columnas: {list(_cols)}")
                except Exception:
                    pass
            for i in range(n):
                for col in ("POSNR", "POSN", "POS", "COMPONENT", "MATNR"):
                    try:
                        v = str(grid.GetCellValue(i, col) or "").strip().lstrip("0") or "0"
                        if v in ("106", "116"):
                            pos_found = "0" + v
                            print(f"    ZPPR0008: posición acero encontrada = {pos_found}")
                            return {"ok": True, "pos": pos_found, "error": ""}
                        break
                    except Exception:
                        pass
            return {"ok": False, "pos": "",
                    "error": f"ZFER {zfer_base} no tiene posición 0106 ni 0116 → no aplica cambio de fórmula sin acero"}
        except Exception as e:
            return {"ok": False, "pos": "", "error": f"ZPPR0008 error grid: {e}"}

    # ── ZMME0001 — Cambio de Fórmula ─────────────────────────────────────────

    _ID_RAD_FORMULA = "wnd[0]/usr/radRB2_A1"
    _ID_TXT_FORMU   = "wnd[0]/usr/txtP_FORMU"

    def zmme0001_ejecutar_formula(self, zfer_base: str, p_color: str, p_franj: str,
                                  formula_nueva: str, zplas_validos: list,
                                  forzar_be: bool = False) -> tuple:
        """
        Igual que zmme0001_ejecutar pero selecciona 'Cambio de Fórmula' (radRB2_A1)
        y llena el campo txtP_FORMU con la fórmula destino.
        """
        self._cerrar_dialogs_abiertos()
        self._navegar("ZMME0001")
        self.session.findById("wnd[0]").maximize()
        self._esperar(T_RAPIDO)

        self.session.findById(self._ID_RAD_HOMOLOG).setFocus()
        self.session.findById(self._ID_RAD_HOMOLOG).select()
        self._esperar(T_RAPIDO)

        self.session.findById(self._ID_MATER_LOW).text = zfer_base
        self._esperar(T_RAPIDO)
        self.session.findById(self._ID_CTX_CENTER).text = "CO01"

        # Cambio de Fórmula
        self.session.findById(self._ID_RAD_FORMULA).setFocus()
        self.session.findById(self._ID_RAD_FORMULA).select()
        self._esperar(T_RAPIDO)

        self.session.findById(self._ID_CTX_P_COLOR).text = p_color
        self.session.findById(self._ID_CTX_P_FRANJ).text = p_franj
        self.session.findById(self._ID_TXT_FORMU).text   = formula_nueva

        # F4 en ZPLA (idéntico al flujo de color)
        self.session.findById(self._ID_CTX_P_ZPLA).setFocus()
        self.session.findById(self._ID_CTX_P_ZPLA).caretPosition = 0
        self.session.findById("wnd[0]").sendVKey(4)
        self._esperar(T_LENTO)

        _ID_POPUP_GRID = "wnd[1]/usr/cntlLO_CONTAINER0500/shellcont/shell"
        grid_popup = None
        for _intento in range(3):
            try:
                grid_popup = self.session.findById(_ID_POPUP_GRID)
                break
            except Exception:
                if _intento < 2:
                    print(f"    [F4-formula] popup no listo, esperando... ({_intento+1})")
                    self._esperar(T_LENTO)
        if grid_popup is None:
            raise RuntimeError("F4 ZPLA popup falló en cambio de fórmula")

        try:
            n_filas   = grid_popup.RowCount
            zplas_sap = []
            descs_sap = []
            for i in range(n_filas):
                val = ""
                for col in ("MATNR", "ZPLA", "ZPLARF", "MATERIAL"):
                    try:
                        val = str(grid_popup.GetCellValue(i, col) or "").strip()
                        if val:
                            break
                    except Exception:
                        pass
                desc = ""
                for col in ("MAKTX", "DESCR", "TEXT", "MAKTG", "BEZEICHNUNG"):
                    try:
                        desc = str(grid_popup.GetCellValue(i, col) or "").strip()
                        if desc:
                            break
                    except Exception:
                        pass
                zplas_sap.append(val)
                descs_sap.append(desc)
                print(f"      F4-formula fila {i}: {val}  desc='{desc}'")

            zplas_set = {z.strip() for z in zplas_validos if z.strip()}
            fila_sel  = 0
            if forzar_be:
                fila_be = next((i for i, d in enumerate(descs_sap) if "BE" in d.upper()), None)
                fila_sel = fila_be if fila_be is not None else 0
            else:
                for i, z in enumerate(zplas_sap):
                    if z in zplas_set:
                        fila_sel = i
                        break

            zpla_seleccionado = zplas_sap[fila_sel] if zplas_sap else ""
            grid_popup.selectedRows = str(fila_sel)
            grid_popup.doubleClickCurrentCell()
            self._esperar(T_RAPIDO)
        except Exception as e:
            raise RuntimeError(f"F4 ZPLA popup (fórmula) falló: {e}")

        self.session.findById(self._ID_CTX_P_FRANJ).setFocus()
        self.session.findById(self._ID_CTX_P_FRANJ).caretPosition = 2
        self._esperar(T_RAPIDO)
        self.session.findById(self._ID_BTN_EXEC).press()
        self._esperar(T_LENTO)

        msg_sap = self._estado_sap()
        try:
            grid       = self.session.findById(self._ID_GRID_RESULT)
            n_filas    = grid.RowCount
            # Debug: dump columnas reales del grid resultado
            try:
                co = grid.ColumnOrder
                all_cols = co.split() if isinstance(co, str) else [str(c) for c in co]
                print(f"    [DEBUG] ZMME0001-formula grid cols: {all_cols} | filas={n_filas}")
            except Exception:
                all_cols = []
            zfer_nuevo = ""
            zfor_nuevo = ""
            # Leer fila 0 (la primera y única fila del resultado)
            for col in ("ZFER", "MATNR_ZFER", "ZFER_NEW", "MATNR", "MAT_ZFER"):
                try:
                    v = str(grid.GetCellValue(0, col) or "").strip()
                    if v and v != col:   # descartar cuando SAP devuelve el nombre de columna
                        zfer_nuevo = v
                        break
                except Exception:
                    pass
            for col in ("ZFOR", "MATNR_ZFOR", "ZFOR_NEW", "MAT_ZFOR"):
                try:
                    v = str(grid.GetCellValue(0, col) or "").strip()
                    if v and v != col:
                        zfor_nuevo = v
                        break
                except Exception:
                    pass
            print(f"    ZMME0001-formula OK: ZFER_NUEVO={zfer_nuevo} | ZFOR={zfor_nuevo} | ZPLA={zpla_seleccionado}")
        except Exception as e:
            raise RuntimeError(f"ZMME0001-formula: grid resultado no leído: {e} | msg={msg_sap}")

        # F3 para volver a pantalla de selección (igual que cambio de color, necesario para paso 4)
        self.session.findById("wnd[0]").sendVKey(3)
        self._esperar(T_RAPIDO)

        return zfer_nuevo, zfor_nuevo, zpla_seleccionado

    # ── MM02 — Desactivar diferencial 06 ─────────────────────────────────────

    _ID_MM02_TBL_PIEZA = ("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000"
                           "/tabsTABSTRIP_CHAR/tabpTAB4"
                           "/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100"
                           "/tblSAPLCTMSCHARS_S")

    def mm02_desactivar_diferencial_06(self, zfer: str):
        """
        Navega a MM02, abre PIEZA tab, scroll a pos 6, fila visual 7 = Z_BEHAVIOR_DIFFERENTIALS.
        Abre popup con sendVKey(2), desmarca fila 5 (valor "06"), guarda y sale.
        IDs confirmados por VBS grabado en QUAS.
        """
        print(f"    MM02 diferencial: desmarcando 06 en {zfer}")
        
        # Navegar a MM02 y abrir tab PIEZA (igual que mm02_actualizar_partnumber)
        self._cerrar_dialogs_abiertos()
        self.session.findById(self._ID_TCODE_BOX).text = "/NMM02"
        self.session.findById("wnd[0]").sendVKey(0)
        self._esperar(T_MEDIO)
        self.session.findById(self._ID_MM02_MATNR).text = zfer
        self.session.findById("wnd[0]").sendVKey(0)
        self._esperar(T_MEDIO)
        for _ in range(3):
            try:
                self.session.findById(self._ID_MM02_TAB03)
                break
            except Exception:
                pass
            if not self._primera_opcion_si_popup():
                self.session.findById("wnd[0]").sendVKey(0)
                self._esperar(T_MEDIO)
        self.session.findById(self._ID_MM02_TAB03).select()
        self._esperar(T_RAPIDO)
        self._primera_opcion_si_popup()
        self.session.findById(self._ID_MM02_TAB4).select()
        self._esperar(T_RAPIDO)

        tbl = self._ID_MM02_TBL_PIEZA
        # Scroll a posición 6 (confirma VBS)
        self.session.findById(tbl).verticalScrollbar.position = 6
        self._esperar(T_RAPIDO)

        # Fila visual 7, columna MWERT (1) = Z_BEHAVIOR_DIFFERENTIALS (confirma VBS)
        campo_name = tbl + "/ctxtRCTMS-MWERT[1,7]"
        self.session.findById(campo_name).setFocus()
        self.session.findById(campo_name).caretPosition = 8
        self.session.findById("wnd[0]").sendVKey(2)   # abre popup de valores
        self._esperar(T_MEDIO)

        # En popup wnd[1]: desmarcar checkbox fila 5 (el "06") → btn[8] para confirmar
        # Confirmado por VBS: selected=False → setFocus → btn[8].press()
        chk = "wnd[1]/usr/tblSAPLCTMSVALUE_S/chkRCTMS-SEL01[0,5]"
        try:
            self.session.findById(chk).selected = False
            self.session.findById(chk).setFocus()
            self._esperar(T_RAPIDO)
            self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
        except Exception as e:
            print(f"    [WARN] Diferencial popup check: {e}")
        self._esperar(T_RAPIDO)

        # Guardar con btn[11] y salir con F3 → OPTION1
        try:
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self._esperar(T_MEDIO)
            self._primera_opcion_si_popup()
        except Exception as e:
            print(f"    [WARN] Diferencial guardar: {e}")
        try:
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self._esperar(T_RAPIDO)
            self._primera_opcion_si_popup()
        except Exception:
            pass

    # ── MM02 — Cambio de plano (tab ZU04) ────────────────────────────────────

    def _plano_base(self, doknr: str) -> str:
        """
        Extrae el núcleo de un DOKNR quitando desde la derecha cualquier
        combinación de letras cortas (versión A/B/AA…) y 'SP'.
        Ej:  "M1344 000 001 A SP" → "M1344 000 001"
             "M1344 000 001 SP"   → "M1344 000 001"
             "M1344 000 001 A"    → "M1344 000 001"
             "M1344 000 001"      → "M1344 000 001"
        """
        return re.sub(r'(\s+[A-Za-z]{1,3})+$', '', doknr.strip()).strip()

    def _buscar_plano_bd(self, doknr_actual: str) -> tuple:
        """
        Dado el DOKNR leído de MM02 (puede tener SP y/o letra al final),
        busca en ODATA_ZFER_RUTAS_JPG el PLANO más reciente sin SP,
        ordenado por ULTIMA_MOD DESC para siempre tomar la versión vigente.
        Returns: (plano_nuevo: str | None, mensaje: str)
        """
        base = self._plano_base(doknr_actual)
        if not base:
            return None, f"DOKNR '{doknr_actual}' no tiene base reconocible"

        try:
            cn  = pyodbc.connect(_DB_SAP_STR, autocommit=True)
            cur = cn.cursor()
            # TOP 1 con NOT LIKE '% SP' excluye SP y "X SP" en un solo query,
            # ordenado por ULTIMA_MOD DESC → siempre la revisión más reciente
            cur.execute(
                "SELECT TOP 1 DOCUMENTO, PLANO "
                "FROM dbo.ODATA_ZFER_RUTAS_JPG "
                "WHERE DOCUMENTO LIKE ? "
                "  AND DOCUMENTO NOT LIKE '% SP' "
                "ORDER BY ULTIMA_MOD DESC",
                f"%{base}%"
            )
            row = cur.fetchone()
            cn.close()
        except Exception as e:
            return None, f"Error BD al buscar plano: {e}"

        if not row:
            return None, (
                f"⚠ PLANO NO ACTUALIZADO — No se encontró ninguna versión sin SP "
                f"para '{base}' en ODATA_ZFER_RUTAS_JPG. "
                f"Revisa que el plano exista en la BD o que no tenga SP en todas sus versiones."
            )

        doc_elegido = str(row[0] or "").strip()

        if not doc_elegido:
            return None, (
                f"⚠ PLANO NO ACTUALIZADO — Se encontró fila pero columna DOCUMENTO está vacía en BD."
            )

        # Se devuelve el DOCUMENTO (nombre del plano) para escribir en el campo DOKNR de MM02.
        # Ej: "M0606 065 001" — no la ruta del archivo (columna PLANO).
        return doc_elegido, f"Plano actualizado: DOKNR ← '{doc_elegido}'"

    def _buscar_doknr_por_material(self, zfer: str, con_sp: bool = False) -> tuple:
        """
        Busca en ODATA_ZFER_RUTAS_JPG el DOCUMENTO (DOKNR) más reciente para un MATERIAL.
        con_sp=False → busca sin SP (flujo con→sin acero)
        con_sp=True  → busca con SP   (flujo sin→con acero)
        Returns: (doknr: str | None, mensaje: str)
        """
        try:
            cn  = pyodbc.connect(_DB_SAP_STR, autocommit=True)
            cur = cn.cursor()
            if con_sp:
                cur.execute(
                    "SELECT TOP 1 DOCUMENTO FROM dbo.ODATA_ZFER_RUTAS_JPG "
                    "WHERE MATERIAL = ? AND CENTRO = 'CO01' AND DOCUMENTO LIKE '% SP' "
                    "ORDER BY ULTIMA_MOD DESC",
                    zfer
                )
            else:
                cur.execute(
                    "SELECT TOP 1 DOCUMENTO FROM dbo.ODATA_ZFER_RUTAS_JPG "
                    "WHERE MATERIAL = ? AND CENTRO = 'CO01' AND DOCUMENTO NOT LIKE '% SP' "
                    "ORDER BY ULTIMA_MOD DESC",
                    zfer
                )
            row = cur.fetchone()
            cn.close()
        except Exception as e:
            return None, f"Error BD buscando DOKNR por material: {e}"

        if not row or not row[0]:
            sp_txt = "con SP" if con_sp else "sin SP"
            return None, f"⚠ Sin plano {sp_txt} en BD para material {zfer}"

        return str(row[0]).strip(), f"DOKNR BD ({('con SP' if con_sp else 'sin SP')}): '{row[0].strip()}'"

    def mm02_cambiar_plano(self, zfer: str, res: "ResultadoItem" = None,
                           zfer_base: str = None) -> bool:
        """
        Lee el DOKNR del ZFER BASE en MM02 (que sí tiene ZU04 extendido),
        busca en BD el plano sin SP con ese nombre, y lo escribe en el ZFER nuevo.
        Retorna True si guardó, False si omitió (sin romper el flujo).
        """
        def _warn(msg):
            print(f"    [WARN] mm02_cambiar_plano: {msg}")
            if res:
                res._log(f"  [PLANO] ADVERTENCIA: {msg}")

        zfer_lectura = zfer_base if zfer_base else zfer
        print(f"    MM02 plano: leyendo DOKNR de {zfer_lectura}, escribiendo en {zfer}")
        try:
            _subZU04 = ("wnd[0]/usr/tabsTABSPR1/tabpZU04"
                        "/ssubTABFRA1:SAPLMGMM:2110"
                        "/subSUB2:SAPLMGD1:3400"
                        "/subDOCU:SAPLCV140:0204")
            _grid_docu = _subZU04 + "/subDOC_ALV:SAPLCV140:0206/cntlALV_CUST_DOC/shellcont/shell"

            # ── 1. Leer DOKNR del ZFER BASE desde SAP (primario) ─────────────
            doknr_base = ""
            try:
                self.session.findById(self._ID_TCODE_BOX).text = "/nmm02"
                self.session.findById("wnd[0]").sendVKey(0)
                self._esperar(T_MEDIO)
                self.session.findById(self._ID_MM02_MATNR).text = zfer_lectura
                self.session.findById(self._ID_MM02_MATNR).caretPosition = len(zfer_lectura)
                self.session.findById("wnd[0]").sendVKey(0)
                self._esperar(T_MEDIO)
                self.session.findById("wnd[0]").sendVKey(0)
                self._esperar(T_MEDIO)
                self.session.findById("wnd[0]/tbar[1]/btn[30]").press()
                self._esperar(T_MEDIO)
                self.session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU04").select()
                self._esperar(T_MEDIO)
                self.session.findById(_subZU04 + "/subBUTTON:SAPLCV140:0203/radGF_ALLE").setFocus()
                self.session.findById(_subZU04 + "/subBUTTON:SAPLCV140:0203/radGF_ALLE").select()
                self._esperar(T_RAPIDO)
                doknr_base = str(self.session.findById(_grid_docu).getCellValue(0, "DOKNR") or "").strip()
                print(f"    MM02 plano: DOKNR SAP base='{doknr_base}'")
            except Exception as e_sap:
                print(f"    MM02 plano: SAP no disponible ({e_sap}), intentando BD…")

            # Fallback: BD si SAP no devolvió nada
            if not doknr_base:
                doknr_base_bd, msg_bd0 = self._buscar_doknr_por_material(zfer_lectura, con_sp=False)
                print(f"    MM02 plano BD fallback: {msg_bd0}")
                if res:
                    res._log(f"  [PLANO] BD fallback: {msg_bd0}")
                if not doknr_base_bd:
                    _warn(f"Sin DOKNR en SAP ni BD para {zfer_lectura} — omitiendo cambio de plano")
                    return False
                doknr_base = doknr_base_bd

            # ── 2. Buscar en BD el plano sin SP usando el nombre base ─────────
            nuevo_plano, msg_bd = self._buscar_plano_bd(doknr_base)
            print(f"    MM02 plano BD: {msg_bd}")
            if res:
                res._log(f"  [PLANO] {msg_bd}")
            if not nuevo_plano:
                return False

            # ── 3. Navegar al ZFER NUEVO y escribir el plano ──────────────────
            self.session.findById(self._ID_TCODE_BOX).text = "/nmm02"
            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)
            self.session.findById(self._ID_MM02_MATNR).text = zfer
            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)
            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)

            try:
                self.session.findById("wnd[0]/tbar[1]/btn[30]").press()
            except Exception:
                _warn(f"btn[30] no disponible en ZFER nuevo {zfer} — omitiendo cambio de plano")
                return False
            self._esperar(T_MEDIO)
            self.session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU04").select()
            self._esperar(T_MEDIO)
            self.session.findById(_subZU04 + "/subBUTTON:SAPLCV140:0203/radGF_ALLE").setFocus()
            self.session.findById(_subZU04 + "/subBUTTON:SAPLCV140:0203/radGF_ALLE").select()
            self._esperar(T_RAPIDO)

            try:
                doknr_actual = str(self.session.findById(_grid_docu).getCellValue(0, "DOKNR") or "").strip()
            except Exception:
                doknr_actual = ""

            if nuevo_plano == doknr_actual:
                print(f"    MM02 plano: sin cambio necesario ('{nuevo_plano}')")
                if res:
                    res._log(f"  [PLANO] Sin cambio necesario ('{nuevo_plano}')")
                return True

            self.session.findById(_grid_docu).modifyCell(0, "DOKNR", nuevo_plano)
            self.session.findById(_grid_docu).currentCellColumn = "DOKNR"
            self.session.findById(_grid_docu).pressEnter()
            self._esperar(T_MEDIO)
            try:
                self.session.findById("wnd[1]/usr/btnBUTTON_1").press()
                self._esperar(T_RAPIDO)
            except Exception:
                pass
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self._esperar(T_LENTO)
            print(f"    MM02 plano guardado: {zfer} → '{nuevo_plano}'")
            if res:
                res._log(f"  [PLANO] Guardado: '{doknr_actual or '(vacío)'}' → '{nuevo_plano}'")
            return True
        except Exception as e:
            _warn(str(e))
            return False

    # ── ZPPR0008 — Validar posición acero por ZPLA (sesión auxiliar) ─────────

    def zppr0008_validar_posicion_acero_zpla(self, zpla: str) -> dict:
        """
        Abre sesión auxiliar, entra a ZPPR0008 modo ZPLA (radRB_2),
        filtra por ZPLA+CO01, busca posición 0106 ó 0116.
        Retorna {"ok": True/False, "pos": "0106"|"0116"|"", "error": ""}
        Si ok=False con pos="" → el ZPLA NO tiene acero → abortar flujo con acero.
        """
        print(f"    ZPPR0008 (aux): validando posición acero para ZPLA={zpla}")
        self.session.createSession()
        self._esperar(T_LENTO)

        idx_nueva = self.conn_sap.Children.Count - 1
        ses = self.conn_sap.Children(idx_nueva)
        ses.findById("wnd[0]").maximize()

        try:
            ses.findById(self._ID_TCODE_BOX).text = "ZPPR0008"
            ses.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)

            # Modo "por material" con radio radRB_2 + campo ctxtS_MATNR2-LOW = ZPLA
            ses.findById("wnd[0]/usr/radRB_2").setFocus()
            ses.findById("wnd[0]/usr/radRB_2").select()
            self._esperar(T_RAPIDO)
            ses.findById("wnd[0]/usr/ctxtS_MATNR2-LOW").text = zpla
            ses.findById("wnd[0]/usr/ctxtS_WERKS2-LOW").text = "CO01"
            ses.findById(self._ID_BTN_EXEC).press()
            self._esperar(T_LENTO)

            # Buscar grid
            grid = None
            for _gid in ("wnd[0]/usr/cntlGRID1/shellcont/shell",
                         "wnd[0]/usr/cntlGRID/shellcont/shell",
                         "wnd[0]/usr/cntlALV/shellcont/shell"):
                try:
                    grid = ses.findById(_gid)
                    break
                except Exception:
                    pass

            if grid is None:
                return {"ok": False, "pos": "",
                        "error": f"ZPPR0008 (aux): grid no encontrado para ZPLA={zpla}"}

            n = grid.RowCount
            print(f"    ZPPR0008 (aux): {n} filas para ZPLA={zpla}")
            for i in range(n):
                for col in ("POSNR", "POSN", "POS", "COMPONENT", "MATNR"):
                    try:
                        v = str(grid.GetCellValue(i, col) or "").strip().lstrip("0") or "0"
                        if v in ("106", "116"):
                            pos_found = "0" + v
                            print(f"    ZPPR0008 (aux): posición acero encontrada = {pos_found}")
                            return {"ok": True, "pos": pos_found, "error": ""}
                        break
                    except Exception:
                        pass
            return {"ok": False, "pos": "",
                    "error": f"ZPLA {zpla} no tiene posición 0106 ni 0116 — no aplica flujo con acero"}

        except Exception as e:
            return {"ok": False, "pos": "", "error": f"ZPPR0008 (aux) error: {e}"}
        finally:
            try:
                ses.findById("wnd[0]").close()
            except Exception:
                pass

    # ── MM02 — Activar diferencial 06 (inverso de desactivar) ────────────────

    def mm02_activar_diferencial_06(self, zfer: str):
        """
        Igual que mm02_desactivar_diferencial_06 pero marca selected=True (activa "06").
        IDs idénticos confirmados por VBS.
        """
        print(f"    MM02 diferencial: activando 06 en {zfer}")
        self._cerrar_dialogs_abiertos()
        self.session.findById(self._ID_TCODE_BOX).text = "/NMM02"
        self.session.findById("wnd[0]").sendVKey(0)
        self._esperar(T_MEDIO)
        self.session.findById(self._ID_MM02_MATNR).text = zfer
        self.session.findById("wnd[0]").sendVKey(0)
        self._esperar(T_MEDIO)
        for _ in range(3):
            try:
                self.session.findById(self._ID_MM02_TAB03)
                break
            except Exception:
                pass
            if not self._primera_opcion_si_popup():
                self.session.findById("wnd[0]").sendVKey(0)
                self._esperar(T_MEDIO)
        self.session.findById(self._ID_MM02_TAB03).select()
        self._esperar(T_RAPIDO)
        self._primera_opcion_si_popup()
        self.session.findById(self._ID_MM02_TAB4).select()
        self._esperar(T_RAPIDO)

        tbl = self._ID_MM02_TBL_PIEZA
        self.session.findById(tbl).verticalScrollbar.position = 6
        self._esperar(T_RAPIDO)

        campo_name = tbl + "/ctxtRCTMS-MWERT[1,7]"
        self.session.findById(campo_name).setFocus()
        self.session.findById(campo_name).caretPosition = 8
        self.session.findById("wnd[0]").sendVKey(2)
        self._esperar(T_MEDIO)

        chk = "wnd[1]/usr/tblSAPLCTMSVALUE_S/chkRCTMS-SEL01[0,5]"
        try:
            self.session.findById(chk).selected = True
            self.session.findById(chk).setFocus()
            self._esperar(T_RAPIDO)
            self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
        except Exception as e:
            print(f"    [WARN] Activar diferencial popup check: {e}")
        self._esperar(T_RAPIDO)

        try:
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self._esperar(T_MEDIO)
            self._primera_opcion_si_popup()
        except Exception as e:
            print(f"    [WARN] Activar diferencial guardar: {e}")
        try:
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self._esperar(T_RAPIDO)
            self._primera_opcion_si_popup()
        except Exception:
            pass

    # ── MM02 — Cambio de plano CON SP (flujo sin acero → con acero) ──────────

    def _buscar_plano_con_sp(self, doknr_actual: str) -> tuple:
        """
        Dado el DOKNR leído de MM02, busca en ODATA_ZFER_RUTAS_JPG el plano
        más reciente QUE SÍ tenga SP, ordenado por ULTIMA_MOD DESC.
        Returns: (plano_nuevo: str | None, mensaje: str)
        """
        base = self._plano_base(doknr_actual)
        if not base:
            return None, f"DOKNR '{doknr_actual}' no tiene base reconocible"

        try:
            cn  = pyodbc.connect(_DB_SAP_STR, autocommit=True)
            cur = cn.cursor()
            cur.execute(
                "SELECT TOP 1 DOCUMENTO "
                "FROM dbo.ODATA_ZFER_RUTAS_JPG "
                "WHERE DOCUMENTO LIKE ? "
                "  AND DOCUMENTO LIKE '% SP' "
                "ORDER BY ULTIMA_MOD DESC",
                f"%{base}%"
            )
            row = cur.fetchone()
            cn.close()
        except Exception as e:
            return None, f"Error BD al buscar plano con SP: {e}"

        if not row:
            return None, (
                f"⚠ PLANO CON SP NO ENCONTRADO — No se encontró versión con SP "
                f"para '{base}' en ODATA_ZFER_RUTAS_JPG."
            )

        doc_elegido = str(row[0] or "").strip()
        if not doc_elegido:
            return None, "⚠ PLANO CON SP NO ENCONTRADO — DOCUMENTO vacío en BD."
        return doc_elegido, f"Plano con SP: DOKNR ← '{doc_elegido}'"

    def mm02_cambiar_plano_con_sp(self, zfer: str, res: "ResultadoItem" = None,
                                  zfer_base: str = None) -> bool:
        """
        Lee el DOKNR del ZFER BASE en MM02, busca en BD el plano CON SP con ese nombre,
        y lo escribe en el ZFER nuevo. Para flujo sin acero → con acero.
        """
        def _warn(msg):
            print(f"    [WARN] mm02_cambiar_plano_con_sp: {msg}")
            if res:
                res._log(f"  [PLANO-SP] ADVERTENCIA: {msg}")

        zfer_lectura = zfer_base if zfer_base else zfer
        print(f"    MM02 plano (con SP): leyendo DOKNR de {zfer_lectura}, escribiendo en {zfer}")
        try:
            _subZU04 = ("wnd[0]/usr/tabsTABSPR1/tabpZU04"
                        "/ssubTABFRA1:SAPLMGMM:2110"
                        "/subSUB2:SAPLMGD1:3400"
                        "/subDOCU:SAPLCV140:0204")
            _grid_docu = _subZU04 + "/subDOC_ALV:SAPLCV140:0206/cntlALV_CUST_DOC/shellcont/shell"

            # ── 1. Leer DOKNR del ZFER BASE desde SAP (primario) ─────────────
            doknr_base = ""
            try:
                self.session.findById(self._ID_TCODE_BOX).text = "/nmm02"
                self.session.findById("wnd[0]").sendVKey(0)
                self._esperar(T_MEDIO)
                self.session.findById(self._ID_MM02_MATNR).text = zfer_lectura
                self.session.findById(self._ID_MM02_MATNR).caretPosition = len(zfer_lectura)
                self.session.findById("wnd[0]").sendVKey(0)
                self._esperar(T_MEDIO)
                self.session.findById("wnd[0]").sendVKey(0)
                self._esperar(T_MEDIO)
                self.session.findById("wnd[0]/tbar[1]/btn[30]").press()
                self._esperar(T_MEDIO)
                self.session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU04").select()
                self._esperar(T_MEDIO)
                self.session.findById(_subZU04 + "/subBUTTON:SAPLCV140:0203/radGF_ALLE").setFocus()
                self.session.findById(_subZU04 + "/subBUTTON:SAPLCV140:0203/radGF_ALLE").select()
                self._esperar(T_RAPIDO)
                doknr_base = str(self.session.findById(_grid_docu).getCellValue(0, "DOKNR") or "").strip()
                print(f"    MM02 plano (con SP): DOKNR SAP base='{doknr_base}'")
            except Exception as e_sap:
                print(f"    MM02 plano (con SP): SAP no disponible ({e_sap}), intentando BD…")

            # Fallback: BD si SAP no devolvió nada
            if not doknr_base:
                doknr_base_bd, msg_bd0 = self._buscar_doknr_por_material(zfer_lectura, con_sp=True)
                print(f"    MM02 plano (con SP) BD fallback: {msg_bd0}")
                if res:
                    res._log(f"  [PLANO-SP] BD fallback: {msg_bd0}")
                if not doknr_base_bd:
                    _warn(f"Sin DOKNR en SAP ni BD para {zfer_lectura} — omitiendo cambio de plano")
                    return False
                doknr_base = doknr_base_bd

            # ── 2. Buscar en BD el plano CON SP usando el nombre base ─────────
            nuevo_plano, msg_bd = self._buscar_plano_con_sp(doknr_base)
            print(f"    MM02 plano (con SP) BD: {msg_bd}")
            if res:
                res._log(f"  [PLANO-SP] {msg_bd}")
            if not nuevo_plano:
                return False

            # ── 3. Navegar al ZFER NUEVO y escribir el plano ──────────────────
            self.session.findById(self._ID_TCODE_BOX).text = "/nmm02"
            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)
            self.session.findById(self._ID_MM02_MATNR).text = zfer
            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)
            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)

            try:
                self.session.findById("wnd[0]/tbar[1]/btn[30]").press()
            except Exception:
                _warn(f"btn[30] no disponible en ZFER nuevo {zfer} — omitiendo cambio de plano")
                return False
            self._esperar(T_MEDIO)
            self.session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU04").select()
            self._esperar(T_MEDIO)
            self.session.findById(_subZU04 + "/subBUTTON:SAPLCV140:0203/radGF_ALLE").setFocus()
            self.session.findById(_subZU04 + "/subBUTTON:SAPLCV140:0203/radGF_ALLE").select()
            self._esperar(T_RAPIDO)

            try:
                doknr_actual = str(self.session.findById(_grid_docu).getCellValue(0, "DOKNR") or "").strip()
            except Exception:
                doknr_actual = ""

            if nuevo_plano == doknr_actual:
                print(f"    MM02 plano (con SP): sin cambio necesario ('{nuevo_plano}')")
                if res:
                    res._log(f"  [PLANO-SP] Sin cambio necesario ('{nuevo_plano}')")
                return True

            self.session.findById(_grid_docu).modifyCell(0, "DOKNR", nuevo_plano)
            self.session.findById(_grid_docu).currentCellColumn = "DOKNR"
            self.session.findById(_grid_docu).pressEnter()
            self._esperar(T_MEDIO)
            try:
                self.session.findById("wnd[1]/usr/btnBUTTON_1").press()
                self._esperar(T_RAPIDO)
            except Exception:
                pass
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self._esperar(T_LENTO)
            print(f"    MM02 plano (con SP) guardado: {zfer} → '{nuevo_plano}'")
            if res:
                res._log(f"  [PLANO-SP] Guardado: '{doknr_actual or '(vacío)'}' → '{nuevo_plano}'")
            return True
        except Exception as e:
            _warn(str(e))
            return False

    # ── CA02 — Cambio de Hoja de Ruta ────────────────────────────────────────

    def _ca02_leer_matnr_vis(self, tbl_path: str, vis_row: int) -> str:
        """Lee MATNR de fila visual vis_row via findById (getCellValue no funciona en esta tabla)."""
        try:
            return str(self.session.findById(f"{tbl_path}/ctxtMAPL-MATNR[2,{vis_row}]").text or "").strip()
        except Exception:
            return ""

    def _ca02_scroll(self, tbl, pos: int):
        try:
            tbl.verticalScrollbar.position = pos
            self._esperar(T_RAPIDO)
        except Exception:
            pass

    def ca02_desasignar_hr(self, zfer_nuevo: str, res=None) -> bool:
        """
        CA02 con MATNR=zfer_nuevo → busca la asignación de HR → la borra → guarda.
        Lee celdas via findById por posición de scroll (getCellValue no funciona aquí).
        """
        def _warn(msg):
            print(f"    [WARN] ca02_desasignar: {msg}")
            if res: res._advertir(f"HR-Desasignar: {msg}")

        print(f"    CA02 desasignar HR: {zfer_nuevo}")
        _TBL = "wnd[1]/usr/tblSAPLCZDITCTRL_1010"
        try:
            self.session.findById(self._ID_TCODE_BOX).text = "/nca02"
            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)
            # Desasignar: solo MATNR, PLNNR debe estar vacío
            self.session.findById("wnd[0]/usr/ctxtRC271-PLNNR").text = ""
            self.session.findById("wnd[0]/usr/ctxtRC27M-MATNR").text = zfer_nuevo
            self.session.findById("wnd[0]/usr/ctxtRC27M-WERKS").text = "CO01"
            self.session.findById("wnd[0]/usr/ctxtRC27M-WERKS").caretPosition = 4
            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)
            self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
            self._esperar(T_MEDIO)
            self.session.findById("wnd[0]/tbar[1]/btn[31]").press()
            self._esperar(T_MEDIO)

            # Si la tabla no existe o la HR no fue asignada, btn[31] puede no abrir popup
            try:
                tbl        = self.session.findById(_TBL)
                vis_rows   = tbl.VisibleRowCount
                max_scroll = tbl.verticalScrollbar.maximum
            except Exception:
                _warn(f"Sin HR asignada para {zfer_nuevo} (popup no disponible — esperado para ZFER nuevo)")
                try: self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                except Exception: pass
                return False

            # Escanear via findById en cada posición de scroll
            fila_scroll = None
            fila_vis    = None
            for sp in range(max_scroll + 1):
                self._ca02_scroll(tbl, sp)
                for vis in range(vis_rows):
                    if self._ca02_leer_matnr_vis(_TBL, vis) == zfer_nuevo:
                        fila_scroll = sp
                        fila_vis    = vis
                        break
                if fila_scroll is not None:
                    break

            if fila_scroll is None:
                _warn(f"No se encontró asignación de HR para {zfer_nuevo} en CA02")
                try: self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                except Exception: pass
                return False

            print(f"    CA02 desasignar: scroll={fila_scroll} vis_row={fila_vis}")
            self.session.findById(f"{_TBL}/ctxtMAPL-MATNR[2,{fila_vis}]").setFocus()
            self.session.findById(f"{_TBL}/ctxtMAPL-MATNR[2,{fila_vis}]").caretPosition = 9
            self._esperar(T_RAPIDO)
            self.session.findById("wnd[1]/tbar[0]/btn[14]").press()
            self._esperar(T_RAPIDO)
            for _id in ("wnd[2]/tbar[0]/btn[0]", "wnd[2]/usr/btnSPOP-OPTION1"):
                try:
                    self.session.findById(_id).press()
                    self._esperar(T_RAPIDO)
                except Exception:
                    pass
            try:
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                self._esperar(T_RAPIDO)
            except Exception:
                pass
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self._esperar(T_LENTO)
            print(f"    CA02 desasignación guardada: {zfer_nuevo}")
            if res: res._log(f"  [HR-DESASIGNAR] OK: {zfer_nuevo}")
            return True
        except Exception as e:
            _warn(str(e))
            return False

    def ca02_asignar_hr(self, zfer_nuevo: str, id_hruta: str, res=None) -> bool:
        """
        CA02 con PLNNR=id_hruta → abre popup de materiales → escribe zfer_nuevo en
        la primera fila vacía → guarda.
        Lee celdas via findById por posición de scroll (getCellValue no funciona aquí).
        """
        def _warn(msg):
            print(f"    [WARN] ca02_asignar: {msg}")
            if res: res._advertir(f"HR-Asignar: {msg}")

        print(f"    CA02 asignar HR {id_hruta} → {zfer_nuevo}")
        _TBL = "wnd[1]/usr/tblSAPLCZDITCTRL_1010"
        try:
            self.session.findById(self._ID_TCODE_BOX).text = "/nca02"
            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)
            # Asignar: solo PLNNR, MATNR debe estar vacío
            self.session.findById("wnd[0]/usr/ctxtRC27M-MATNR").text = ""
            self.session.findById("wnd[0]/usr/ctxtRC271-PLNNR").text = str(id_hruta)
            self.session.findById("wnd[0]/usr/ctxtRC27M-WERKS").text = "CO01"
            self.session.findById("wnd[0]/usr/ctxtRC27M-WERKS").caretPosition = 4
            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)
            self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
            self._esperar(T_MEDIO)
            self.session.findById("wnd[0]/tbar[1]/btn[31]").press()
            self._esperar(T_MEDIO)

            tbl = self.session.findById(_TBL)
            vis_rows   = tbl.VisibleRowCount
            max_scroll = tbl.verticalScrollbar.maximum

            # Buscar primera fila vacía escaneando desde el final via findById
            fila_scroll = None
            fila_vis    = None
            for sp in range(max_scroll, -1, -1):
                self._ca02_scroll(tbl, sp)
                for vis in range(vis_rows - 1, -1, -1):
                    val = self._ca02_leer_matnr_vis(_TBL, vis)
                    if not val:
                        fila_scroll = sp
                        fila_vis    = vis
                    else:
                        if fila_scroll is not None:
                            break
                if fila_scroll is not None:
                    break

            if fila_scroll is None:
                _warn("No se encontró fila vacía en la tabla de materiales de CA02")
                try: self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                except Exception: pass
                return False

            print(f"    CA02 asignar: scroll={fila_scroll} vis_row={fila_vis}")
            _matnr = f"{_TBL}/ctxtMAPL-MATNR[2,{fila_vis}]"
            _werks  = f"{_TBL}/ctxtMAPL-WERKS[3,{fila_vis}]"
            _plnal  = f"{_TBL}/txtMAPL-PLNAL[0,{fila_vis}]"
            self.session.findById(_matnr).setFocus()
            self._esperar(T_RAPIDO)
            try:
                self.session.findById(_plnal).text = "1"
            except Exception:
                pass
            self.session.findById(_matnr).text = zfer_nuevo
            self.session.findById(_werks).text  = "CO01"
            self.session.findById(_werks).caretPosition = 4
            self._esperar(T_RAPIDO)
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self._esperar(T_MEDIO)
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self._esperar(T_LENTO)
            try:
                self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                self._esperar(T_RAPIDO)
            except Exception:
                pass
            print(f"    CA02 asignación guardada: HR {id_hruta} → {zfer_nuevo}")
            if res: res._log(f"  [HR-ASIGNAR] OK: HR={id_hruta} → {zfer_nuevo}")
            return True
        except Exception as e:
            _warn(str(e))
            return False

    # ── CEWB — Eliminar posición acero ───────────────────────────────────────

    def cewb_eliminar_posicion_acero(self, material: str, pos_acero: str):
        """
        Navega a CEWB, filtra por material (ZFER nuevo) y pos_acero, selecciona fila 0,
        elimina y guarda.
        """
        zpla_nuevo = material  # compatibilidad interna — campo MATNR en CEWB
        print(f"    CEWB: eliminando pos {pos_acero} de material={material}")
        self._navegar("CEWB")
        self._esperar(T_MEDIO)

        # Popup inicial de área de trabajo
        try:
            self.session.findById("wnd[1]/usr/ctxtCWB_WORKAREA-WORK_AREA").text = "SAP_ITEM"
            self.session.findById("wnd[1]/usr/ctxtCWB_WORKAREA-WORK_AREA").caretPosition = 8
            self.session.findById("wnd[1]").sendVKey(0)
            self._esperar(T_MEDIO)
        except Exception:
            pass

        # Filtros: ZPLA, CO01, posición
        _base_cewb = ("wnd[0]/usr/subSELECTION_CRITERIA:SAPLCPSC:1250"
                      "/tabsTAB_STRIP_SEL/tabpITMS"
                      "/ssubSUBPAGE:SAPLCPSC:3345")
        try:
            self.session.findById(_base_cewb + "/ctxtMBMMATNR-LOW").text  = zpla_nuevo
            self.session.findById(_base_cewb + "/ctxtMBMWERKS-LOW").text  = "CO01"
            self.session.findById(_base_cewb + "/txtITMPOSNR-LOW").text   = pos_acero
            self.session.findById(_base_cewb + "/txtITMPOSNR-LOW").setFocus()
            self.session.findById(_base_cewb + "/txtITMPOSNR-LOW").caretPosition = 4
        except Exception as e:
            print(f"    [WARN] CEWB filtros: {e}")

        self.session.findById(self._ID_BTN_EXEC).press()
        self._esperar(T_LENTO)

        # Intentar seleccionar el tab de resultados si no está activo
        _tab_activo = "tabpITM_TGEN"
        for _tab in ("tabpITM_TGEN", "tabpITM_STR", "tabpITM_DOC"):
            try:
                self.session.findById(f"wnd[0]/usr/tabsTAB_STRIP_ITM/{_tab}").select()
                _tab_activo = _tab
                self._esperar(T_RAPIDO)
                break
            except Exception:
                pass

        # Intentar múltiples rutas para la tabla de resultados
        _tbl_candidates = [
            f"wnd[0]/usr/tabsTAB_STRIP_ITM/{_tab_activo}/ssubSUBPAGE:SAPLCSOV:3205/tblSAPLCSOVTC_3205",
            "wnd[0]/usr/tabsTAB_STRIP_ITM/tabpITM_TGEN/ssubSUBPAGE:SAPLCSOV:3205/tblSAPLCSOVTC_3205",
            "wnd[0]/usr/tabsTAB_STRIP_ITM/tabpITM_STR/ssubSUBPAGE:SAPLCSOV:3205/tblSAPLCSOVTC_3205",
        ]
        _tbl_cewb = None
        for _cand in _tbl_candidates:
            try:
                self.session.findById(_cand).getAbsoluteRow(0)
                _tbl_cewb = _cand
                print(f"    CEWB: tabla encontrada en {_cand}")
                break
            except Exception:
                pass
            
        if _tbl_cewb is None:
            # Diagnóstico: mostrar hijos de wnd[0]/usr para ayudar a grabar VBS
            try:
                _usr = self.session.findById("wnd[0]/usr")
                print(f"    [DIAG] CEWB wnd[0]/usr hijos: {[_usr.Children(i).Id for i in range(min(_usr.Children.Count, 10))]}")
            except Exception as _de:
                print(f"    [DIAG] CEWB no pudo inspeccionar usr: {_de}")
            print(f"    [WARN] CEWB eliminar: tabla tblSAPLCSOVTC_3205 no encontrada — grabar VBS en CEWB para confirmar ruta")
        else:
            try:
                self.session.findById(_tbl_cewb).getAbsoluteRow(0).selected = True
                self.session.findById(_tbl_cewb + "/txtITM_CLASS_VIEW-ITM_LOCK[0,0]").setFocus()
                self.session.findById(_tbl_cewb + "/txtITM_CLASS_VIEW-ITM_LOCK[0,0]").caretPosition = 0
                self._esperar(T_RAPIDO)
                self.session.findById("wnd[0]/tbar[1]/btn[14]").press()   # botón borrar
                self._esperar(T_MEDIO)

                # Confirmaciones
                for btn in ("wnd[1]/tbar[0]/btn[0]", "wnd[1]/usr/btnSPOP-OPTION1"):
                    try:
                        self.session.findById(btn).press()
                        self._esperar(T_RAPIDO)
                    except Exception:
                        pass

                # Guardar y salir
                self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                self._esperar(T_LENTO)
                self.session.findById("wnd[0]").sendVKey(3)   # Back
                self._esperar(T_RAPIDO)
                try:
                    self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                except Exception:
                    pass
            except Exception as e:
                print(f"    [WARN] CEWB eliminar: {e}")

    # ── Procesar una combinación ──────────────────────────────────────────────

    def procesar(self, zfer_base: str, color_codigo: str, color_nombre: str,
                 franja: str = "00", pn_base: str = "", zpla: str = "",
                 nivel: str = "", tipo_pieza: str = "",
                 step_callback=None) -> ResultadoItem:
        """
        Procesa una combinación zfer_base + color en SAP.
        franja, pn_base y zpla vienen resueltos desde la BD por app.py.
        color_codigo es el código SAP directo (ej: "19").
        """
        res = ResultadoItem(
            batch_id     = str(uuid.uuid4())[:8],
            zfer_base    = zfer_base,
            color_codigo = color_codigo,
            tipo_pieza   = str(tipo_pieza or ""),
            estado       = "EN_PROCESO",
            fecha_inicio = datetime.datetime.now(),
        )

        def _cb(paso_num: int, desc: str):
            res._log(f"PASO {paso_num}/5: {desc}")
            if step_callback:
                try:
                    step_callback(paso_num, desc)
                except Exception:
                    pass
                
        try:
            res._log(f"=== Inicio: {zfer_base} → color {color_codigo} ({color_nombre}) ===")
            res._log(f"  Franja={franja}  PN_base={pn_base}  ZPLA={zpla}")

            p_color = color_codigo.strip()
            p_franj = franja or "00"

            # PASO 1 — ZPPP0042 validar
            _cb(1, "Validando ZFER base en SAP (ZPPP0042)")
            val = self.zppp0042_validar(zfer_base)
            if not val["ok"]:
                raise RuntimeError(f"ZPPP0042: {val['error']}")
            res._log(f"  VERID={val['verid']} — OK")

            # ── Caso especial BE: nivel 02/03 + tipo_pieza 009/090 ───────────────
            # Si se cumple la condición, verificar si posición 0100 del ZFER base
            # en ZPPR0008 tiene clase que termine en "800". Si es así, en el F4
            # del ZPLA hay que seleccionar la fila que contenga "BE" en descripción.
            forzar_be = False
            nivel_norm     = str(nivel or "").strip().lstrip("0") or "0"
            tipopza_norm   = str(tipo_pieza or "").strip().lstrip("0") or "0"
            es_caso_be     = nivel_norm in ("2", "3") and tipopza_norm in ("9", "90")
            if es_caso_be:
                res._log(f"  Caso especial BE: nivel={nivel} tipo_pieza={tipo_pieza} — verificando pos 0100 en ZPPR0008")
                try:
                    # Reutilizamos _leer_clases_zpla_sap pero sobre el ZFER base
                    clases_zfer = self._leer_clases_zpla_sap(zfer_base)
                    clase_0100  = clases_zfer.get("0100", clases_zfer.get("100", ""))
                    res._log(f"  Pos 0100 → clase: '{clase_0100}'")
                    if clase_0100.upper().endswith("800") or clase_0100.upper().endswith("800_"):
                        forzar_be = True
                        res._log("  → Clase termina en 800: se seleccionará ZPLA con 'BE'")
                    else:
                        res._log("  → Clase NO termina en 800: proceso normal")
                except Exception as e:
                    res._log(f"  [WARN] No pudo verificar pos 0100: {e} — proceso normal")

            # PASO 2 — ZMME0001 ejecutar
            _cb(2, f"Creando nueva variante de color en SAP (ZMME0001) — color {p_color}")
            # zpla puede ser string simple o lista separada por comas
            zplas_validos = [z.strip() for z in str(zpla).split(",") if z.strip()]
            zfer_nuevo, zfor_nuevo, zpla_usado = self.zmme0001_ejecutar(
                zfer_base, p_color, p_franj, zplas_validos, forzar_be=forzar_be
            )
            res.zfer_nuevo = zfer_nuevo
            res.zfor_nuevo = zfor_nuevo
            res.zpla       = zpla_usado

            # PASO 3 — ZPPR0020 esperar fases
            _cb(3, f"Esperando aprobación del proceso SAP (ZPPR0020) — {zfer_nuevo}")
            fase_res = self.zppr0020_esperar_fases(zfer_nuevo)
            if not fase_res["ok"]:
                raise RuntimeError(
                    f"ZPPR0020 falló — {fase_res['fase_error']}: {fase_res['detalle']}"
                )
            if not zpla_usado and fase_res.get("zpla"):
                zpla_usado = fase_res["zpla"]
                res.zpla   = zpla_usado
            res._log(f"  ZPPR0020 OK | ZPLA={zpla_usado}")

            # PASO 4 — ZMME0001 BOM
            _cb(4, "Comparando y copiando estructura de materiales (BOM)")
            # Re-establecer campos (SAP puede haberlos perdido al volver de ZPPR0020)
            try:
                self.session.findById(self._ID_RAD_HOMOLOG).setFocus()
                self.session.findById(self._ID_RAD_HOMOLOG).select()
                self._esperar(T_RAPIDO)
                self.session.findById(self._ID_CTX_CENTER).text = "CO01"
                self.session.findById(self._ID_RAD_COLOR).setFocus()
                self.session.findById(self._ID_RAD_COLOR).select()
                self._esperar(T_RAPIDO)
                self.session.findById(self._ID_CTX_P_COLOR).text = p_color
                self.session.findById(self._ID_CTX_P_FRANJ).text = p_franj
                # ZPLA: re-escribir si está vacío
                zpla_actual = ""
                try:
                    zpla_actual = self.session.findById(self._ID_CTX_P_ZPLA).text.strip()
                except Exception:
                    pass
                if not zpla_actual and zpla_usado:
                    self.session.findById(self._ID_CTX_P_ZPLA).text = f" {zpla_usado}"
            except Exception as e_p4:
                print(f"     [WARN] Re-establecer campos paso 4: {e_p4}")

            # Cambiar material a ZFER_NUEVO
            self.session.findById(self._ID_MATER_LOW).text = zfer_nuevo
            self.session.findById(self._ID_MATER_LOW).caretPosition = len(zfer_nuevo)
            self._esperar(T_RAPIDO)

            # Leer clases del ZPLA antes del loop BOM
            clases = {}
            if zpla_usado:
                clases = self._leer_clases_zpla_sap(zpla_usado)
                res._log(f"  Clases leídas desde SAP: {clases}")
            else:
                res._log("  [WARN] Sin ZPLA para leer clases")

            posiciones = self.bom_con_retry(zpla_usado, clases, on_retry=lambda m: res._advertir(m))
            res.posiciones_bom = posiciones
            res.bom_detalle    = [{"posnr": p["pos"], "clase_destino": clases.get(p["pos"], clases.get(p["pos"].lstrip("0"), ""))} for p in posiciones]
            res._log(f"  Posiciones BOM procesadas ({len(posiciones)}): {posiciones}")

            # PASO 5 — MM02 actualizar PARTNUMBER
            if pn_base and p_color:
                _cb(5, f"Actualizando número de parte en SAP (MM02) — {zfer_nuevo}")
                nuevo_pn = self._construir_nuevo_pn(pn_base, p_color)
                if nuevo_pn and nuevo_pn != pn_base:
                    self.mm02_actualizar_partnumber(zfer_nuevo, nuevo_pn)
                    if zfor_nuevo:
                        self.mm02_actualizar_partnumber(zfor_nuevo, nuevo_pn)
                    res._log(f"  PARTNUMBER → {nuevo_pn}")
                else:
                    res._log("  PARTNUMBER sin cambio necesario")
            else:
                res._log("PASO 5: omitido (sin PN base)")

            res.estado    = "OK"
            res.fecha_fin = datetime.datetime.now()
            res._log(f"=== COMPLETADO OK ({res.duracion_seg}s) ===")
            
        except Exception as e:
            res.estado    = "ERROR"
            res.error     = str(e)
            res.fecha_fin = datetime.datetime.now()
            res._log(f"=== ERROR: {e} ===")

        self._log_bd(res)
        return res


    # ── Procesar fórmula sin acero ───────────────────────────────────────────

    def procesar_formula_sin_acero(
            self, zfer_base: str, formula_nueva: str,
            color_codigo: str, color_nombre: str,
            franja: str = "00", pn_base: str = "", zpla: str = "",
            nivel: str = "", tipo_pieza: str = "",
            step_callback=None) -> ResultadoItem:
        """
        Flujo completo: cambio de fórmula con acero → sin acero.
        Pasos: ZPPR0008 validar pos acero → ZPPP0042 → ZMME0001 fórmula →
               ZPPR0020 → ZMME0001 BOM → MM02 (PN+diferencial+plano) → CEWB borrar pos.
        """
        res = ResultadoItem(
            batch_id     = str(uuid.uuid4())[:8],
            zfer_base    = zfer_base,
            color_codigo = color_codigo,
            formula      = str(formula_nueva or ""),
            tipo_pieza   = str(tipo_pieza or ""),
            estado       = "EN_PROCESO",
            fecha_inicio = datetime.datetime.now(),
        )

        def _cb(paso_num: int, desc: str):
            res._log(f"PASO {paso_num}/7: {desc}")
            if step_callback:
                try:
                    step_callback(paso_num, desc)
                except Exception:
                    pass

        try:
            res._log(f"=== Inicio fórmula sin acero: {zfer_base} → fórmula {formula_nueva} color {color_codigo} ===")
            res._log(f"  Franja={franja}  PN_base={pn_base}  ZPLA={zpla}")

            p_color = color_codigo.strip()
            p_franj = franja or "00"
            zplas_validos = [z.strip() for z in str(zpla).split(",") if z.strip()]
            zpla_base = zplas_validos[0] if zplas_validos else ""

            # PASO 0 — Validar posición acero en ZPPR0008 (usa ZFER base, modo material)
            _cb(0, f"Validando posición acero en ZPPR0008 (ZFER={zfer_base})")
            val_acero = self.zppr0008_validar_posicion_acero(zfer_base)
            if not val_acero["ok"]:
                raise RuntimeError(val_acero["error"])
            pos_acero = val_acero["pos"]   # "0106" ó "0116"
            res._log(f"  Posición acero encontrada: {pos_acero}")

            # PASO 1 — ZPPP0042
            _cb(1, "Validando ZFER base en SAP (ZPPP0042)")
            val = self.zppp0042_validar(zfer_base)
            if not val["ok"]:
                raise RuntimeError(f"ZPPP0042: {val['error']}")
            res._log(f"  VERID={val['verid']} — OK")

            # Caso BE
            forzar_be    = False
            nivel_norm   = str(nivel or "").strip().lstrip("0") or "0"
            tipopza_norm = str(tipo_pieza or "").strip().lstrip("0") or "0"
            if nivel_norm in ("2", "3") and tipopza_norm in ("9", "90"):
                try:
                    clases_zfer = self._leer_clases_zpla_sap(zfer_base)
                    clase_0100  = clases_zfer.get("0100", clases_zfer.get("100", ""))
                    if clase_0100.upper().endswith("800") or clase_0100.upper().endswith("800_"):
                        forzar_be = True
                        res._log("  Caso BE activo")
                except Exception as e:
                    res._log(f"  [WARN] BE check: {e}")

            # PASO 2 — ZMME0001 fórmula
            _cb(2, f"Homologando cambio de fórmula en SAP (ZMME0001) — {formula_nueva}")
            zfer_nuevo, zfor_nuevo, zpla_usado = self.zmme0001_ejecutar_formula(
                zfer_base, p_color, p_franj, formula_nueva, zplas_validos, forzar_be=forzar_be
            )
            res.zfer_nuevo = zfer_nuevo
            res.zfor_nuevo = zfor_nuevo
            res.zpla       = zpla_usado

            # PASO 3 — ZPPR0020
            _cb(3, f"Esperando aprobación del proceso SAP (ZPPR0020) — {zfer_nuevo}")
            fase_res = self.zppr0020_esperar_fases(zfer_nuevo)
            if not fase_res["ok"]:
                raise RuntimeError(f"ZPPR0020 falló — {fase_res['fase_error']}: {fase_res['detalle']}")
            if not zpla_usado and fase_res.get("zpla"):
                zpla_usado = fase_res["zpla"]
                res.zpla   = zpla_usado
            res._log(f"  ZPPR0020 OK | ZPLA={zpla_usado}")

            # PASO 4 — ZMME0001 BOM (igual que cambio de color)
            _cb(4, "Comparando y copiando estructura de materiales (BOM)")
            try:
                self.session.findById(self._ID_RAD_HOMOLOG).setFocus()
                self.session.findById(self._ID_RAD_HOMOLOG).select()
                self._esperar(T_RAPIDO)
                self.session.findById(self._ID_CTX_CENTER).text = "CO01"
                self.session.findById(self._ID_RAD_FORMULA).setFocus()
                self.session.findById(self._ID_RAD_FORMULA).select()
                self._esperar(T_RAPIDO)
                self.session.findById(self._ID_CTX_P_COLOR).text = p_color
                self.session.findById(self._ID_CTX_P_FRANJ).text = p_franj
                self.session.findById(self._ID_TXT_FORMU).text   = formula_nueva
                zpla_actual = ""
                try:
                    zpla_actual = self.session.findById(self._ID_CTX_P_ZPLA).text.strip()
                except Exception:
                    pass
                if not zpla_actual and zpla_usado:
                    self.session.findById(self._ID_CTX_P_ZPLA).text = f" {zpla_usado}"
            except Exception as e_p4:
                print(f"     [WARN] Re-establecer campos paso 4: {e_p4}")

            self.session.findById(self._ID_MATER_LOW).text = zfer_nuevo
            self.session.findById(self._ID_MATER_LOW).caretPosition = len(zfer_nuevo)
            self._esperar(T_RAPIDO)

            clases = {}
            if zpla_usado:
                clases = self._leer_clases_zpla_sap(zpla_usado)
                res._log(f"  Clases leídas desde SAP: {clases}")
            else:
                res._log("  [WARN] Sin ZPLA para leer clases")

            posiciones = self.bom_con_retry(zpla_usado, clases, on_retry=lambda m: res._advertir(m))
            res.posiciones_bom = posiciones
            res.bom_detalle    = [{"posnr": p["pos"], "clase_destino": clases.get(p["pos"], clases.get(p["pos"].lstrip("0"), ""))} for p in posiciones]
            res._log(f"  Posiciones BOM procesadas ({len(posiciones)}): {posiciones}")

            # PASO 5 — MM02 extendido: PN + diferencial 06 + plano
            _cb(5, f"Actualizando MM02 (PN, diferencial, plano) — {zfer_nuevo}")
            nuevo_pn = self._construir_nuevo_pn_formula(pn_base, formula_nueva, p_color)
            res._log(f"  Nuevo PN={nuevo_pn}")

            for mat in ([zfer_nuevo] + ([zfor_nuevo] if zfor_nuevo else [])):
                # 5a — PARTNUMBER
                if nuevo_pn and nuevo_pn != pn_base:
                    self.mm02_actualizar_partnumber(mat, nuevo_pn)
                # 5b — Diferencial 06
                self.mm02_desactivar_diferencial_06(mat)

            # 5c — Plano: solo para ZFER nuevo (no ZFOR)
            self.mm02_cambiar_plano(zfer_nuevo, res, zfer_base=zfer_base)

            # PASO 6 — CEWB: eliminar posición acero del ZFER nuevo
            _cb(6, f"Eliminando posición acero {pos_acero} en CEWB (ZFER={zfer_nuevo})")
            if zfer_nuevo and pos_acero:
                self.cewb_eliminar_posicion_acero(zfer_nuevo, pos_acero)
                res._log(f"  CEWB: pos {pos_acero} eliminada de ZFER={zfer_nuevo}")
            else:
                res._log("  CEWB: omitido (sin ZFER nuevo o sin pos_acero)")

            # PASO 7 — Volver a pantalla inicial (ZMME0001) para siguiente combinación
            _cb(7, "Volviendo a pantalla inicial")
            try:
                # Cerrar cualquier popup pendiente antes de navegar
                for _btn in ("wnd[1]/tbar[0]/btn[0]", "wnd[1]/usr/btnSPOP-OPTION1"):
                    try:
                        self.session.findById(_btn).press()
                        self._esperar(T_RAPIDO)
                    except Exception:
                        pass
                self._navegar("ZMME0001")
                self._esperar(T_MEDIO)
            except Exception as e:
                print(f"    [WARN] PASO 7 navegación: {e}")

            res.estado    = "OK"
            res.fecha_fin = datetime.datetime.now()
            res._log(f"=== COMPLETADO OK ({res.duracion_seg}s) ===")

        except Exception as e:
            res.estado    = "ERROR"
            res.error     = str(e)
            res.fecha_fin = datetime.datetime.now()
            res._log(f"=== ERROR: {e} ===")

        self._log_bd(res)
        return res

    # ── Flujo mismo acero (sin→sin o con→con) ───────────────────────────────

    def procesar_formula_mismo_acero(
            self, zfer_base: str, formula_nueva: str,
            color_codigo: str, color_nombre: str,
            franja: str = "00", pn_base: str = "", zpla: str = "",
            nivel: str = "", tipo_pieza: str = "",
            cambio_hr: bool = False,
            step_callback=None) -> ResultadoItem:
        """
        Flujo cambio de fórmula sin cambio de acero (sin→sin o con→con).
        Pasos: ZPPP0042 → ZMME0001 fórmula → ZPPR0020 → BOM → MM02 (solo PN) → CA02 (si cambio_hr).
        NO hace diferencial 06, NO cambia plano, NO toca CS02/CEWB.
        """
        res = ResultadoItem(
            batch_id     = str(uuid.uuid4())[:8],
            zfer_base    = zfer_base,
            color_codigo = color_codigo,
            formula      = str(formula_nueva or ""),
            tipo_pieza   = str(tipo_pieza or ""),
            estado       = "EN_PROCESO",
            fecha_inicio = datetime.datetime.now(),
        )

        n_pasos = 6 if cambio_hr else 5

        def _cb(paso_num: int, desc: str):
            res._log(f"PASO {paso_num}/{n_pasos}: {desc}")
            if step_callback:
                try:
                    step_callback(paso_num, desc)
                except Exception:
                    pass

        try:
            res._log(f"=== Inicio fórmula mismo acero: {zfer_base} → fórmula {formula_nueva} color {color_codigo} ===")
            res._log(f"  Franja={franja}  PN_base={pn_base}  ZPLA={zpla}  cambio_hr={cambio_hr}")

            p_color = color_codigo.strip()
            p_franj = franja or "00"
            zplas_validos = [z.strip() for z in str(zpla).split(",") if z.strip()]
            res._log(f"  color_nombre={color_nombre or '—'}")

            # PASO 0 — ZPPP0042
            _cb(0, f"Validando ZFER base en SAP (ZPPP0042) — {zfer_base}")
            val = self.zppp0042_validar(zfer_base)
            if not val["ok"]:
                raise RuntimeError(f"ZPPP0042: {val['error']}")
            res._log(f"  VERID={val['verid']} — OK")

            # Caso BE
            forzar_be    = False
            nivel_norm   = str(nivel or "").strip().lstrip("0") or "0"
            tipopza_norm = str(tipo_pieza or "").strip().lstrip("0") or "0"
            if nivel_norm in ("2", "3") and tipopza_norm in ("9", "90"):
                try:
                    clases_zfer = self._leer_clases_zpla_sap(zfer_base)
                    clase_0100  = clases_zfer.get("0100", clases_zfer.get("100", ""))
                    if clase_0100.upper().endswith("800") or clase_0100.upper().endswith("800_"):
                        forzar_be = True
                        res._log("  Caso BE activo")
                except Exception as e:
                    res._log(f"  [WARN] BE check: {e}")

            # PASO 1 — ZMME0001 Cambio de Fórmula
            _cb(1, f"Homologando cambio de fórmula en SAP (ZMME0001) — {formula_nueva}")
            zfer_nuevo, zfor_nuevo, zpla_usado = self.zmme0001_ejecutar_formula(
                zfer_base, p_color, p_franj, formula_nueva, zplas_validos, forzar_be=forzar_be
            )
            res.zfer_nuevo = zfer_nuevo
            res.zfor_nuevo = zfor_nuevo
            res.zpla       = zpla_usado

            # PASO 2 — ZPPR0020
            _cb(2, f"Esperando aprobación del proceso SAP (ZPPR0020) — {zfer_nuevo}")
            fase_res = self.zppr0020_esperar_fases(zfer_nuevo)
            if not fase_res["ok"]:
                raise RuntimeError(f"ZPPR0020 falló — {fase_res['fase_error']}: {fase_res['detalle']}")
            if not zpla_usado and fase_res.get("zpla"):
                zpla_usado = fase_res["zpla"]
                res.zpla   = zpla_usado
            res._log(f"  ZPPR0020 OK | ZPLA={zpla_usado}")

            # PASO 3 — ZMME0001 BOM
            _cb(3, "Comparando y copiando estructura de materiales (BOM)")
            try:
                self.session.findById(self._ID_RAD_HOMOLOG).setFocus()
                self.session.findById(self._ID_RAD_HOMOLOG).select()
                self._esperar(T_RAPIDO)
                self.session.findById(self._ID_CTX_CENTER).text = "CO01"
                self.session.findById(self._ID_RAD_FORMULA).setFocus()
                self.session.findById(self._ID_RAD_FORMULA).select()
                self._esperar(T_RAPIDO)
                self.session.findById(self._ID_CTX_P_COLOR).text = p_color
                self.session.findById(self._ID_CTX_P_FRANJ).text = p_franj
                self.session.findById(self._ID_TXT_FORMU).text   = formula_nueva
                zpla_actual = ""
                try:
                    zpla_actual = self.session.findById(self._ID_CTX_P_ZPLA).text.strip()
                except Exception:
                    pass
                if not zpla_actual and zpla_usado:
                    self.session.findById(self._ID_CTX_P_ZPLA).text = f" {zpla_usado}"
            except Exception as e_p3:
                print(f"     [WARN] Re-establecer campos paso 3: {e_p3}")

            self.session.findById(self._ID_MATER_LOW).text = zfer_nuevo
            self.session.findById(self._ID_MATER_LOW).caretPosition = len(zfer_nuevo)
            self._esperar(T_RAPIDO)

            clases = {}
            if zpla_usado:
                clases = self._leer_clases_zpla_sap(zpla_usado)
                res._log(f"  Clases leídas desde SAP: {clases}")
            else:
                res._log("  [WARN] Sin ZPLA para leer clases")

            posiciones = self.bom_con_retry(zpla_usado, clases, on_retry=lambda m: res._advertir(m))
            res.posiciones_bom = posiciones
            res.bom_detalle    = [{"posnr": p["pos"], "clase_destino": clases.get(p["pos"], clases.get(p["pos"].lstrip("0"), ""))} for p in posiciones]
            res._log(f"  Posiciones BOM procesadas ({len(posiciones)}): {posiciones}")

            # PASO 4 — MM02: solo actualizar PN (sin diferencial, sin plano)
            _cb(4, f"Actualizando MM02 (solo PN) — {zfer_nuevo}")
            nuevo_pn = self._construir_nuevo_pn_formula(pn_base, formula_nueva, p_color)
            res._log(f"  Nuevo PN={nuevo_pn}")
            for mat in ([zfer_nuevo] + ([zfor_nuevo] if zfor_nuevo else [])):
                if nuevo_pn and nuevo_pn != pn_base:
                    self.mm02_actualizar_partnumber(mat, nuevo_pn)

            # PASO 5 — CA02 (solo si cambio_hr=True)
            if cambio_hr:
                _cb(5, f"Buscando y asignando hoja de ruta (CA02) — {zfer_nuevo}")
                try:
                    from app import _hr_buscar_candidata
                    hr_id, hr_desc, hr_err = _hr_buscar_candidata(zfer_base, zfer_nuevo)
                    if hr_err:
                        res._log(f"  [WARN] HR candidata no encontrada: {hr_err}")
                    else:
                        res._log(f"  HR candidata: {hr_id} ({hr_desc})")
                        self.ca02_desasignar_hr(zfer_nuevo, res)
                        ok_asi = self.ca02_asignar_hr(zfer_nuevo, hr_id, res)
                        if ok_asi:
                            res._log(f"  CA02 OK: HR={hr_id} → {zfer_nuevo}")
                        else:
                            res._log(f"  [WARN] CA02 asignación falló — revisar manualmente")
                except Exception as e_hr:
                    res._log(f"  [WARN] CA02: {e_hr}")

            # Volver a pantalla inicial
            try:
                for _btn in ("wnd[1]/tbar[0]/btn[0]", "wnd[1]/usr/btnSPOP-OPTION1"):
                    try:
                        self.session.findById(_btn).press()
                        self._esperar(T_RAPIDO)
                    except Exception:
                        pass
                self._navegar("ZMME0001")
                self._esperar(T_MEDIO)
            except Exception as e:
                print(f"    [WARN] Navegación final: {e}")

            res.estado    = "OK"
            res.fecha_fin = datetime.datetime.now()
            res._log(f"=== COMPLETADO OK ({res.duracion_seg}s) ===")

        except Exception as e:
            res.estado    = "ERROR"
            res.error     = str(e)
            res.fecha_fin = datetime.datetime.now()
            res._log(f"=== ERROR: {e} ===")

        self._log_bd(res)
        return res

    # ── CS02 — Agregar posición acero en BOM del ZFOR ────────────────────────

    def cs02_agregar_posicion_acero(self, zfor: str, pos_acero: str, zhal: str, res=None) -> bool:
        """
        CS02 con MATNR=zfor, Utilización=1, Centro=CO01.
        Escanea tabla TCMA buscando fila con POSTP=='' (vacía).
        Escribe POSNR=pos_acero, POSTP=l, IDNRK=zhal, MENGE=1 → guarda.

        IDs confirmados por VBS (QUAS):
          Pantalla inicial: ctxtRC29N-MATNR, ctxtRC29N-WERKS, ctxtRC29N-STLAN
          Tabla: wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT
            col 0 = txtRC29P-POSNR   col 1 = ctxtRC29P-POSTP   col 2 = ctxtRC29P-IDNRK
          Sub-screen MENGE: wnd[0]/usr/subPOS_PHPT:SAPLCSDI:0830/txtRC29P-MENGE
          Guardar: wnd[0]/tbar[0]/btn[11]
        Detección fila vacía: POSTP=='' (SAP pre-llena POSNR con 9310/9320/... — NO usar POSNR)
        """
        _TBL = ("wnd[0]/usr/tabsTS_ITOV/tabpTCMA"
                "/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT")

        def _warn(msg):
            print(f"    [WARN] cs02_agregar: {msg}")
            if res: res._log(f"  [CS02] ⚠ {msg}")

        def _leer_postp(vis):
            try:
                return str(self.session.findById(f"{_TBL}/ctxtRC29P-POSTP[1,{vis}]").text or "").strip()
            except Exception:
                return None  # None = fila no existe

        def _leer_posnr(vis):
            try:
                return str(self.session.findById(f"{_TBL}/txtRC29P-POSNR[0,{vis}]").text or "").strip()
            except Exception:
                return None

        def _leer_idnrk(vis):
            try:
                return str(self.session.findById(f"{_TBL}/ctxtRC29P-IDNRK[2,{vis}]").text or "").strip()
            except Exception:
                return ""

        # Normalizar: "0116" → "116"
        pos_num = str(pos_acero).lstrip("0") or pos_acero

        print(f"    CS02: ZFOR={zfor} pos={pos_num} ZHAL={zhal}")
        if res: res._log(f"  [CS02] Iniciando: ZFOR={zfor} pos={pos_num} ZHAL={zhal}")

        try:
            # ── Navegar a CS02 ──────────────────────────────────────────────
            self._cerrar_dialogs_abiertos()
            self.session.findById(self._ID_TCODE_BOX).text = "/ncs02"
            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)

            # ── Pantalla inicial: ZFOR + Centro + Utilización ───────────────
            try:
                self.session.findById("wnd[0]/usr/ctxtRC29N-MATNR").text = zfor
            except Exception as e:
                raise RuntimeError(f"No se pudo escribir ZFOR en pantalla inicial CS02: {e}")

            try:
                self.session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = "CO01"
            except Exception as e:
                _warn(f"Centro CO01 no disponible: {e}")

            try:
                self.session.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = "1"
            except Exception as e:
                _warn(f"Utilización=1 no disponible: {e}")

            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)

            # ── Seleccionar tab TCMA (componentes) ─────────────────────────
            try:
                self.session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA").select()
                self._esperar(T_RAPIDO)
            except Exception:
                pass  # ya puede estar activo

            # ── Obtener parámetros de la tabla ─────────────────────────────
            tbl      = None
            max_sb   = 0
            vis_rows = 19
            try:
                tbl      = self.session.findById(_TBL)
                max_sb   = tbl.verticalScrollbar.maximum
                vis_rows = tbl.VisibleRowCount
                print(f"    CS02 tabla: VisibleRows={vis_rows} scrollMax={max_sb}")
                if res: res._log(f"  [CS02] Tabla: vis={vis_rows} scroll_max={max_sb}")
            except Exception as e:
                _warn(f"No se pudo leer tabla CS02 ({e}) — usando defaults vis=19 scroll=0")

            # ── Verificar si la posición ya existe (evitar duplicado) ───────
            for sp_check in range(max_sb + 1):
                if tbl:
                    self._ca02_scroll(tbl, sp_check)
                for vis in range(vis_rows):
                    posnr = _leer_posnr(vis)
                    if posnr is None:
                        break
                    postp = _leer_postp(vis)
                    idnrk = _leer_idnrk(vis)
                    # Posición ya existe si POSNR coincide y POSTP no está vacío
                    if posnr == pos_num and postp not in ("", None):
                        _warn(f"Posición {pos_num} ya existe en BOM (IDNRK={idnrk}) — omitiendo CS02")
                        if res: res._log(f"  [CS02] Posición {pos_num} ya existía con IDNRK={idnrk} — no se duplicó")
                        return True  # no es error, ya está
                if sp_check == max_sb:
                    break

            # ── Buscar primera fila vacía (POSTP=='') ───────────────────────
            fila_vis    = None
            fila_scroll = 0
            for sp in range(max_sb + 1):
                if tbl:
                    self._ca02_scroll(tbl, sp)
                for vis in range(vis_rows):
                    posnr = _leer_posnr(vis)
                    if posnr is None:
                        break  # fin de tabla
                    postp = _leer_postp(vis)
                    if postp == "":  # fila disponible
                        fila_vis    = vis
                        fila_scroll = sp
                        print(f"    CS02 fila vacía: vis={vis} scroll={sp} POSNR={posnr!r}")
                        if res: res._log(f"  [CS02] Fila vacía en vis={vis} scroll={sp}")
                        break
                if fila_vis is not None:
                    break

            if fila_vis is None:
                msg = f"No se encontró fila vacía en BOM del ZFOR {zfor} — revisar CS02 manualmente"
                _warn(msg)
                if res: res._log(f"  [CS02] ❌ {msg}")
                return False

            # ── Escribir POSNR, POSTP, IDNRK ──────────────────────────────
            try:
                self.session.findById(f"{_TBL}/txtRC29P-POSNR[0,{fila_vis}]").text = pos_num
            except Exception as e:
                raise RuntimeError(f"No se pudo escribir POSNR={pos_num}: {e}")

            try:
                self.session.findById(f"{_TBL}/ctxtRC29P-POSTP[1,{fila_vis}]").text = "l"
            except Exception as e:
                raise RuntimeError(f"No se pudo escribir POSTP=l: {e}")

            try:
                campo_idnrk = f"{_TBL}/ctxtRC29P-IDNRK[2,{fila_vis}]"
                self.session.findById(campo_idnrk).text = zhal
                self.session.findById(campo_idnrk).setFocus()
                self.session.findById(campo_idnrk).caretPosition = len(zhal)
            except Exception as e:
                raise RuntimeError(f"No se pudo escribir IDNRK={zhal}: {e}")

            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)

            # ── Sub-screen MENGE=1 (aparece tras Enter) ────────────────────
            try:
                menge_id = "wnd[0]/usr/subPOS_PHPT:SAPLCSDI:0830/txtRC29P-MENGE"
                self.session.findById(menge_id).text = "1"
                self.session.findById(menge_id).caretPosition = 1
                self.session.findById("wnd[0]").sendVKey(0)
                self._esperar(T_RAPIDO)
            except Exception as e_menge:
                _warn(f"Sub-screen MENGE no apareció ({e_menge}) — puede ser normal si SAP lo tomó automático")

            # ── Cerrar popups de confirmación si aparecen ──────────────────
            for btn_popup in ("wnd[1]/usr/btnSPOP-OPTION1", "wnd[1]/tbar[0]/btn[0]"):
                try:
                    self.session.findById(btn_popup).press()
                    self._esperar(T_RAPIDO)
                except Exception:
                    pass

            # ── Guardar ────────────────────────────────────────────────────
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self._esperar(T_LENTO)

            # ── Verificar que no quedó mensaje de error en statusbar ───────
            try:
                status = str(self.session.findById("wnd[0]/sbar").text or "").strip()
                if status:
                    print(f"    CS02 statusbar: {status!r}")
                    if res: res._log(f"  [CS02] Statusbar tras guardar: {status}")
                    if any(w in status.upper() for w in ("ERROR", "INCORRECTO", "NO EXISTE", "NO SE PUEDE")):
                        _warn(f"Posible error al guardar CS02: {status}")
                        return False
            except Exception:
                pass

            print(f"    CS02 OK: ZFOR={zfor} pos={pos_num} ZHAL={zhal}")
            if res: res._log(f"  [CS02] ✓ Guardado: pos={pos_num} ZHAL={zhal}")
            return True

        except Exception as e:
            _warn(str(e))
            if res: res._log(f"  [CS02] ❌ Excepción: {e}")
            return False

    # ─────────────────────────────────────────────────────────────────────────
    # FLUJO: sin acero → con acero
    # ─────────────────────────────────────────────────────────────────────────

    def procesar_formula_con_acero(
            self, zfer_base: str, formula_nueva: str,
            color_codigo: str, color_nombre: str,
            franja: str = "00", pn_base: str = "", zpla: str = "",
            nivel: str = "", tipo_pieza: str = "",
            zhal: str = "",
            step_callback=None) -> ResultadoItem:
        """
        Flujo completo: cambio de fórmula sin acero → con acero.
        Pasos:
          0 — Validar versión en ZPPP0042
          1 — ZMME0001 Cambio de Fórmula (igual que sin_acero)
          2 — Validar posición acero en ZPPR0008 con el ZPLA sugerido (sesión aux)
          3 — ZPPR0020 esperar fases
          4 — ZMME0001 BOM (igual que sin_acero)
          5 — MM02: PN + activar diferencial 06 + plano CON SP
          6 — CEWB agregar posición acero (PENDIENTE — se implementará después)
        """
        res = ResultadoItem(
            batch_id     = str(uuid.uuid4())[:8],
            zfer_base    = zfer_base,
            color_codigo = color_codigo,
            formula      = str(formula_nueva or ""),
            tipo_pieza   = str(tipo_pieza or ""),
            estado       = "EN_PROCESO",
            fecha_inicio = datetime.datetime.now(),
        )

        def _cb(paso_num: int, desc: str):
            res._log(f"PASO {paso_num}/6: {desc}")
            if step_callback:
                try:
                    step_callback(paso_num, desc)
                except Exception:
                    pass

        try:
            res._log(f"=== Inicio fórmula con acero: {zfer_base} → fórmula {formula_nueva} color {color_codigo} ===")
            res._log(f"  Franja={franja}  PN_base={pn_base}  ZPLA={zpla}")

            p_color = color_codigo.strip()
            p_franj = franja or "00"
            zplas_validos = [z.strip() for z in str(zpla).split(",") if z.strip()]

            # PASO 0 — Validar versión en ZPPP0042 (sin validar pos acero)
            _cb(0, f"Validando ZFER base en SAP (ZPPP0042) — {zfer_base}")
            val = self.zppp0042_validar(zfer_base)
            if not val["ok"]:
                raise RuntimeError(f"ZPPP0042: {val['error']}")
            res._log(f"  VERID={val['verid']} — OK")

            # Caso BE
            forzar_be    = False
            nivel_norm   = str(nivel or "").strip().lstrip("0") or "0"
            tipopza_norm = str(tipo_pieza or "").strip().lstrip("0") or "0"
            if nivel_norm in ("2", "3") and tipopza_norm in ("9", "90"):
                try:
                    clases_zfer = self._leer_clases_zpla_sap(zfer_base)
                    clase_0100  = clases_zfer.get("0100", clases_zfer.get("100", ""))
                    if clase_0100.upper().endswith("800") or clase_0100.upper().endswith("800_"):
                        forzar_be = True
                        res._log("  Caso BE activo")
                except Exception as e:
                    res._log(f"  [WARN] BE check: {e}")

            # PASO 1 — ZMME0001 Cambio de Fórmula
            _cb(1, f"Homologando cambio de fórmula en SAP (ZMME0001) — {formula_nueva}")
            zfer_nuevo, zfor_nuevo, zpla_usado = self.zmme0001_ejecutar_formula(
                zfer_base, p_color, p_franj, formula_nueva, zplas_validos, forzar_be=forzar_be
            )
            res.zfer_nuevo = zfer_nuevo
            res.zfor_nuevo = zfor_nuevo
            res.zpla       = zpla_usado

            # PASO 2 — Validar posición acero en ZPPR0008 con el ZPLA sugerido
            _cb(2, f"Validando posición acero en ZPPR0008 (ZPLA={zpla_usado})")
            if not zpla_usado:
                raise RuntimeError("No se obtuvo ZPLA desde ZMME0001 — no se puede validar acero en ZPPR0008")
            val_acero = self.zppr0008_validar_posicion_acero_zpla(zpla_usado)
            if not val_acero["ok"]:
                raise RuntimeError(val_acero["error"])
            pos_acero = val_acero["pos"]
            res._log(f"  Posición acero confirmada en ZPLA: {pos_acero}")

            # PASO 3 — ZPPR0020
            _cb(3, f"Esperando aprobación del proceso SAP (ZPPR0020) — {zfer_nuevo}")
            fase_res = self.zppr0020_esperar_fases(zfer_nuevo)
            if not fase_res["ok"]:
                raise RuntimeError(f"ZPPR0020 falló — {fase_res['fase_error']}: {fase_res['detalle']}")
            if not zpla_usado and fase_res.get("zpla"):
                zpla_usado = fase_res["zpla"]
                res.zpla   = zpla_usado
            res._log(f"  ZPPR0020 OK | ZPLA={zpla_usado}")

            # PASO 4 — ZMME0001 BOM (igual que sin_acero)
            _cb(4, "Comparando y copiando estructura de materiales (BOM)")
            try:
                self.session.findById(self._ID_RAD_HOMOLOG).setFocus()
                self.session.findById(self._ID_RAD_HOMOLOG).select()
                self._esperar(T_RAPIDO)
                self.session.findById(self._ID_CTX_CENTER).text = "CO01"
                self.session.findById(self._ID_RAD_FORMULA).setFocus()
                self.session.findById(self._ID_RAD_FORMULA).select()
                self._esperar(T_RAPIDO)
                self.session.findById(self._ID_CTX_P_COLOR).text = p_color
                self.session.findById(self._ID_CTX_P_FRANJ).text = p_franj
                self.session.findById(self._ID_TXT_FORMU).text   = formula_nueva
                zpla_actual = ""
                try:
                    zpla_actual = self.session.findById(self._ID_CTX_P_ZPLA).text.strip()
                except Exception:
                    pass
                if not zpla_actual and zpla_usado:
                    self.session.findById(self._ID_CTX_P_ZPLA).text = f" {zpla_usado}"
            except Exception as e_p4:
                print(f"     [WARN] Re-establecer campos paso 4: {e_p4}")

            self.session.findById(self._ID_MATER_LOW).text = zfer_nuevo
            self.session.findById(self._ID_MATER_LOW).caretPosition = len(zfer_nuevo)
            self._esperar(T_RAPIDO)

            clases = {}
            if zpla_usado:
                clases = self._leer_clases_zpla_sap(zpla_usado)
                res._log(f"  Clases leídas desde SAP: {clases}")
            else:
                res._log("  [WARN] Sin ZPLA para leer clases")

            posiciones = self.bom_con_retry(zpla_usado, clases, on_retry=lambda m: res._advertir(m))
            res.posiciones_bom = posiciones
            res.bom_detalle    = [{"posnr": p["pos"], "clase_destino": clases.get(p["pos"], clases.get(p["pos"].lstrip("0"), ""))} for p in posiciones]
            res._log(f"  Posiciones BOM procesadas ({len(posiciones)}): {posiciones}")

            # PASO 5 — MM02: PN + activar diferencial 06 + plano CON SP
            _cb(5, f"Actualizando MM02 (PN, diferencial 06, plano SP) — {zfer_nuevo}")
            nuevo_pn = self._construir_nuevo_pn_formula(pn_base, formula_nueva, p_color)
            res._log(f"  Nuevo PN={nuevo_pn}")

            for mat in ([zfer_nuevo] + ([zfor_nuevo] if zfor_nuevo else [])):
                # 5a — PARTNUMBER
                if nuevo_pn and nuevo_pn != pn_base:
                    self.mm02_actualizar_partnumber(mat, nuevo_pn)
                # 5b — Activar diferencial 06 (lo marca, no desmarca)
                self.mm02_activar_diferencial_06(mat)

            # 5c — Plano con SP: solo para ZFER nuevo
            self.mm02_cambiar_plano_con_sp(zfer_nuevo, res, zfer_base=zfer_base)

            # PASO 6 — CS02: agregar posición acero en BOM del ZFOR
            _cb(6, f"Agregando posición acero {pos_acero} en CS02 (ZFOR={zfor_nuevo})")
            if zfor_nuevo and zhal:
                ok_cs02 = self.cs02_agregar_posicion_acero(zfor_nuevo, pos_acero, zhal, res)
                if ok_cs02:
                    res._log(f"  CS02: posición {pos_acero} agregada con ZHAL={zhal}")
                else:
                    res._log(f"  [WARN] CS02: falló agregar posición {pos_acero} — revisar BOM manualmente")
            elif not zfor_nuevo:
                res._log(f"  [WARN] CS02 omitido: sin ZFOR nuevo")
            else:
                res._log(f"  [WARN] CS02 omitido: sin código ZHAL")

            # Volver a pantalla inicial
            try:
                for _btn in ("wnd[1]/tbar[0]/btn[0]", "wnd[1]/usr/btnSPOP-OPTION1"):
                    try:
                        self.session.findById(_btn).press()
                        self._esperar(T_RAPIDO)
                    except Exception:
                        pass
                self._navegar("ZMME0001")
                self._esperar(T_MEDIO)
            except Exception as e:
                print(f"    [WARN] Navegación final: {e}")

            res.estado    = "OK"
            res.fecha_fin = datetime.datetime.now()
            res._log(f"=== COMPLETADO OK ({res.duracion_seg}s) ===")

        except Exception as e:
            res.estado    = "ERROR"
            res.error     = str(e)
            res.fecha_fin = datetime.datetime.now()
            res._log(f"=== ERROR: {e} ===")

        self._log_bd(res)
        return res
    

# ── Función de entrada (usada desde app.py vía threading) ────────────────────

def procesar_combinacion(zfer_base: str, color_codigo: str, color_nombre: str,
                         franja: str = "00", pn_base: str = "",
                         zpla: str = "", nivel: str = "", tipo_pieza: str = "",
                         step_callback=None) -> ResultadoItem:
    auto = AutomatizadorSAP()
    if not auto.conectar():
        r = ResultadoItem(
            batch_id=str(uuid.uuid4())[:8], zfer_base=zfer_base,
            color_codigo=color_codigo, estado="ERROR",
            error="SAP GUI no disponible — verifica que SAP esté abierto y scripting habilitado.",
            fecha_inicio=datetime.datetime.now(), fecha_fin=datetime.datetime.now(),
        )
        return r
    return auto.procesar(zfer_base, color_codigo, color_nombre, franja, pn_base, zpla,
                         nivel=nivel, tipo_pieza=tipo_pieza,
                         step_callback=step_callback)


def leer_bom_material(material: str) -> dict:
    """
    Standalone: conecta a SAP, va a ZPPR0008 con el material,
    retorna {"ok": True, "posiciones": [int,...], "filas": [{pos,nombre},...], "error": ""}
    """
    auto = AutomatizadorSAP()
    if not auto.conectar():
        return {"ok": False, "posiciones": [], "filas": [],
                "error": "SAP GUI no disponible — verifica que SAP esté abierto y scripting habilitado."}
    return auto.zppr0008_leer_bom_completo(material)


def procesar_combinacion_formula_con_acero(
        zfer_base: str, formula_nueva: str,
        color_codigo: str, color_nombre: str,
        franja: str = "00", pn_base: str = "", zpla: str = "",
        nivel: str = "", tipo_pieza: str = "",
        zhal: str = "",
        step_callback=None) -> "ResultadoItem":
    """Entrada pública para flujo sin acero → con acero."""
    auto = AutomatizadorSAP()
    if not auto.conectar():
        r = ResultadoItem(
            batch_id=str(uuid.uuid4())[:8], zfer_base=zfer_base,
            color_codigo=color_codigo, estado="ERROR",
            error="SAP GUI no disponible.",
            fecha_inicio=datetime.datetime.now(), fecha_fin=datetime.datetime.now(),
        )
        return r
    return auto.procesar_formula_con_acero(
        zfer_base, formula_nueva, color_codigo, color_nombre,
        franja, pn_base, zpla, nivel=nivel, tipo_pieza=tipo_pieza,
        zhal=zhal, step_callback=step_callback,
    )


def procesar_combinacion_formula_mismo_acero(
        zfer_base: str, formula_nueva: str,
        color_codigo: str, color_nombre: str,
        franja: str = "00", pn_base: str = "", zpla: str = "",
        nivel: str = "", tipo_pieza: str = "",
        cambio_hr: bool = False,
        step_callback=None) -> "ResultadoItem":
    """Entrada pública para flujo mismo acero (sin→sin o con→con)."""
    auto = AutomatizadorSAP()
    if not auto.conectar():
        r = ResultadoItem(
            batch_id=str(uuid.uuid4())[:8], zfer_base=zfer_base,
            color_codigo=color_codigo, estado="ERROR",
            error="SAP GUI no disponible.",
            fecha_inicio=datetime.datetime.now(), fecha_fin=datetime.datetime.now(),
        )
        return r
    return auto.procesar_formula_mismo_acero(
        zfer_base, formula_nueva, color_codigo, color_nombre,
        franja, pn_base, zpla, nivel=nivel, tipo_pieza=tipo_pieza,
        cambio_hr=cambio_hr, step_callback=step_callback,
    )


def procesar_combinacion_formula_sin_acero(
        zfer_base: str, formula_nueva: str,
        color_codigo: str, color_nombre: str,
        franja: str = "00", pn_base: str = "", zpla: str = "",
        nivel: str = "", tipo_pieza: str = "",
        step_callback=None) -> ResultadoItem:
    auto = AutomatizadorSAP()
    if not auto.conectar():
        r = ResultadoItem(
            batch_id=str(uuid.uuid4())[:8], zfer_base=zfer_base,
            color_codigo=color_codigo, estado="ERROR",
            error="SAP GUI no disponible.",
            fecha_inicio=datetime.datetime.now(), fecha_fin=datetime.datetime.now(),
        )
        return r
    return auto.procesar_formula_sin_acero(
        zfer_base, formula_nueva, color_codigo, color_nombre,
        franja, pn_base, zpla, nivel=nivel, tipo_pieza=tipo_pieza,
        step_callback=step_callback,
    )


def cambiar_hoja_ruta(zfer_nuevo: str, id_hruta: str) -> dict:
    """
    Standalone: desasigna la HR actual del zfer_nuevo en CA02 y asigna id_hruta.
    Retorna {"ok": True/False, "error": str}
    """
    auto = AutomatizadorSAP()
    if not auto.conectar():
        return {"ok": False, "error": "SAP GUI no disponible"}
    auto.ca02_desasignar_hr(zfer_nuevo)
    ok_asi = auto.ca02_asignar_hr(zfer_nuevo, id_hruta)
    if ok_asi:
        return {"ok": True, "error": ""}
    return {"ok": False, "error": "ca02_asignar falló — revisa logs SAP"}