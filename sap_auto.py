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

# ── Tiempos de espera ─────────────────────────────────────────────────────────
T_RAPIDO = 1.5
T_MEDIO  = 3.5
T_LENTO  = 7.0

_SAP_USER = os.environ.get("SAP_USER", "PROGRAING") #PROGRAING

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
    estado:         str   = "PENDIENTE"   # EN_PROCESO | OK | ERROR
    error:          str   = ""
    fecha_inicio:   Optional[datetime.datetime] = None
    fecha_fin:      Optional[datetime.datetime] = None
    log:            list  = field(default_factory=list)
    # Campos extra para log BD
    formula:        str   = ""
    tipo_pieza:     str   = ""
    acero:          str   = ""

    @property
    def duracion_seg(self) -> float:
        if self.fecha_inicio and self.fecha_fin:
            return round((self.fecha_fin - self.fecha_inicio).total_seconds(), 1)
        return 0.0

    def _log(self, msg: str):
        print(f"  [SAP] {msg}")
        self.log.append(msg)


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

    def _esperar(self, seg: float = T_RAPIDO):
        time.sleep(seg)

    def _navegar(self, tcode: str):
        self.session.findById(self._ID_TCODE_BOX).text = f"/N{tcode}"
        self.session.findById("wnd[0]").sendVKey(0)
        self._esperar(T_MEDIO)

    def _estado_sap(self) -> str:
        try:
            return self.session.findById(self._ID_STATUSBAR).text.strip()
        except Exception:
            return ""

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
                                intervalo_seg: int = 10,
                                max_espera_seg: int = 600) -> dict:
        """
        Abre sesión auxiliar SAP para ZPPR0020 (deja ZMME0001 intacta en sesión
        principal). Polling hasta > 7 fases con 'S', o error 'E', o timeout.
        Cierra sesión auxiliar al terminar y re-adquiere sesión principal.
        """
        print("     Abriendo sesión auxiliar para ZPPR0020...")
        self.session.createSession()
        self._esperar(T_LENTO)

        idx_nueva = self.conn_sap.Children.Count - 1
        ses2 = self.conn_sap.Children(idx_nueva)
        self._esperar(T_LENTO)   # dar tiempo a que la sesión arranque
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
            self._esperar(T_LENTO * 2)   # máquina lenta: doble espera

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
                    msg = self._estado_sap()
                    print(f"    Ejecutar BOM: {msg}")
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
        self._esperar(T_LENTO)

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
        try:
            cn  = pyodbc.connect(_DB_LOCAL_STR, autocommit=True)
            cur = cn.cursor()
            # Migrar batch_id a varchar si aún es uniqueidentifier (dos pasos separados)
            try:
                cur.execute(
                    "ALTER TABLE dbo.M5_LogEjecucion "
                    "ALTER COLUMN batch_id varchar(50) NULL"
                )
            except Exception:
                pass  # ya es varchar o tabla no existe — ignorar

            _vals = (
                str(res.batch_id)[:50],
                str(res.zfer_base or "")[:50],
                str(getattr(res, "tipo_pieza", "") or "")[:50],
                str(getattr(res, "formula",    "") or "")[:50],
                str(res.color_codigo or "")[:20],
                str(getattr(res, "acero",      "") or "")[:50],
                str(res.estado or "")[:20],
                str(res.error)[:2000] if res.error else None,
                res.fecha_inicio,
                res.fecha_fin,
            )
            cur.execute(
                "INSERT INTO dbo.M5_LogEjecucion "
                "(batch_id, pedido_origen, tipo_pieza, formula, color_codigo, acero_variante, "
                " estado, detalle_error, fecha_inicio, fecha_fin) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                _vals
            )
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

        # Fila visual 7 = Z_BEHAVIOR_DIFFERENTIALS (con scroll en pos 6)
        campo_name = tbl + "/ctxtRCTMS-MNAME[0,7]"
        self.session.findById(campo_name).setFocus()
        self.session.findById(campo_name).caretPosition = 16
        self.session.findById("wnd[0]").sendVKey(2)   # abre popup de valores
        self._esperar(T_MEDIO)

        # En popup wnd[1]: desmarcar checkbox fila 5 (el "06")
        try:
            chk = "wnd[1]/usr/tblSAPLCTMSVALUE_S/chkRCTMS-SEL01[0,5]"
            self.session.findById(chk).selected = False
            self.session.findById(chk).setFocus()
            self.session.findById("wnd[1]").sendVKey(2)   # confirmar selección
            self._esperar(T_RAPIDO)
        except Exception as e:
            print(f"    [WARN] Diferencial popup check: {e}")

        # Cerrar popup
        try:
            self.session.findById("wnd[1]").close()
        except Exception:
            try:
                self.session.findById("wnd[1]").sendVKey(12)
            except Exception:
                pass
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

    def _buscar_plano_bd(self, doknr_actual: str) -> tuple:
        """
        Dado el DOKNR leído de MM02 (puede tener SP y/o letra al final),
        busca en ODATA_ZFER_RUTAS_JPG el PLANO más reciente sin SP,
        ordenado por ULTIMA_MOD DESC para siempre tomar la versión vigente.
        Returns: (plano_nuevo: str | None, mensaje: str)
        """
        # Quitar SP y cualquier sufijo de letra(s) para obtener la base de búsqueda
        # Ej: "M1344 000 001 A SP" → "M1344 000 001"
        #     "M1344 000 001 SP"   → "M1344 000 001"
        base = re.sub(r'(\s+[A-Z]+)?\s+SP\s*$', '', doknr_actual, flags=re.IGNORECASE).strip()
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

        doc_elegido  = str(row[0] or "").strip()
        plano_elegido = str(row[1] or "").strip()

        if not plano_elegido:
            return None, (
                f"⚠ PLANO NO ACTUALIZADO — Se encontró DOCUMENTO='{doc_elegido}' "
                f"pero la columna PLANO está vacía en BD."
            )

        return plano_elegido, f"Plano actualizado: '{doc_elegido}' → '{plano_elegido}'"

    def mm02_cambiar_plano(self, zfer: str, res: "ResultadoItem" = None) -> bool:
        """
        Navega a MM02 → btn[30] → tabpZU04 → radGF_ALLE → lee DOKNR actual →
        busca en ODATA_ZFER_RUTAS_JPG el plano más reciente sin SP →
        reemplaza DOKNR con ese plano → guarda.
        Retorna True si guardó, False si omitió (sin romper el flujo).
        """
        def _warn(msg):
            print(f"    [WARN] mm02_cambiar_plano: {msg}")
            if res:
                res._log(f"  [PLANO] ADVERTENCIA: {msg}")

        print(f"    MM02 plano: procesando {zfer}")
        try:
            _subZU04 = ("wnd[0]/usr/tabsTABSPR1/tabpZU04"
                        "/ssubTABFRA1:SAPLMGMM:2110"
                        "/subSUB2:SAPLMGD1:3400"
                        "/subDOCU:SAPLCV140:0204")
            _grid_docu = _subZU04 + "/subDOC_ALV:SAPLCV140:0206/cntlALV_CUST_DOC/shellcont/shell"

            # /nmm02 → material → Enter × 2
            self.session.findById(self._ID_TCODE_BOX).text = "/nmm02"
            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)
            self.session.findById(self._ID_MM02_MATNR).text = zfer
            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)
            self.session.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)

            # Vista ingeniería → tab ZU04 → radio GF_ALLE
            self.session.findById("wnd[0]/tbar[1]/btn[30]").press()
            self._esperar(T_MEDIO)
            self.session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU04").select()
            self._esperar(T_MEDIO)
            self.session.findById(_subZU04 + "/subBUTTON:SAPLCV140:0203/radGF_ALLE").setFocus()
            self.session.findById(_subZU04 + "/subBUTTON:SAPLCV140:0203/radGF_ALLE").select()
            self._esperar(T_RAPIDO)

            # Leer DOKNR actual
            try:
                doknr_actual = self.session.findById(_grid_docu).getCellValue(0, "DOKNR")
                print(f"    MM02 plano: DOKNR actual='{doknr_actual}'")
            except Exception as e:
                _warn(f"No pudo leer DOKNR: {e}")
                return False

            if not doknr_actual or not str(doknr_actual).strip():
                _warn("DOKNR vacío en MM02, omitiendo cambio de plano")
                return False

            # Buscar plano en BD
            nuevo_plano, msg_bd = self._buscar_plano_bd(str(doknr_actual).strip())
            print(f"    MM02 plano BD: {msg_bd}")
            if res:
                res._log(f"  [PLANO] {msg_bd}")

            if not nuevo_plano:
                # No encontró plano válido → continuar sin cambiar, ya logueado
                return False

            if nuevo_plano == str(doknr_actual).strip():
                print(f"    MM02 plano: sin cambio necesario ('{nuevo_plano}')")
                if res:
                    res._log(f"  [PLANO] Sin cambio necesario ('{nuevo_plano}')")
                return True

            # Reemplazar DOKNR → popup validación → guardar
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
                res._log(f"  [PLANO] Guardado: '{doknr_actual}' → '{nuevo_plano}'")
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

            posiciones = self.bom_con_retry(zpla_usado, clases)
            res.posiciones_bom = posiciones
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

            posiciones = self.bom_con_retry(zpla_usado, clases)
            res.posiciones_bom = posiciones
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
            self.mm02_cambiar_plano(zfer_nuevo, res)

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
                self._navegar("ZMME0001")
                self._esperar(T_MEDIO)
            except Exception:
                pass

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
