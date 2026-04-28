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

import win32com.client
import time
import pyodbc
import datetime
import uuid
from dataclasses import dataclass, field
from typing import Optional

# ── Tiempos de espera ─────────────────────────────────────────────────────────
T_RAPIDO = 0.4
T_MEDIO  = 1.2
T_LENTO  = 2.5

_SAP_USER = "PROGRAING"

# ── BD Local ──────────────────────────────────────────────────────────────────
_DB_LOCAL_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    r"SERVER=localhost\SQLEXPRESS;"
    "DATABASE=MODULO_5;"
    "Trusted_Connection=yes;"
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
                          zplas_validos: list) -> tuple:
        """
        ZMME0001 → Homologar → Cambio de Color → F4 en ZPLA (SAP sugiere) →
        valida contra zplas_validos del combinador → doble clic → F8.
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
        self._esperar(T_MEDIO)

        # 7. Leer popup y seleccionar ZPLA que coincida con los del combinador
        _ID_POPUP_GRID = "wnd[1]/usr/cntlLO_CONTAINER0500/shellcont/shell"
        zpla_seleccionado = ""
        fila_seleccionada = 0

        try:
            grid_popup = self.session.findById(_ID_POPUP_GRID)
            n_filas    = grid_popup.RowCount
            print(f"    F4 ZPLA popup: {n_filas} sugerencias SAP")

            # Leer todos los ZPLAs que sugiere SAP
            zplas_sap = []
            for i in range(n_filas):
                val = ""
                for col in ("MATNR", "ZPLA", "ZPLARF", "MATERIAL"):
                    try:
                        val = str(grid_popup.GetCellValue(i, col) or "").strip()
                        if val:
                            break
                    except Exception:
                        pass
                zplas_sap.append(val)
                print(f"      Sugerencia SAP fila {i}: {val}")

            # Buscar la primera sugerencia de SAP que esté en nuestros ZPLAs válidos
            zplas_set = {z.strip() for z in zplas_validos if z.strip()}
            fila_seleccionada = 0   # por defecto tomar la primera sugerencia de SAP
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
        ses2.findById("wnd[0]").maximize()

        try:
            ses2.findById(self._ID_TCODE_BOX).text = "ZPPR0020"
            ses2.findById("wnd[0]").sendVKey(0)
            self._esperar(T_MEDIO)
            ses2.findById(self._ID_ZPPR_USER).text   = _SAP_USER
            ses2.findById(self._ID_ZPPR_CENTRO).text = "CO01"
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

        try:
            n_filas = grid.RowCount
            if not n_filas:
                return resultado
            try:
                cols = list(grid.ColumnOrder)
            except Exception:
                cols = []

            for i in range(n_filas):
                zfer_fila = ""
                for col in cols:
                    if "MAT_ZFER" in col.upper() or col.upper() in ("ZFER", "MATNR_ZFER"):
                        try:
                            zfer_fila = str(grid.GetCellValue(i, col) or "").strip()
                            if zfer_fila:
                                break
                        except Exception:
                            pass
                if zfer_fila != zfer_nuevo:
                    continue

                resultado["encontrado"] = True
                for col in cols:
                    if "MAT_ZPLA" in col.upper() or col.upper() in ("ZPLA", "MATNR_ZPLA"):
                        try:
                            v = str(grid.GetCellValue(i, col) or "").strip()
                            if v:
                                resultado["zpla"] = v
                                break
                        except Exception:
                            pass

                fases  = {}
                n_fase = 1
                for col in cols:
                    if col.upper().startswith("PHASE"):
                        try:
                            v = str(grid.GetCellValue(i, col) or "").strip()
                            fases[f"Fase {n_fase}"] = v
                            n_fase += 1
                        except Exception:
                            pass
                resultado["fases"] = fases
                n_s = sum(1 for v in fases.values() if v.upper() == "S")
                n_e = sum(1 for v in fases.values() if v.upper() == "E")
                print(f"    ZPPR0020 fila {i}: ZPLA={resultado['zpla']} | S={n_s} | E={n_e}")
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
        self.session.findById(self._ID_BTN_COMP).press()
        self._esperar(T_MEDIO)

        posiciones = []
        # Intentar GuiTableControl clásico
        try:
            tabla = self.session.findById("wnd[1]/usr/tblZMME0001T_COMP")
            for i in range(tabla.RowCount):
                try:
                    pos = tabla.GetCell(i, 0).text.strip()
                    if pos:
                        posiciones.append(pos)
                except Exception:
                    pass
        except Exception:
            pass

        # Fallback GuiGridView
        if not posiciones:
            try:
                grid = self.session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell")
                for i in range(grid.RowCount):
                    try:
                        pos = grid.GetCellValue(i, "POSNR").strip()
                        if pos:
                            posiciones.append(pos)
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

        vistos = set()
        return [p for p in posiciones if not (p in vistos or vistos.add(p))]

    def zmme0001_agregar_filas_bom(self, posiciones: list, zpla: str,
                                    clases_dict: dict = None):
        """
        clases_dict: {posicion: clase} obtenido de _leer_clases_zpla_sap().
        Si no se pasa, queda vacío (sin clase destino).
        """
        if clases_dict is None:
            clases_dict = {}
        for idx, pos in enumerate(posiciones):
            self.session.findById(self._ID_BTN_INSERT).press()
            self._esperar(T_RAPIDO)
            # SAP espera la posición sin ceros a la izquierda (0205 → 205)
            pos_sin_ceros = str(int(pos)) if pos.isdigit() else pos
            self.session.findById(
                f"{self._ID_TBL_LISTA}/txtWA_LISTA-POSNR[0,{idx}]"
            ).text = pos_sin_ceros
            self._esperar(T_RAPIDO)
            # Buscar clase con o sin padding
            clase = clases_dict.get(pos.zfill(4), clases_dict.get(pos, ""))
            if clase:
                self.session.findById(
                    f"{self._ID_TBL_LISTA}/ctxtWA_LISTA-CLASE_DESTINO[3,{idx}]"
                ).text = clase
                self._esperar(T_RAPIDO)
            print(f"    Fila {idx}: POS={pos_sin_ceros} CLASE={clase or '(sin clase)'}")

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
            cn = pyodbc.connect(_DB_LOCAL_STR, autocommit=True)
            cn.cursor().execute(
                "INSERT INTO dbo.M5_LogEjecucion "
                "(batch_id, pedido_origen, tipo_pieza, formula, color_codigo, acero_variante, "
                " estado, detalle_error, fecha_inicio, fecha_fin) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (res.batch_id, res.zfer_base, "", "", res.color_codigo, "",
                 res.estado, res.error or None,
                 res.fecha_inicio, res.fecha_fin)
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

    # ── Procesar una combinación ──────────────────────────────────────────────

    def procesar(self, zfer_base: str, color_codigo: str, color_nombre: str,
                 franja: str = "00", pn_base: str = "", zpla: str = "",
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

            # PASO 2 — ZMME0001 ejecutar
            _cb(2, f"Creando nueva variante de color en SAP (ZMME0001) — color {p_color}")
            # zpla puede ser string simple o lista separada por comas
            zplas_validos = [z.strip() for z in str(zpla).split(",") if z.strip()]
            zfer_nuevo, zfor_nuevo, zpla_usado = self.zmme0001_ejecutar(
                zfer_base, p_color, p_franj, zplas_validos
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

            posiciones = self.zmme0001_leer_posiciones_popup()
            res.posiciones_bom = posiciones
            res._log(f"  Posiciones BOM ({len(posiciones)}): {posiciones}")

            if posiciones and zpla_usado:
                # Leer posicion→clase desde SAP (sesión auxiliar)
                clases = self._leer_clases_zpla_sap(zpla_usado)
                res._log(f"  Clases leídas desde SAP: {clases}")
                self.zmme0001_agregar_filas_bom(posiciones, zpla_usado, clases)
            elif posiciones and not zpla_usado:
                res._log("  [WARN] Sin ZPLA para leer clases")
                self.zmme0001_agregar_filas_bom(posiciones, "", {})

            ok_bom = self.zmme0001_segunda_comparar_y_copy()
            if not ok_bom:
                raise RuntimeError("Segunda Comparar BOM devolvió error — revisar Clave Destino")

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


# ── Función de entrada (usada desde app.py vía threading) ────────────────────

def procesar_combinacion(zfer_base: str, color_codigo: str, color_nombre: str,
                         franja: str = "00", pn_base: str = "",
                         zpla: str = "", step_callback=None) -> ResultadoItem:
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
                         step_callback=step_callback)
