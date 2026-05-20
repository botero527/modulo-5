"""
test_ca02.py — Prueba aislada de CA02 desasignar + asignar HR.
getCellValue no funciona en esta tabla → lee celdas via findById por cada posición de scroll.

Uso:
    py test_ca02.py desasignar <zfer_nuevo>
    py test_ca02.py asignar    <zfer_nuevo> <id_hruta>
    py test_ca02.py ambos      <zfer_nuevo> <id_hruta>

Ejemplo:
    py test_ca02.py ambos 700163269 53025306
"""
import sys, time, win32com.client

_TBL = "wnd[1]/usr/tblSAPLCZDITCTRL_1010"

def get_session():
    sap  = win32com.client.GetObject("SAPGUI")
    app  = sap.GetScriptingEngine
    conn = app.Children(0)
    return conn.Children(0)

def esperar(seg=1.0):
    time.sleep(seg)

def leer_matnr(session, vis_row: int) -> str:
    """Lee MATNR de la fila visual vis_row via findById."""
    try:
        return str(session.findById(f"{_TBL}/ctxtMAPL-MATNR[2,{vis_row}]").text or "").strip()
    except Exception:
        return ""

def abrir_ca02_con_matnr(session, zfer_nuevo: str):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nca02"
    session.findById("wnd[0]").sendVKey(0)
    esperar(2)
    session.findById("wnd[0]/usr/ctxtRC27M-MATNR").text = zfer_nuevo
    session.findById("wnd[0]/usr/ctxtRC27M-WERKS").text = "CO01"
    session.findById("wnd[0]/usr/ctxtRC27M-WERKS").caretPosition = 4
    session.findById("wnd[0]").sendVKey(0)
    esperar(1)
    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    esperar(3)
    session.findById("wnd[0]/tbar[1]/btn[31]").press()
    esperar(3)

def abrir_ca02_con_plnnr(session, id_hruta: str):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nca02"
    session.findById("wnd[0]").sendVKey(0)
    esperar(2)
    session.findById("wnd[0]/usr/ctxtRC27M-MATNR").text = ""
    session.findById("wnd[0]/usr/ctxtRC27M-WERKS").text = "CO01"
    session.findById("wnd[0]/usr/ctxtRC271-PLNNR").text = str(id_hruta)
    session.findById("wnd[0]/usr/ctxtRC27M-WERKS").caretPosition = 4
    session.findById("wnd[0]").sendVKey(0)
    esperar(1)
    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    esperar(8)
    session.findById("wnd[0]/tbar[1]/btn[31]").press()
    esperar(8)

def get_tbl_info(session):
    """Retorna (tbl, total, vis_rows, max_scroll)."""
    tbl = session.findById(_TBL)
    total    = tbl.RowCount
    vis_rows = tbl.VisibleRowCount
    try:
        max_scroll = tbl.verticalScrollbar.maximum
    except Exception:
        max_scroll = max(0, total - vis_rows)
    return tbl, total, vis_rows, max_scroll

def set_scroll(tbl, pos: int):
    try:
        tbl.verticalScrollbar.position = pos
        esperar(0.2)
    except Exception as e:
        print(f"  [scroll] ERROR: {e}")

# ── CA02 DESASIGNAR ───────────────────────────────────────────────────────────
def ca02_desasignar(session, zfer_nuevo: str):
    print(f"\n=== CA02 DESASIGNAR {zfer_nuevo} ===")
    abrir_ca02_con_matnr(session, zfer_nuevo)

    tbl, total, vis_rows, max_scroll = get_tbl_info(session)
    print(f"  Tabla: total={total} vis={vis_rows} max_scroll={max_scroll}")

    # Escanear todas las posiciones de scroll buscando el MATNR
    fila_scroll = None
    fila_vis    = None
    for scroll_pos in range(max_scroll + 1):
        set_scroll(tbl, scroll_pos)
        for vis in range(vis_rows):
            val = leer_matnr(session, vis)
            if val == zfer_nuevo:
                fila_scroll = scroll_pos
                fila_vis    = vis
                print(f"  Encontrado: scroll={scroll_pos} vis_row={vis} MATNR={val}")
                break
        if fila_scroll is not None:
            break

    if fila_scroll is None:
        print(f"  [WARN] No se encontró {zfer_nuevo} (esperado para ZFER nuevo sin HR)")
        try: session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except: pass
        return False

    # Ya estamos en el scroll correcto, usar vis_row directamente
    id_matnr = f"{_TBL}/ctxtMAPL-MATNR[2,{fila_vis}]"
    session.findById(id_matnr).setFocus()
    session.findById(id_matnr).caretPosition = 9
    esperar(0.3)
    session.findById("wnd[1]/tbar[0]/btn[14]").press()
    esperar(1)
    for btn_id in ("wnd[2]/tbar[0]/btn[0]", "wnd[2]/usr/btnSPOP-OPTION1"):
        try:
            session.findById(btn_id).press()
            esperar(0.5)
        except:
            pass
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    esperar(1)
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    esperar(2)
    print("  CA02 DESASIGNAR OK")
    return True

# ── CA02 ASIGNAR ──────────────────────────────────────────────────────────────
def ca02_asignar(session, zfer_nuevo: str, id_hruta: str):
    print(f"\n=== CA02 ASIGNAR {zfer_nuevo} → HR {id_hruta} ===")
    abrir_ca02_con_plnnr(session, id_hruta)

    tbl, total, vis_rows, max_scroll = get_tbl_info(session)
    print(f"  Tabla: total={total} vis={vis_rows} max_scroll={max_scroll}")

    # Ir al final y buscar primera fila vacía (de abajo a arriba)
    fila_scroll = None
    fila_vis    = None
    for scroll_pos in range(max_scroll, -1, -1):
        set_scroll(tbl, scroll_pos)
        for vis in range(vis_rows - 1, -1, -1):
            val = leer_matnr(session, vis)
            if not val:
                fila_scroll = scroll_pos
                fila_vis    = vis
            else:
                # primera fila con dato desde abajo → las vacías son después de aquí
                if fila_scroll is not None:
                    break
        if fila_scroll is not None:
            break

    if fila_scroll is None:
        print("  ERROR: No hay fila vacía")
        try: session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except: pass
        return False

    print(f"  Fila vacía: scroll={fila_scroll} vis_row={fila_vis}")
    # Ya estamos en el scroll correcto
    _matnr = f"{_TBL}/ctxtMAPL-MATNR[2,{fila_vis}]"
    _werks  = f"{_TBL}/ctxtMAPL-WERKS[3,{fila_vis}]"
    _plnal  = f"{_TBL}/txtMAPL-PLNAL[0,{fila_vis}]"

    session.findById(_matnr).setFocus()
    esperar(0.3)
    try:
        session.findById(_plnal).text = "1"
    except Exception as e:
        print(f"  [WARN] PLNAL: {e}")
    session.findById(_matnr).text = zfer_nuevo
    session.findById(_werks).text  = "CO01"
    session.findById(_werks).caretPosition = 4
    esperar(0.5)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    esperar(2)
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    esperar(3)
    print("  CA02 ASIGNAR OK")
    return True

# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    modo = sys.argv[1].lower()
    zfer = sys.argv[2]
    session = get_session()
    session.findById("wnd[0]").maximize()

    if modo == "desasignar":
        ca02_desasignar(session, zfer)
    elif modo == "asignar":
        if len(sys.argv) < 4:
            print("Falta id_hruta")
            sys.exit(1)
        ca02_asignar(session, zfer, sys.argv[3])
    elif modo == "ambos":
        if len(sys.argv) < 4:
            print("Falta id_hruta")
            sys.exit(1)
        ca02_desasignar(session, zfer)
        ca02_asignar(session, zfer, sys.argv[3])
    else:
        print(f"Modo desconocido: {modo}")
        print(__doc__)
