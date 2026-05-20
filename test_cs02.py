"""
test_cs02.py — Prueba aislada de CS02 agregar posición acero.
Ejecutar con SAP GUI abierto y sesión activa.

Uso:
    py test_cs02.py <zfor> <pos_acero> <zhal>

Ejemplo (valores del log real):
    py test_cs02.py 501143183 0116 600006711
"""
import sys, time, win32com.client

_TBL = ("wnd[0]/usr/tabsTS_ITOV/tabpTCMA"
        "/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT")

def get_session():
    sap  = win32com.client.GetObject("SAPGUI")
    app  = sap.GetScriptingEngine
    conn = app.Children(0)
    return conn.Children(0)

def esperar(seg=1.0):
    time.sleep(seg)

def leer_posnr(session, vis):
    try:
        return str(session.findById(f"{_TBL}/txtRC29P-POSNR[0,{vis}]").text or "").strip()
    except Exception:
        return None   # None = fila no existe

def leer_postp(session, vis):
    try:
        return str(session.findById(f"{_TBL}/ctxtRC29P-POSTP[1,{vis}]").text or "").strip()
    except Exception:
        return None
    
def diag_tabla(session):
    print("\n── Diagnóstico tabla CS02 ──")
    try:
        tbl = session.findById(_TBL)
        print(f"  RowCount        = {tbl.RowCount}")
        print(f"  VisibleRowCount = {tbl.VisibleRowCount}")
        try:    print(f"  scrollbar max   = {tbl.verticalScrollbar.maximum}")
        except Exception as e: print(f"  scrollbar       ERROR: {e}")
    except Exception as e:
        print(f"  findById(_TBL)  ERROR: {e}")

    print("\n  Primeras 20 filas vía findById:")
    for vis in range(20):
        posnr = leer_posnr(session, vis)
        if posnr is None:
            print(f"    vis={vis}: (fila no existe — fin de tabla)")
            break
        postp = leer_postp(session, vis) or ""
        idnrk = ""
        try:
            idnrk = str(session.findById(f"{_TBL}/ctxtRC29P-IDNRK[2,{vis}]").text or "").strip()
        except Exception:
            pass
        print(f"    vis={vis}: POSNR='{posnr}' POSTP='{postp}' IDNRK='{idnrk}'")

def cs02_agregar(session, zfor, pos_acero, zhal):
    pos_num = str(pos_acero).lstrip("0") or pos_acero
    print(f"\n=== CS02 AGREGAR pos={pos_num} ZHAL={zhal} en ZFOR={zfor} ===")

    session.findById("wnd[0]/tbar[0]/okcd").text = "/ncs02"
    session.findById("wnd[0]").sendVKey(0)
    esperar(2)

    # Pantalla "Imagen inicial" — llenar MATNR y Utilización=1
    session.findById("wnd[0]/usr/ctxtRC29N-MATNR").text = zfor
    try:
        session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = "CO01"  # Centro
        print("  Centro = CO01 ✓")
    except Exception as e:
        print(f"  [WARN] Centro: {e}")
    try:
        session.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = "1"     # Utilización
        print("  Utilización = 1 ✓")
    except Exception as e:
        print(f"  [WARN] Utilización: {e}")
    session.findById("wnd[0]").sendVKey(0)
    esperar(2)

    # Seleccionar tab TCMA si no está activo
    try:
        session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA").select()
        esperar(1)
    except Exception as e:
        print(f"  [WARN] tab TCMA: {e}")

    # Diagnóstico
    diag_tabla(session)

    # Buscar primera fila vacía — escanear desde el FINAL hacia arriba
    fila_vis    = None
    fila_scroll = None
    try:
        tbl      = session.findById(_TBL)
        vis_rows = tbl.VisibleRowCount
        max_sb   = tbl.verticalScrollbar.maximum
    except Exception as e:
        print(f"  [WARN] tabla: {e}")
        tbl = None; vis_rows = 19; max_sb = 0

    # Filas vacías en CS02: POSTP está vacío (K=clase, L=material, ''=sin asignar)
    print(f"\n  Buscando fila vacía (max_scroll={max_sb} vis_rows={vis_rows})...")
    for sp in range(max_sb + 1):
        if tbl:
            try:
                tbl.verticalScrollbar.position = sp
                esperar(0.2)
            except Exception:
                pass
        for vis in range(vis_rows):
            posnr = leer_posnr(session, vis)
            if posnr is None:
                break   # fuera de rango
            postp = leer_postp(session, vis)
            if postp == "":   # fila disponible
                fila_vis    = vis
                fila_scroll = sp
                print(f"  Fila vacía encontrada: vis={vis} scroll={sp} POSNR='{posnr}'")
                break
        if fila_vis is not None:
            break

    if fila_vis is not None:
        print(f"  Fila vacía: vis={fila_vis} scroll={fila_scroll}")

    if fila_vis is None:
        print("  ERROR: No se encontró fila vacía")
        return False

    # Escribir
    print(f"\n  Escribiendo en vis_row={fila_vis}...")
    try:
        session.findById(f"{_TBL}/txtRC29P-POSNR[0,{fila_vis}]").text = pos_num
        print(f"    POSNR = {pos_num} ✓")
    except Exception as e:
        print(f"    POSNR ERROR: {e}")

    try:
        session.findById(f"{_TBL}/ctxtRC29P-POSTP[1,{fila_vis}]").text = "l"
        print(f"    POSTP = l ✓")
    except Exception as e:
        print(f"    POSTP ERROR: {e}")

    try:
        session.findById(f"{_TBL}/ctxtRC29P-IDNRK[2,{fila_vis}]").text = zhal
        session.findById(f"{_TBL}/ctxtRC29P-IDNRK[2,{fila_vis}]").setFocus()
        session.findById(f"{_TBL}/ctxtRC29P-IDNRK[2,{fila_vis}]").caretPosition = len(zhal)
        print(f"    IDNRK = {zhal} ✓")
    except Exception as e:
        print(f"    IDNRK ERROR: {e}")

    print("  sendVKey(0) para confirmar...")
    session.findById("wnd[0]").sendVKey(0)
    esperar(2)

    # Sub-screen MENGE
    try:
        menge_id = "wnd[0]/usr/subPOS_PHPT:SAPLCSDI:0830/txtRC29P-MENGE"
        session.findById(menge_id).text = "1"
        session.findById(menge_id).caretPosition = 1
        session.findById("wnd[0]").sendVKey(0)
        esperar(1)
        print("    MENGE = 1 ✓")
    except Exception as e:
        print(f"    MENGE ERROR (puede ser normal si ya se llenó): {e}")

    # Guardar
    print("  Guardando con btn[11]...")
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    esperar(3)
    print("  CS02 AGREGAR OK")
    return True

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)

    zfor      = sys.argv[1]
    pos_acero = sys.argv[2]
    zhal      = sys.argv[3]

    session = get_session()
    session.findById("wnd[0]").maximize()
    cs02_agregar(session, zfor, pos_acero, zhal)
