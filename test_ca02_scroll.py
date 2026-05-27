"""
test_ca02_scroll.py — Diagnóstico CA02 popup materiales
Deja el popup ABIERTO en SAP antes de correr esto.
"""
import win32com.client, time

_TBL = "wnd[1]/usr/tblSAPLCZDITCTRL_1010"

def conectar():
    try:
        sap = win32com.client.GetObject("SAPGUI")
        ses = sap.GetScriptingEngine.Children(0).Children(0)
        print(f"[OK] SAP: {ses.Info.SystemName}")
        return ses
    except Exception as e:
        print(f"[ERROR] {e}"); return None

def listar_controles_wnd1(ses):
    """Enumera todos los controles de wnd[1] para encontrar botones y campos."""
    print("\n=== CONTROLES wnd[1] ===")
    try:
        wnd = ses.findById("wnd[1]")
        print(f"  Título: '{wnd.Text}'")
    except Exception as e:
        print(f"  [ERROR] wnd[1]: {e}"); return
    # tbar[0]
    for i in range(30):
        for tbar in ("tbar[0]", "tbar[1]"):
            try:
                btn = ses.findById(f"wnd[1]/{tbar}/btn[{i}]")
                print(f"  {tbar}/btn[{i}]: '{btn.Text}' tooltip='{btn.Tooltip}'")
            except Exception:
                pass
    # usr — buscar campos de texto
    for ctrl_name in ("txtFIRST_LINE", "txtPOSITION", "txtENTRADA", "txtLINE"):
        try:
            ctrl = ses.findById(f"wnd[1]/usr/{ctrl_name}")
            print(f"  usr/{ctrl_name}: '{ctrl.Text}'")
        except Exception:
            pass

def probar_scroll(ses, zfer_buscar: str):
    tbl = ses.findById(_TBL)
    vis_rows   = tbl.VisibleRowCount
    max_scroll = tbl.verticalScrollbar.maximum
    print(f"\n=== TABLA: vis={vis_rows} max_scroll={max_scroll} total_est={max_scroll+vis_rows} ===")

    def leer(sp, vis):
        try:
            # Re-fetch tbl en cada lectura para evitar objeto stale
            t = ses.findById(_TBL)
            return str(t.findById(f"ctxtMAPL-MATNR[2,{vis}]").text or "").strip()
        except Exception:
            try:
                return str(ses.findById(f"{_TBL}/ctxtMAPL-MATNR[2,{vis}]").text or "").strip()
            except Exception:
                return ""

    def scroll_a(sp):
        """Intenta scroll por 3 métodos distintos y reporta cuál funcionó."""
        # Método 1: re-fetch + verticalScrollbar.position
        try:
            t = ses.findById(_TBL)
            t.verticalScrollbar.position = sp
            time.sleep(0.2)
            return "verticalScrollbar"
        except Exception as e1:
            pass
        # Método 2: firstVisibleRow
        try:
            t = ses.findById(_TBL)
            t.firstVisibleRow = sp
            time.sleep(0.2)
            return "firstVisibleRow"
        except Exception as e2:
            pass
        # Método 3: Page Down repetido desde posición 0
        return None

    # Probar métodos de scroll
    print("\n[PRUEBA SCROLL]")
    for test_sp in (0, 100, 400, max_scroll):
        m = scroll_a(test_sp)
        # leer primera fila después del scroll
        val0 = leer(test_sp, 0)
        print(f"  scroll={test_sp:4d} método={m or 'FALLÓ'} → fila0='{val0}'")

    # Decidir método
    metodo = None
    for test_sp in (50,):
        m = scroll_a(test_sp)
        if m:
            metodo = m
            break

    if not metodo:
        print("\n[ERROR] Ningún método de scroll funciona. Intentando Page Down...")
        # Método PgDn: volver a 0 y bajar con teclas
        scroll_a(0)
        encontrado_pgdn = False
        for _ in range(500):
            for vis in range(vis_rows):
                val = leer(0, vis)
                if val == zfer_buscar:
                    print(f"[ENCONTRADO vía PgDn] vis_row={vis}")
                    encontrado_pgdn = True
                    break
            if encontrado_pgdn:
                break
            ses.findById("wnd[1]").sendVKey(82)  # Page Down
            time.sleep(0.1)
        if not encontrado_pgdn:
            print(f"[NO ENCONTRADO] {zfer_buscar} no apareció con PgDn")
        return

    # Scan con el método que funcionó
    print(f"\n[SCAN] método='{metodo}' buscando {zfer_buscar}...")
    encontrado = False
    for sp in range(0, max_scroll + 1, max(1, vis_rows)):
        scroll_a(sp)
        for vis in range(vis_rows):
            val = leer(sp, vis)
            if val == zfer_buscar:
                print(f"\n*** ENCONTRADO: scroll={sp} vis_row={vis} ***")
                encontrado = True
                break
        if encontrado:
            break
        if (sp // vis_rows) % 10 == 0:
            rango = leer(sp, 0)
            print(f"  sp={sp:4d}: '{rango}'")

    if not encontrado:
        print(f"\n[NO ENCONTRADO en 0..{max_scroll}]")
        print("  → El material puede estar más allá del max_scroll.")
        print("  → Probando saltar directo al final con 'Posicionar'...")
        # Intentar usar el campo Entrada + botón Posicionar
        for campo_id in ("wnd[1]/usr/txtFIRST_LINE", "wnd[1]/usr/txtENTRADA",
                         "wnd[1]/usr/txtLINE", "wnd[1]/usr/txtPOSITION"):
            try:
                ctrl = ses.findById(campo_id)
                ctrl.text = zfer_buscar
                print(f"  Escribí {zfer_buscar} en {campo_id}")
                # Buscar botón Posicionar
                for bi in range(25):
                    try:
                        btn = ses.findById(f"wnd[1]/tbar[0]/btn[{bi}]")
                        tt = btn.Tooltip.lower()
                        if "posicionar" in tt or "position" in tt or "goto" in tt:
                            btn.press()
                            time.sleep(0.3)
                            val0 = leer(0, 0)
                            print(f"  Después de Posicionar: fila0='{val0}'")
                            if val0 == zfer_buscar:
                                print(f"  *** ENCONTRADO vía Posicionar ***")
                            break
                    except Exception:
                        pass
                break
            except Exception:
                pass

if __name__ == "__main__":
    ses = conectar()
    if not ses: exit(1)
    listar_controles_wnd1(ses)
    zfer = input("\nZFER a buscar: ").strip()
    if zfer:
        probar_scroll(ses, zfer)
    input("\nEnter para cerrar...")
