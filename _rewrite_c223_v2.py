content = open('sap_mantenimiento.py', encoding='utf-8').read()

start = content.find('\ndef c223_mantenimiento(')
end   = content.find('\ndef zinpg0004_actualizar(')

nueva_c223 = '''
def c223_mantenimiento(zfers: list, hr_id: str) -> dict:
    """
    C223 en sesion auxiliar: entra ZFERs uno a uno en ctxtMKAL-MATNR + Enter,
    luego llena PRDAT (fecha hoy via F4 calendario) y PLNNR (HR) por cada fila visible.
    Basado en VBS real grabado.
    """
    import datetime as _dt
    ses0, ses2, conn, err = _conectar_sap_nueva_sesion()
    if ses2 is None:
        return {"ok": False, "error": f"No se pudo crear sesion auxiliar SAP: {err}", "detalles": []}

    detalles = []
    def _log(msg):
        print(f"[C223-MANT] {msg}")
        detalles.append(msg)

    import datetime as _dt2
    hoy        = _dt2.datetime.today().strftime("%d.%m.%Y")   # DD.MM.YYYY para campos texto
    hoy_cal    = _dt2.datetime.today().strftime("%Y%m%d")     # YYYYMMDD para el calendario F4
    _TBL       = "wnd[0]/usr/ssubSUBSCR_1200:SAPLCMFV:1200/tblSAPLCMFVT_MKAL"
    _MATNR     = "wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-MATNR"
    _WERKS     = "wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-WERKS"

    _log(f"C223: {len(zfers)} ZFERs | HR={hr_id} | fecha={hoy}")

    try:
        # Navegar a C223 con reintentos
        nav_ok = False
        for _n in range(15):
            try:
                try: ses2.findById("wnd[1]").sendVKey(0)
                except Exception: pass
                ses2.findById("wnd[0]/tbar[0]/okcd").text = "C223"
                ses2.findById("wnd[0]").sendVKey(0)
                _esperar_ocupado(ses2, T_MEDIO)
                nav_ok = True
                _log(f"Nav C223 OK (intento {_n+1})")
                break
            except Exception as e_nav:
                _log(f"[WAIT] intento {_n+1}: {e_nav}")
                time.sleep(1.0)
        if not nav_ok:
            return {"ok": False, "error": "No se pudo navegar a C223", "detalles": detalles}

        # Cerrar popups post-nav
        for _ in range(3):
            try: ses2.findById("wnd[1]").sendVKey(0); time.sleep(0.3)
            except Exception: break

        # Planta
        _log(">> WERKS = CO01")
        ses2.findById(_WERKS).text = "CO01"

        # Entrar ZFERs uno a uno en el campo MATNR + Enter
        _log(f">> Entrando {len(zfers)} ZFERs en MATNR...")
        for i, zfer in enumerate(zfers):
            ses2.findById(_MATNR).text = zfer
            ses2.findById(_MATNR).setFocus()
            ses2.findById(_MATNR).caretPosition = len(zfer)
            ses2.findById("wnd[0]").sendVKey(0)   # Enter
            _esperar_ocupado(ses2, 2.0)
            # Cerrar popups que aparezcan
            for _ in range(2):
                try: ses2.findById("wnd[1]").sendVKey(0); time.sleep(0.3)
                except Exception: break
            _log(f"   [{i+1}/{len(zfers)}] {zfer} OK")

        # Leer tabla
        try:
            tbl      = ses2.findById(_TBL)
            vis_rows = tbl.VisibleRowCount
            max_sb   = tbl.verticalScrollbar.maximum
        except Exception as e_tbl:
            return {"ok": False, "error": f"Tabla C223 no encontrada: {e_tbl}", "detalles": detalles}

        _log(f"Tabla: vis={vis_rows} max_scroll={max_sb}")

        # Ajustar ancho columna 16
        try:
            ses2.findById(_TBL).columns.elementAt(16).width = 15
        except Exception:
            pass

        # Llenar PRDAT (col 9) y PLNNR (col 16) en todas las filas
        # Para PRDAT: intentar texto directo primero, si falla usar F4 calendario
        for sp in range(0, max_sb + 1, max(1, vis_rows)):
            try:
                ses2.findById(_TBL).verticalScrollbar.position = sp
                time.sleep(0.15)
            except Exception:
                pass

            for vis in range(vis_rows):
                # PRDAT[9,vis] — fecha
                try:
                    campo_prdat = ses2.findById(f"{_TBL}/ctxtMKAL_EXPAND-PRDAT[9,{vis}]")
                    campo_prdat.text = hoy
                except Exception:
                    # Fallback: F4 calendario
                    try:
                        ses2.findById(f"{_TBL}/ctxtMKAL_EXPAND-PRDAT[9,{vis}]").setFocus()
                        ses2.findById(f"{_TBL}/ctxtMKAL_EXPAND-PRDAT[9,{vis}]").caretPosition = 2
                        ses2.findById("wnd[0]").sendVKey(4)   # F4
                        time.sleep(0.5)
                        ses2.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").focusDate = hoy_cal
                        ses2.findById("wnd[1]/tbar[0]/btn[0]").press()
                        time.sleep(0.3)
                    except Exception as e_prdat:
                        _log(f"[WARN] PRDAT[{vis}]: {e_prdat}")

                # PLNNR[16,vis] — HR
                try:
                    ses2.findById(f"{_TBL}/ctxtMKAL_EXPAND-PLNNR[16,{vis}]").text = hr_id
                    ses2.findById(f"{_TBL}/ctxtMKAL_EXPAND-PLNNR[16,{vis}]").setFocus()
                    ses2.findById(f"{_TBL}/ctxtMKAL_EXPAND-PLNNR[16,{vis}]").caretPosition = len(hr_id)
                    ses2.findById("wnd[0]").sendVKey(0)   # Enter para confirmar
                    _esperar_ocupado(ses2, 1.0)
                    for _ in range(2):
                        try: ses2.findById("wnd[1]").sendVKey(0); time.sleep(0.3)
                        except Exception: break
                except Exception as e_plnnr:
                    _log(f"[WARN] PLNNR[{vis}]: {e_plnnr}")

        # MAAL
        _log("MAAL...")
        ses2.findById("wnd[0]/usr/ssubSUBSCR_1200:SAPLCMFV:1200/btnMAAL").press()
        _esperar_ocupado(ses2, T_MEDIO)
        try: ses2.findById("wnd[1]").sendVKey(0)
        except Exception: pass

        # PRUEFEN
        _log("PRUEFEN...")
        ses2.findById("wnd[0]/usr/ssubSUBSCR_1200:SAPLCMFV:1200/btnPRUEFEN").press()
        _esperar_ocupado(ses2, T_MEDIO)

        # btn[3]
        try:
            ses2.findById("wnd[0]/tbar[0]/btn[3]").press()
            _esperar_ocupado(ses2, T_RAPIDO)
        except Exception:
            pass

        sbar_txt, sbar_tipo = _leer_sbar(ses2)
        _log(f"Pre-guardado: \'{sbar_txt}\' tipo={sbar_tipo}")
        if sbar_tipo == "E":
            return {"ok": False, "error": f"Error C223: {sbar_txt}", "detalles": detalles}

        # Guardar
        _log("Guardando btn[11]...")
        ses2.findById("wnd[0]/tbar[0]/btn[11]").press()
        _esperar_ocupado(ses2, T_LENTO)

        msg_final, tipo_final = _leer_sbar(ses2)
        _log(f"Guardado: \'{msg_final}\' tipo={tipo_final}")
        try: ses2.findById("wnd[1]").sendVKey(0)
        except Exception: pass

        if tipo_final == "E":
            return {"ok": False, "error": f"Error guardando C223: {msg_final}", "detalles": detalles}

        return {"ok": True, "mensaje": msg_final or "C223 OK", "error": "", "detalles": detalles}

    except Exception as e:
        _log(f"Excepcion: {e}")
        return {"ok": False, "error": str(e), "detalles": detalles}

    finally:
        try:
            ses2.findById("wnd[0]/tbar[0]/okcd").text = "/i"
            ses2.findById("wnd[0]").sendVKey(0)
        except Exception:
            try: ses2.findById("wnd[0]").close()
            except Exception: pass
        _log("Sesion auxiliar C223 cerrada.")

'''

new_content = content[:start] + nueva_c223 + content[end:]
open('sap_mantenimiento.py', 'w', encoding='utf-8').write(new_content)
print('OK')
