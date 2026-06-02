import datetime

code = '''

def c223_mantenimiento(zfers: list, hr_id: str) -> dict:
    """C223: fecha hoy + PLNNR=hr_id para lista de ZFERs via seleccion multiple."""
    import datetime as _dt
    ses, err = _conectar_sap()
    if ses is None:
        return {"ok": False, "error": f"SAP GUI no disponible: {err}", "detalles": []}

    detalles = []
    def _log(msg):
        print(f"[C223-MANT] {msg}")
        detalles.append(msg)

    hoy = _dt.datetime.today().strftime("%d.%m.%Y")
    _log(f"C223: {len(zfers)} ZFERs | HR={hr_id} | fecha={hoy}")

    _TBL    = "wnd[0]/usr/ssubSUBSCR_1200:SAPLCMFV:1200/tblSAPLCMFVT_MKAL"
    _TBL_MV = "wnd[2]/usr/tabsTAB_STRIP/tabpSINGLE/ssubSCREEN:SAPLALDB:3010/tblSAPLALDBSINGLE"

    try:
        ses.findById("wnd[0]/tbar[0]/okcd").text = "/nC223"
        ses.findById("wnd[0]").sendVKey(0)
        _esperar_ocupado(ses, T_MEDIO)
        _cerrar_popup(ses)

        ses.findById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/btnG_ICON_SELECTION_FV").press()
        _esperar_ocupado(ses, T_MEDIO)

        ses.findById("wnd[1]/usr/btn%_RANG_MAT_%_APP_%-VALU_PUSH").press()
        _esperar_ocupado(ses, T_MEDIO)

        # Limpiar previos
        try:
            ses.findById("wnd[2]/tbar[0]/btn[24]").press()
            _esperar_ocupado(ses, T_RAPIDO)
        except Exception:
            pass

        # Leer filas visibles del multi-value
        try:
            vis_mv = ses.findById(_TBL_MV).VisibleRowCount
        except Exception:
            vis_mv = 20

        _log(f"Multi-value vis={vis_mv}, total ZFERs={len(zfers)}")
        for i, zfer in enumerate(zfers):
            vis_row = i % vis_mv
            if vis_row == 0 and i > 0:
                try:
                    ses.findById(_TBL_MV).verticalScrollbar.position = i
                    time.sleep(0.1)
                except Exception:
                    pass
            try:
                ses.findById(f"{_TBL_MV}/ctxtRSCSEL_255-SLOW_I[1,{vis_row}]").text = zfer
            except Exception as e_mv:
                _log(f"[WARN] mv fila {i}: {e_mv}")

        # Confirmar wnd[2] y wnd[1]
        ses.findById("wnd[2]/tbar[0]/btn[8]").press()
        _esperar_ocupado(ses, T_MEDIO)
        ses.findById("wnd[1]/tbar[0]/btn[8]").press()
        _esperar_ocupado(ses, T_LENTO)

        # Tabla principal
        try:
            tbl = ses.findById(_TBL)
            vis_rows = tbl.VisibleRowCount
            max_sb   = tbl.verticalScrollbar.maximum
        except Exception as e_tbl:
            return {"ok": False, "error": f"Tabla C223 no encontrada: {e_tbl}", "detalles": detalles}

        _log(f"Tabla C223: vis={vis_rows} max_scroll={max_sb}")

        # Ajustar ancho columna 16 para edicion
        try:
            ses.findById(_TBL).columns.elementAt(16).width = 15
        except Exception:
            pass

        # Recorrer todas las filas
        for sp in range(0, max_sb + 1, max(1, vis_rows)):
            try:
                ses.findById(_TBL).verticalScrollbar.position = sp
                time.sleep(0.15)
            except Exception:
                pass
            for vis in range(vis_rows):
                try:
                    ses.findById(f"{_TBL}/ctxtMKAL_EXPAND-ADATU[10,{vis}]").text = hoy
                except Exception:
                    pass
                try:
                    ses.findById(f"{_TBL}/ctxtMKAL_EXPAND-PLNNR[16,{vis}]").text = hr_id
                except Exception:
                    pass

        _log("MAAL...")
        try:
            ses.findById(f"{_TBL}/ctxtMKAL_EXPAND-ADATU[10,0]").setFocus()
            ses.findById(f"{_TBL}/ctxtMKAL_EXPAND-ADATU[10,0]").caretPosition = 10
        except Exception:
            pass

        ses.findById("wnd[0]/usr/ssubSUBSCR_1200:SAPLCMFV:1200/btnMAAL").press()
        _esperar_ocupado(ses, T_MEDIO)
        txt_pop, _ = _leer_popup(ses)
        if txt_pop:
            _log(f"Popup MAAL: {txt_pop}")
            _cerrar_popup(ses)

        _log("PRUEFEN...")
        ses.findById("wnd[0]/usr/ssubSUBSCR_1200:SAPLCMFV:1200/btnPRUEFEN").press()
        _esperar_ocupado(ses, T_MEDIO)

        try:
            ses.findById("wnd[0]/tbar[0]/btn[3]").press()
            _esperar_ocupado(ses, T_RAPIDO)
        except Exception:
            pass

        sbar_txt, sbar_tipo = _leer_sbar(ses)
        _log(f"Pre-guardado: \'{sbar_txt}\' tipo={sbar_tipo}")
        if sbar_tipo == "E":
            return {"ok": False, "error": f"Error C223: {sbar_txt}", "detalles": detalles}

        _log("Guardando btn[11]...")
        ses.findById("wnd[0]/tbar[0]/btn[11]").press()
        _esperar_ocupado(ses, T_LENTO)

        msg_final, tipo_final = _leer_sbar(ses)
        _log(f"Guardado: \'{msg_final}\' tipo={tipo_final}")
        txt_pop2, _ = _leer_popup(ses)
        if txt_pop2:
            _log(f"Popup post-guardado: {txt_pop2}")
            _cerrar_popup(ses)

        if tipo_final == "E":
            return {"ok": False, "error": f"Error guardando C223: {msg_final}", "detalles": detalles}

        return {"ok": True, "mensaje": msg_final or "C223 OK", "error": "", "detalles": detalles}

    except Exception as e:
        _log(f"Excepcion: {e}")
        return {"ok": False, "error": str(e), "detalles": detalles}


def zinpg0004_actualizar(zfers: list = None) -> dict:
    """ZINPG0004: actualiza version fabricacion. zfers=None ejecuta para todos."""
    ses, err = _conectar_sap()
    if ses is None:
        return {"ok": False, "error": f"SAP GUI no disponible: {err}", "detalles": []}

    detalles = []
    def _log(msg):
        print(f"[ZINPG0004] {msg}")
        detalles.append(msg)

    _TBL_MV = "wnd[1]/usr/tabsTAB_STRIP/tabpSINGLE/ssubSCREEN:SAPLALDB:3010/tblSAPLALDBSINGLE"

    try:
        ses.findById("wnd[0]/tbar[0]/okcd").text = "/nZINPG0004"
        ses.findById("wnd[0]").sendVKey(0)
        _esperar_ocupado(ses, T_MEDIO)
        _cerrar_popup(ses)

        ses.findById("wnd[0]/usr/ctxtPA_WERKS").text = "CO01"
        ses.findById("wnd[0]/usr/ctxtPA_VERID").text = "5000"
        ses.findById("wnd[0]/usr/ctxtPA_VERID").setFocus()
        ses.findById("wnd[0]/usr/ctxtPA_VERID").caretPosition = 4

        ses.findById("wnd[0]/usr/btn%_SO_MATNR_%_APP_%-VALU_PUSH").press()
        _esperar_ocupado(ses, T_MEDIO)

        if zfers:
            _log(f"Entrando {len(zfers)} ZFERs...")
            try:
                vis_mv = ses.findById(_TBL_MV).VisibleRowCount
            except Exception:
                vis_mv = 20
            for i, zfer in enumerate(zfers):
                vis_row = i % vis_mv
                if vis_row == 0 and i > 0:
                    try:
                        ses.findById(_TBL_MV).verticalScrollbar.position = i
                        time.sleep(0.1)
                    except Exception:
                        pass
                try:
                    ses.findById(f"{_TBL_MV}/ctxtRSCSEL_255-SLOW_I[1,{vis_row}]").text = zfer
                except Exception:
                    pass
        else:
            _log("Limpiando filtro (todos los materiales)...")
            try:
                ses.findById("wnd[1]/tbar[0]/btn[24]").press()
                _esperar_ocupado(ses, T_RAPIDO)
            except Exception:
                pass

        ses.findById("wnd[1]/tbar[0]/btn[8]").press()
        _esperar_ocupado(ses, T_MEDIO)

        _log("F8 Ejecutar...")
        ses.findById("wnd[0]/tbar[1]/btn[8]").press()
        _esperar_con_polling(ses, max_seg=300, intervalo=2.0)

        try:
            ses.findById("wnd[0]/tbar[1]/btn[20]").press()
            _esperar_ocupado(ses, T_MEDIO)
        except Exception as e20:
            _log(f"[WARN] btn[20]: {e20}")

        msg_final, tipo_final = _leer_sbar(ses)
        _log(f"Resultado: \'{msg_final}\' tipo={tipo_final}")

        if tipo_final == "E":
            return {"ok": False, "error": f"Error ZINPG0004: {msg_final}", "detalles": detalles}

        return {"ok": True, "mensaje": msg_final or "ZINPG0004 OK", "error": "", "detalles": detalles}

    except Exception as e:
        _log(f"Excepcion: {e}")
        return {"ok": False, "error": str(e), "detalles": detalles}
'''

with open('sap_mantenimiento.py', 'a', encoding='utf-8') as f:
    f.write(code)
print('OK')
