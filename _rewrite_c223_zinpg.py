
content = open('sap_mantenimiento.py', encoding='utf-8').read()

# Find and replace c223_mantenimiento + zingp0004_actualizar completely
import re

# Find start of c223_mantenimiento
start_c223 = content.find('\ndef c223_mantenimiento(')
# Find start of zinpg0004
start_zinpg = content.find('\ndef zinpg0004_actualizar(')
# Find end of file
end_of_file = len(content)

new_fns = '''

def _conectar_sap_nueva_sesion():
    """
    Abre una sesion auxiliar nueva en SAP (igual que hace el homologador con ZPPR0020).
    Retorna (ses_principal, ses_nueva, app, error).
    ses_nueva es la sesion limpia para usar. Cerrar con ses_nueva.findById("wnd[0]/tbar[0]/okcd").text="/i".
    """
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        app     = sap_gui.GetScriptingEngine
        conn    = app.Children(0)
        ses0    = conn.Children(0)   # sesion principal (puede estar bloqueada)
        ses0.createSession()
        time.sleep(3)
        # La nueva sesion es la ultima
        ses_nueva = conn.Children(conn.Children.Count - 1)
        ses_nueva.findById("wnd[0]").maximize()
        return ses0, ses_nueva, conn, None
    except Exception as e:
        return None, None, None, str(e)


def c223_mantenimiento(zfers: list, hr_id: str) -> dict:
    """C223 en sesion auxiliar: fecha hoy + PLNNR=hr_id para lista de ZFERs."""
    import datetime as _dt
    ses0, ses2, conn, err = _conectar_sap_nueva_sesion()
    if ses2 is None:
        return {"ok": False, "error": f"No se pudo crear sesion auxiliar SAP: {err}", "detalles": []}

    detalles = []
    def _log(msg):
        print(f"[C223-MANT] {msg}")
        detalles.append(msg)

    hoy = _dt.datetime.today().strftime("%d.%m.%Y")
    _log(f"C223 sesion auxiliar: {len(zfers)} ZFERs | HR={hr_id} | fecha={hoy}")

    _TBL    = "wnd[0]/usr/ssubSUBSCR_1200:SAPLCMFV:1200/tblSAPLCMFVT_MKAL"
    _TBL_MV = "wnd[2]/usr/tabsTAB_STRIP/tabpSINGLE/ssubSCREEN:SAPLALDB:3010/tblSAPLALDBSINGLE"

    try:
        # Navegar a C223 en la sesion nueva (limpia)
        ses2.findById("wnd[0]/tbar[0]/okcd").text = "C223"
        ses2.findById("wnd[0]").sendVKey(0)
        _esperar_ocupado(ses2, T_MEDIO)

        # Cerrar popups post-navegacion
        for _ in range(3):
            try:
                ses2.findById("wnd[1]").sendVKey(0)
                time.sleep(0.3)
            except Exception:
                break

        # Abrir seleccion de materiales
        ses2.findById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/btnG_ICON_SELECTION_FV").press()
        _esperar_ocupado(ses2, T_MEDIO)

        # Abrir entrada multiple
        ses2.findById("wnd[1]/usr/btn%_RANG_MAT_%_APP_%-VALU_PUSH").press()
        _esperar_ocupado(ses2, T_MEDIO)

        # Limpiar entradas previas
        try:
            ses2.findById("wnd[2]/tbar[0]/btn[24]").press()
            _esperar_ocupado(ses2, T_RAPIDO)
        except Exception:
            pass

        # Entrar ZFERs en la tabla del multi-value
        try:
            vis_mv = ses2.findById(_TBL_MV).VisibleRowCount
        except Exception:
            vis_mv = 20

        _log(f"Multi-value vis={vis_mv}, ZFERs={len(zfers)}")
        for i, zfer in enumerate(zfers):
            vis_row = i % vis_mv
            if vis_row == 0 and i > 0:
                try:
                    ses2.findById(_TBL_MV).verticalScrollbar.position = i
                    time.sleep(0.1)
                except Exception:
                    pass
            try:
                ses2.findById(f"{_TBL_MV}/ctxtRSCSEL_255-SLOW_I[1,{vis_row}]").text = zfer
            except Exception as e_mv:
                _log(f"[WARN] mv fila {i} ({zfer}): {e_mv}")

        # Confirmar wnd[2] y wnd[1]
        ses2.findById("wnd[2]/tbar[0]/btn[8]").press()
        _esperar_ocupado(ses2, T_MEDIO)
        ses2.findById("wnd[1]/tbar[0]/btn[8]").press()
        _esperar_ocupado(ses2, T_LENTO)

        # Tabla principal
        try:
            tbl      = ses2.findById(_TBL)
            vis_rows = tbl.VisibleRowCount
            max_sb   = tbl.verticalScrollbar.maximum
        except Exception as e_tbl:
            return {"ok": False, "error": f"Tabla C223 no encontrada: {e_tbl}", "detalles": detalles}

        _log(f"Tabla C223: vis={vis_rows} max_scroll={max_sb}")

        # Ajustar columna 16
        try:
            ses2.findById(_TBL).columns.elementAt(16).width = 15
        except Exception:
            pass

        # Llenar fecha y HR en todas las filas
        for sp in range(0, max_sb + 1, max(1, vis_rows)):
            try:
                ses2.findById(_TBL).verticalScrollbar.position = sp
                time.sleep(0.15)
            except Exception:
                pass
            for vis in range(vis_rows):
                try:
                    ses2.findById(f"{_TBL}/ctxtMKAL_EXPAND-ADATU[10,{vis}]").text = hoy
                except Exception:
                    pass
                try:
                    ses2.findById(f"{_TBL}/ctxtMKAL_EXPAND-PLNNR[16,{vis}]").text = hr_id
                except Exception:
                    pass

        # Foco en ultima celda
        try:
            ses2.findById(f"{_TBL}/ctxtMKAL_EXPAND-ADATU[10,0]").setFocus()
            ses2.findById(f"{_TBL}/ctxtMKAL_EXPAND-ADATU[10,0]").caretPosition = 10
        except Exception:
            pass

        # MAAL
        _log("MAAL...")
        ses2.findById("wnd[0]/usr/ssubSUBSCR_1200:SAPLCMFV:1200/btnMAAL").press()
        _esperar_ocupado(ses2, T_MEDIO)
        try:
            ses2.findById("wnd[1]").sendVKey(0)
        except Exception:
            pass

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

        # Guardar btn[11]
        _log("Guardando...")
        ses2.findById("wnd[0]/tbar[0]/btn[11]").press()
        _esperar_ocupado(ses2, T_LENTO)

        msg_final, tipo_final = _leer_sbar(ses2)
        _log(f"Guardado: \'{msg_final}\' tipo={tipo_final}")

        try:
            ses2.findById("wnd[1]").sendVKey(0)
        except Exception:
            pass

        if tipo_final == "E":
            return {"ok": False, "error": f"Error guardando C223: {msg_final}", "detalles": detalles}

        return {"ok": True, "mensaje": msg_final or "C223 OK", "error": "", "detalles": detalles}

    except Exception as e:
        _log(f"Excepcion: {e}")
        return {"ok": False, "error": str(e), "detalles": detalles}

    finally:
        # Cerrar sesion auxiliar
        try:
            ses2.findById("wnd[0]/tbar[0]/okcd").text = "/i"
            ses2.findById("wnd[0]").sendVKey(0)
        except Exception:
            try:
                ses2.findById("wnd[0]").close()
            except Exception:
                pass
        _log("Sesion auxiliar C223 cerrada.")


def zinpg0004_actualizar(zfers: list = None) -> dict:
    """ZINPG0004 en sesion auxiliar. zfers=None ejecuta para todos."""
    ses0, ses2, conn, err = _conectar_sap_nueva_sesion()
    if ses2 is None:
        return {"ok": False, "error": f"No se pudo crear sesion auxiliar SAP: {err}", "detalles": []}

    detalles = []
    def _log(msg):
        print(f"[ZINPG0004] {msg}")
        detalles.append(msg)

    _TBL_MV = "wnd[1]/usr/tabsTAB_STRIP/tabpSINGLE/ssubSCREEN:SAPLALDB:3010/tblSAPLALDBSINGLE"

    try:
        _log("Navegando a ZINPG0004 en sesion auxiliar...")
        ses2.findById("wnd[0]/tbar[0]/okcd").text = "ZINPG0004"
        ses2.findById("wnd[0]").sendVKey(0)
        _esperar_ocupado(ses2, T_MEDIO)

        for _ in range(3):
            try:
                ses2.findById("wnd[1]").sendVKey(0)
                time.sleep(0.3)
            except Exception:
                break

        ses2.findById("wnd[0]/usr/ctxtPA_WERKS").text = "CO01"
        ses2.findById("wnd[0]/usr/ctxtPA_VERID").text = "5000"
        ses2.findById("wnd[0]/usr/ctxtPA_VERID").setFocus()
        ses2.findById("wnd[0]/usr/ctxtPA_VERID").caretPosition = 4

        ses2.findById("wnd[0]/usr/btn%_SO_MATNR_%_APP_%-VALU_PUSH").press()
        _esperar_ocupado(ses2, T_MEDIO)

        if zfers:
            _log(f"Entrando {len(zfers)} ZFERs...")
            try:
                vis_mv = ses2.findById(_TBL_MV).VisibleRowCount
            except Exception:
                vis_mv = 20
            for i, zfer in enumerate(zfers):
                vis_row = i % vis_mv
                if vis_row == 0 and i > 0:
                    try:
                        ses2.findById(_TBL_MV).verticalScrollbar.position = i
                        time.sleep(0.1)
                    except Exception:
                        pass
                try:
                    ses2.findById(f"{_TBL_MV}/ctxtRSCSEL_255-SLOW_I[1,{vis_row}]").text = zfer
                except Exception:
                    pass
        else:
            _log("Limpiando filtro (todos los materiales)...")
            try:
                ses2.findById("wnd[1]/tbar[0]/btn[24]").press()
                _esperar_ocupado(ses2, T_RAPIDO)
            except Exception:
                pass

        ses2.findById("wnd[1]/tbar[0]/btn[8]").press()
        _esperar_ocupado(ses2, T_MEDIO)

        _log("F8 Ejecutar...")
        ses2.findById("wnd[0]/tbar[1]/btn[8]").press()
        _esperar_con_polling(ses2, max_seg=300, intervalo=2.0)

        try:
            ses2.findById("wnd[0]/tbar[1]/btn[20]").press()
            _esperar_ocupado(ses2, T_MEDIO)
        except Exception as e20:
            _log(f"[WARN] btn[20]: {e20}")

        msg_final, tipo_final = _leer_sbar(ses2)
        _log(f"Resultado: \'{msg_final}\' tipo={tipo_final}")

        if tipo_final == "E":
            return {"ok": False, "error": f"Error ZINPG0004: {msg_final}", "detalles": detalles}

        return {"ok": True, "mensaje": msg_final or "ZINPG0004 OK", "error": "", "detalles": detalles}

    except Exception as e:
        _log(f"Excepcion: {e}")
        return {"ok": False, "error": str(e), "detalles": detalles}

    finally:
        try:
            ses2.findById("wnd[0]/tbar[0]/okcd").text = "/i"
            ses2.findById("wnd[0]").sendVKey(0)
        except Exception:
            try:
                ses2.findById("wnd[0]").close()
            except Exception:
                pass
        _log("Sesion auxiliar ZINPG0004 cerrada.")
'''

# Replace from c223_mantenimiento to end of file
new_content = content[:start_c223] + new_fns
open('sap_mantenimiento.py', 'w', encoding='utf-8').write(new_content)
print('OK')
