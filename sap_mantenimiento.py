"""
sap_mantenimiento.py — Automatización SAP para Mantenimiento de Hojas de Ruta
Proceso independiente del homologador (sap_auto.py).

Transacciones:
  ZPPP0084 — Desasignar / Asignar materiales a HR masivamente vía Excel
"""

import time
import win32com.client

T_RAPIDO = 0.5
T_MEDIO  = 1.5
T_LENTO  = 8.0   # más generoso que el homologador — procesos masivos tardan más


def _conectar_sap():
    """Retorna (session, error). session=None si falla."""
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        app     = sap_gui.GetScriptingEngine
        conn    = app.Children(0)
        ses     = conn.Children(0)
        return ses, None
    except Exception as e:
        return None, str(e)


def _esperar_ocupado(ses, max_seg: float = T_LENTO):
    """Espera a que SAP deje de estar ocupado, con tope máximo."""
    t0 = time.time()
    time.sleep(0.1)
    while time.time() - t0 < max_seg:
        try:
            if not ses.Busy:
                break
        except Exception:
            break
        time.sleep(0.2)


def _leer_sbar(ses) -> tuple[str, str]:
    """Retorna (texto, tipo) del status bar. tipo: 'E'=error, 'W'=warning, 'S'=success, ''=neutro."""
    try:
        sbar = ses.findById("wnd[0]/sbar")
        texto = str(sbar.text or "").strip()
        tipo  = str(sbar.messageType or "").strip().upper()
        return texto, tipo
    except Exception:
        return "", ""


def _leer_popup(ses) -> tuple[str, str]:
    """
    Lee wnd[1] si existe. Retorna (texto, tipo_icono).
    tipo_icono: 'E'=error, 'W'=warning, 'S'=success, 'I'=info, ''=no hay popup.
    """
    try:
        wnd1 = ses.findById("wnd[1]")
        # Intentar leer texto principal del popup
        textos = []
        for ctrl_id in ("usr/txtMessageTxt", "usr/lblMessageTxt",
                        "usr/txtSY-MSGV1",  "usr/lblG_TITLE"):
            try:
                t = str(ses.findById(f"wnd[1]/{ctrl_id}").text or "").strip()
                if t:
                    textos.append(t)
            except Exception:
                pass
        # Si no encontró texto por IDs conocidos, leer título de la ventana
        titulo = ""
        try:
            titulo = str(wnd1.text or "").strip()
        except Exception:
            pass
        texto_final = " | ".join(textos) if textos else titulo

        # Tipo de icono
        tipo = ""
        try:
            tipo = str(ses.findById("wnd[1]/usr/radMsgTyp").text or "").upper()
        except Exception:
            pass

        return texto_final, tipo
    except Exception:
        return "", ""   # no hay wnd[1]


def _hay_popup_error(ses) -> tuple[bool, str]:
    """True si hay popup con mensaje de error. Retorna (es_error, texto)."""
    texto, tipo = _leer_popup(ses)
    if not texto:
        return False, ""
    es_error = tipo in ("E", "A") or any(
        w in texto.lower() for w in ("error", "incorrecto", "no se puede", "fallo", "aborted")
    )
    return es_error, texto


def _cerrar_popup(ses):
    """Cierra wnd[1] con Enter."""
    try:
        ses.findById("wnd[1]").sendVKey(0)
        time.sleep(T_RAPIDO)
    except Exception:
        pass


def _esperar_con_polling(ses, max_seg: float, intervalo: float = 1.0) -> str:
    """
    Espera hasta max_seg mientras SAP procesa.
    Retorna el mensaje de sbar cuando termine o 'TIMEOUT' si se agotó el tiempo.
    Detecta popups de error durante la espera.
    """
    t0 = time.time()
    while time.time() - t0 < max_seg:
        time.sleep(intervalo)
        # Verificar si SAP sigue ocupado
        try:
            ocupado = ses.Busy
        except Exception:
            ocupado = False

        # Revisar popup durante procesamiento
        es_err, txt_err = _hay_popup_error(ses)
        if es_err:
            return f"POPUP_ERROR: {txt_err}"

        if not ocupado:
            sbar_txt, sbar_tipo = _leer_sbar(ses)
            return f"{sbar_tipo}:{sbar_txt}" if sbar_txt else "DONE"

    return "TIMEOUT"


def zppp0084_desasignar(excel_path: str) -> dict:
    """
    Ejecuta ZPPP0084 modo Desasignar con el Excel indicado.
    Espera el procesamiento, detecta errores de SAP (popups, sbar, ventanas) y los reporta.
    Retorna {ok, mensaje, error, detalles:[]}
    """
    ses, err = _conectar_sap()
    if ses is None:
        return {"ok": False, "error": f"SAP GUI no disponible: {err}", "mensaje": "", "detalles": []}

    detalles = []

    def _log(msg):
        print(f"[ZPPP0084] {msg}")
        detalles.append(msg)

    try:
        _log(f"Iniciando desasignación | archivo: {excel_path}")

        # ── Navegar a ZPPP0084 ────────────────────────────────────────────────
        ses.findById("wnd[0]/tbar[0]/okcd").text = "/nZPPP0084"
        ses.findById("wnd[0]").sendVKey(0)
        _esperar_ocupado(ses, T_MEDIO)

        # Cerrar popups post-navegación
        for _ in range(3):
            texto_pop, _ = _leer_popup(ses)
            if not texto_pop:
                break
            _log(f"Popup post-nav: '{texto_pop}' — cerrando")
            _cerrar_popup(ses)

        # Verificar que llegamos a ZPPP0084
        sbar_txt, _ = _leer_sbar(ses)
        _log(f"Pantalla cargada | sbar: '{sbar_txt}'")
        if "autorización" in sbar_txt.lower() or "autorizacion" in sbar_txt.lower() or "not authorized" in sbar_txt.lower():
            return {"ok": False,
                    "error": f"Sin autorización para ZPPP0084 — pide acceso al admin SAP. ({sbar_txt})",
                    "mensaje": sbar_txt, "detalles": detalles}

        # ── Configurar campos ─────────────────────────────────────────────────
        ses.findById("wnd[0]/usr/radP_DESI").setFocus()
        ses.findById("wnd[0]/usr/radP_DESI").select()
        time.sleep(T_RAPIDO)

        ses.findById("wnd[0]/usr/ctxtP_WERK").text = "CO01"
        ses.findById("wnd[0]/usr/ctxtP_ARCH").text = excel_path.replace("/", "\\")
        ses.findById("wnd[0]/usr/ctxtP_WERK").caretPosition = 4
        time.sleep(T_RAPIDO)

        # Verificar que el archivo se escribió bien
        ruta_leida = ""
        try:
            ruta_leida = ses.findById("wnd[0]/usr/ctxtP_ARCH").text
        except Exception:
            pass
        _log(f"Archivo configurado: '{ruta_leida}'")

        # ── F8 Ejecutar ───────────────────────────────────────────────────────
        _log("Presionando F8 (Ejecutar)...")
        ses.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Esperar procesamiento — puede tardar bastante con muchos registros
        # Poll cada 2s hasta 5 minutos
        resultado_espera = _esperar_con_polling(ses, max_seg=300, intervalo=2.0)
        _log(f"Post-F8: '{resultado_espera}'")

        if resultado_espera.startswith("POPUP_ERROR"):
            msg_err = resultado_espera.replace("POPUP_ERROR: ", "")
            _cerrar_popup(ses)
            return {"ok": False, "error": f"SAP reportó error: {msg_err}",
                    "mensaje": msg_err, "detalles": detalles}

        if resultado_espera == "TIMEOUT":
            return {"ok": False, "error": "SAP no respondió en 5 minutos — verifica manualmente",
                    "mensaje": "", "detalles": detalles}

        # ── btn[5] ────────────────────────────────────────────────────────────
        _log("Presionando btn[5]...")
        try:
            ses.findById("wnd[0]/tbar[1]/btn[5]").press()
            _esperar_ocupado(ses, T_MEDIO)
            sbar5, tipo5 = _leer_sbar(ses)
            _log(f"Post-btn5 sbar: '{sbar5}' tipo={tipo5}")
            # SAP a veces reporta mensajes de éxito con tipo "E" (ej: "Ya fueron procesados todos")
            # Solo fallar si claramente es un error real (no un mensaje de completado)
            _MENSAJES_OK = ("procesados", "pendientes", "completado", "grabado", "correcto", "exitoso")
            if tipo5 == "E" and not any(w in sbar5.lower() for w in _MENSAJES_OK):
                return {"ok": False, "error": f"Error tras btn[5]: {sbar5}",
                        "mensaje": sbar5, "detalles": detalles}
        except Exception as e5:
            _log(f"btn[5] no disponible: {e5}")

        # ── btn[15] x2 ────────────────────────────────────────────────────────
        for i in range(2):
            _log(f"Presionando btn[15] ({i+1}/2)...")
            try:
                ses.findById("wnd[0]/tbar[0]/btn[15]").press()
                _esperar_ocupado(ses, T_MEDIO)
                sbar15, tipo15 = _leer_sbar(ses)
                _log(f"Post-btn15 sbar: '{sbar15}' tipo={tipo15}")

                # Verificar popup después de btn[15]
                es_err, txt_pop = _hay_popup_error(ses)
                if es_err:
                    _log(f"Popup error detectado: '{txt_pop}'")
                    _cerrar_popup(ses)
                    return {"ok": False, "error": f"Error SAP: {txt_pop}",
                            "mensaje": txt_pop, "detalles": detalles}
                elif txt_pop:
                    _log(f"Popup informativo: '{txt_pop}' — cerrando")
                    _cerrar_popup(ses)

                if tipo15 == "E":
                    return {"ok": False, "error": f"Error SAP: {sbar15}",
                            "mensaje": sbar15, "detalles": detalles}
            except Exception as e15:
                _log(f"btn[15] no disponible ({i+1}/2): {e15}")

        # ── Resultado final ───────────────────────────────────────────────────
        mensaje_final, tipo_final = _leer_sbar(ses)
        _log(f"Resultado final: sbar='{mensaje_final}' tipo={tipo_final}")

        if tipo_final == "E":
            return {"ok": False, "error": f"SAP error final: {mensaje_final}",
                    "mensaje": mensaje_final, "detalles": detalles}

        _log("Desasignación completada exitosamente")
        return {"ok": True, "mensaje": mensaje_final or "Proceso completado",
                "error": "", "detalles": detalles}

    except Exception as e:
        _log(f"Excepción inesperada: {e}")
        return {"ok": False, "error": str(e), "mensaje": "", "detalles": detalles}


def zppp0084_asignar(excel_path: str) -> dict:
    """
    Ejecuta ZPPP0084 modo Asignar con el Excel indicado.
    VBS grabado: NO selecciona radio — 'Asignar' es el default al abrir la transacción.
    Secuencia: ZPPP0084 → CO01 → archivo → F8 → btn[5] → btn[15] x2
    """
    ses, err = _conectar_sap()
    if ses is None:
        return {"ok": False, "error": f"SAP GUI no disponible: {err}", "mensaje": "", "detalles": []}

    detalles = []
    def _log(msg):
        print(f"[ZPPP0084-ASIGNAR] {msg}")
        detalles.append(msg)

    try:
        _log(f"Iniciando asignacion | archivo: {excel_path}")

        # Navegar a ZPPP0084 (radio Asignar ya viene seleccionado por defecto)
        ses.findById("wnd[0]/tbar[0]/okcd").text = "ZPPP0084"
        ses.findById("wnd[0]").sendVKey(0)
        _esperar_ocupado(ses, T_MEDIO)

        for _ in range(3):
            texto_pop, _ = _leer_popup(ses)
            if not texto_pop:
                break
            _log(f"Popup post-nav: '{texto_pop}' — cerrando")
            _cerrar_popup(ses)

        sbar_asig, _ = _leer_sbar(ses)
        _log(f"Pantalla asignar | sbar: '{sbar_asig}'")
        if "autorización" in sbar_asig.lower() or "autorizacion" in sbar_asig.lower() or "not authorized" in sbar_asig.lower():
            return {"ok": False,
                    "error": f"Sin autorización para ZPPP0084 — pide acceso al admin SAP. ({sbar_asig})",
                    "mensaje": sbar_asig, "detalles": detalles}

        # Seleccionar radio Asignar explícitamente (por si venía de otra pantalla)
        try:
            ses.findById("wnd[0]/usr/radP_ASIG").setFocus()
            ses.findById("wnd[0]/usr/radP_ASIG").select()
            time.sleep(T_RAPIDO)
        except Exception as e_rad:
            _log(f"[WARN] No se pudo seleccionar radP_ASIG: {e_rad} — continuando con default")

        ses.findById("wnd[0]/usr/ctxtP_WERK").text = "CO01"
        ses.findById("wnd[0]/usr/ctxtP_ARCH").text = excel_path.replace("/", "\\")
        ses.findById("wnd[0]/usr/ctxtP_ARCH").setFocus()
        ses.findById("wnd[0]/usr/ctxtP_ARCH").caretPosition = len(excel_path)
        time.sleep(T_RAPIDO)

        _log("Presionando F8 (Ejecutar)...")
        ses.findById("wnd[0]/tbar[1]/btn[8]").press()

        resultado_espera = _esperar_con_polling(ses, max_seg=300, intervalo=2.0)
        _log(f"Post-F8: '{resultado_espera}'")

        if resultado_espera.startswith("POPUP_ERROR"):
            msg_err = resultado_espera.replace("POPUP_ERROR: ", "")
            _cerrar_popup(ses)
            return {"ok": False, "error": f"SAP error: {msg_err}", "mensaje": msg_err, "detalles": detalles}

        if resultado_espera == "TIMEOUT":
            return {"ok": False, "error": "SAP no respondio en 5 minutos", "mensaje": "", "detalles": detalles}

        try:
            ses.findById("wnd[0]/tbar[1]/btn[5]").press()
            _esperar_ocupado(ses, T_MEDIO)
            sbar5, tipo5 = _leer_sbar(ses)
            _log(f"Post-btn5: '{sbar5}' tipo={tipo5}")
            if tipo5 == "E":
                return {"ok": False, "error": f"Error tras btn[5]: {sbar5}", "mensaje": sbar5, "detalles": detalles}
        except Exception as e5:
            _log(f"btn[5] no disponible: {e5}")

        for i in range(2):
            try:
                ses.findById("wnd[0]/tbar[0]/btn[15]").press()
                _esperar_ocupado(ses, T_MEDIO)
                sbar15, tipo15 = _leer_sbar(ses)
                es_err, txt_pop = _hay_popup_error(ses)
                if es_err:
                    _cerrar_popup(ses)
                    return {"ok": False, "error": f"Error SAP: {txt_pop}", "mensaje": txt_pop, "detalles": detalles}
                elif txt_pop:
                    _cerrar_popup(ses)
                if tipo15 == "E":
                    return {"ok": False, "error": f"Error SAP: {sbar15}", "mensaje": sbar15, "detalles": detalles}
            except Exception as e15:
                _log(f"btn[15] no disponible ({i+1}/2): {e15}")

        mensaje_final, tipo_final = _leer_sbar(ses)
        _log(f"Resultado final: '{mensaje_final}' tipo={tipo_final}")
        if tipo_final == "E":
            return {"ok": False, "error": f"SAP error final: {mensaje_final}", "mensaje": mensaje_final, "detalles": detalles}

        _log("Asignacion completada exitosamente")
        return {"ok": True, "mensaje": mensaje_final or "Proceso completado", "error": "", "detalles": detalles}

    except Exception as e:
        _log(f"Excepcion inesperada: {e}")
        return {"ok": False, "error": str(e), "mensaje": "", "detalles": detalles}



def _conectar_sap_nueva_sesion():
    """
    Abre una sesion auxiliar nueva en SAP.
    Verifica que realmente se creó una sesión nueva (si createSession falla en sesión
    bloqueada, conn.Children.Count no aumenta y devolveríamos la misma sesión bloqueada).
    """
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        app     = sap_gui.GetScriptingEngine
        conn    = app.Children(0)

        # Intentar con cualquier sesión disponible para createSession
        count_antes = conn.Children.Count
        print(f"[SAP-AUX] Sesiones antes: {count_antes}")

        # Probar cada sesión hasta que una pueda crear la nueva
        creada = False
        for idx in range(count_antes):
            try:
                conn.Children(idx).createSession()
                time.sleep(3)
                count_nueva = conn.Children.Count
                if count_nueva > count_antes:
                    creada = True
                    print(f"[SAP-AUX] Sesion creada desde idx={idx} | total={count_nueva}")
                    break
            except Exception as e_cs:
                print(f"[SAP-AUX] createSession idx={idx} fallo: {e_cs}")
                continue

        if not creada:
            return None, None, None, "createSession no creó sesión nueva — todas las sesiones bloqueadas"

        ses_nueva = conn.Children(conn.Children.Count - 1)

        # Esperar que la sesión cargue y cerrar popups de bienvenida
        for _ in range(30):
            time.sleep(0.5)
            # Cerrar cualquier popup (wnd[1]) que bloquee
            try:
                ses_nueva.findById("wnd[1]").sendVKey(0)
            except Exception:
                pass
            # Verificar que se puede ESCRIBIR en okcd (no solo leerlo)
            try:
                ses_nueva.findById("wnd[0]/tbar[0]/okcd").text = ""
                break
            except Exception:
                pass

        ses_nueva.findById("wnd[0]").maximize()
        return conn.Children(0), ses_nueva, conn, None
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
        # Navegar a C223 con reintentos — sesion nueva puede tardar en estabilizarse
        nav_ok = False
        for _n in range(15):
            try:
                # Cerrar popups que puedan bloquear
                try: ses2.findById("wnd[1]").sendVKey(0)
                except Exception: pass
                ses2.findById("wnd[0]/tbar[0]/okcd").text = "C223"
                ses2.findById("wnd[0]").sendVKey(0)
                _esperar_ocupado(ses2, T_MEDIO)
                nav_ok = True
                _log(f"Navegacion C223 OK (intento {_n+1})")
                break
            except Exception as e_nav:
                _log(f"[WAIT] Nav C223 intento {_n+1}: {e_nav}")
                time.sleep(1.0)
        if not nav_ok:
            return {"ok": False, "error": "No se pudo navegar a C223 tras 15 intentos", "detalles": detalles}

        # Cerrar popups post-navegacion
        for _ in range(3):
            try:
                ses2.findById("wnd[1]").sendVKey(0)
                time.sleep(0.3)
            except Exception:
                break

        _log(">> Paso: btnG_ICON_SELECTION_FV")
        ses2.findById("wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/btnG_ICON_SELECTION_FV").press()
        _esperar_ocupado(ses2, T_MEDIO)

        _log(">> Paso: btn RANG_MAT multi-value")
        ses2.findById("wnd[1]/usr/btn%_RANG_MAT_%_APP_%-VALU_PUSH").press()
        _esperar_ocupado(ses2, T_MEDIO)

        _log(">> Paso: btn[24] limpiar")
        try:
            ses2.findById("wnd[2]/tbar[0]/btn[24]").press()
            _esperar_ocupado(ses2, T_RAPIDO)
        except Exception as e24:
            _log(f"[INFO] btn24: {e24}")

        try:
            vis_mv = ses2.findById(_TBL_MV).VisibleRowCount
        except Exception:
            vis_mv = 20

        _log(f">> Paso: entrar ZFERs en multi-value (vis={vis_mv})")
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
                _log(f"  ZFER [{i}]: {zfer} OK")
            except Exception as e_mv:
                _log(f"[WARN] mv fila {i} ({zfer}): {e_mv}")

        _log(">> Paso: confirmar wnd[2] btn[8]")
        ses2.findById("wnd[2]/tbar[0]/btn[8]").press()
        _esperar_ocupado(ses2, T_MEDIO)

        _log(">> Paso: confirmar wnd[1] btn[8]")
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
        _log(f"Pre-guardado: '{sbar_txt}' tipo={sbar_tipo}")
        if sbar_tipo == "E":
            return {"ok": False, "error": f"Error C223: {sbar_txt}", "detalles": detalles}

        # Guardar btn[11]
        _log("Guardando...")
        ses2.findById("wnd[0]/tbar[0]/btn[11]").press()
        _esperar_ocupado(ses2, T_LENTO)

        msg_final, tipo_final = _leer_sbar(ses2)
        _log(f"Guardado: '{msg_final}' tipo={tipo_final}")

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
        _log(f"Resultado: '{msg_final}' tipo={tipo_final}")

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
