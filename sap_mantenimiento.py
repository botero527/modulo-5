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
            _MENSAJES_OK = ("procesados", "pendientes", "completado", "grabado", "correcto", "exitoso")
            if tipo5 == "E" and not any(w in sbar5.lower() for w in _MENSAJES_OK):
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
    """
    C223 sesion auxiliar — flujo VBS correcto:
    Por cada ZFER:
      1. WERKS=CO01, MATNR=zfer, Enter
      2. Si error SAP (material no existe) -> skip, reportar, continuar
      3. ADATU[10,0] = hoy, Enter -> popup btnSPOP-OPTION1
      4. PLNNR[16,0] = hr_id, Enter
    Al final: MAAL -> PRUEFEN -> guardar btn[11]
    """
    import datetime as _dt2
    ses0, ses2, conn, err = _conectar_sap_nueva_sesion()
    if ses2 is None:
        return {"ok": False, "error": f"No se pudo crear sesion auxiliar SAP: {err}", "detalles": []}

    detalles = []
    errores_zfer = []

    def _log(msg):
        print(f"[C223-MANT] {msg}")
        detalles.append(msg)

    hoy    = _dt2.datetime.today().strftime("%d.%m.%Y")
    _TBL   = "wnd[0]/usr/ssubSUBSCR_1200:SAPLCMFV:1200/tblSAPLCMFVT_MKAL"
    _MATNR = "wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-MATNR"
    _WERKS = "wnd[0]/usr/subSUBSCR_1100:SAPLCMFV:1100/ctxtMKAL-WERKS"

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
                _log(f"[WAIT] {_n+1}: {e_nav}")
                time.sleep(1.0)
        if not nav_ok:
            return {"ok": False, "error": "No se pudo navegar a C223", "detalles": detalles}

        for _ in range(3):
            try: ses2.findById("wnd[1]").sendVKey(0); time.sleep(0.3)
            except Exception: break

        # WERKS una sola vez
        ses2.findById(_WERKS).text = "CO01"

        # Procesar cada ZFER
        for i, zfer in enumerate(zfers):
            _log(f"[{i+1}/{len(zfers)}] Procesando {zfer}...")

            # Entrar MATNR + Enter
            try:
                ses2.findById(_MATNR).text = zfer.strip()
                ses2.findById(_MATNR).setFocus()
                ses2.findById(_MATNR).caretPosition = len(zfer.strip())
                ses2.findById("wnd[0]").sendVKey(0)
                _esperar_ocupado(ses2, 2.0)
                # Popup "¿Crear nueva versión?" → OPTION1 = Sí
                try:
                    ses2.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                    time.sleep(0.5)
                    _log(f"  Popup post-MATNR OPTION1 confirmado")
                except Exception:
                    pass
            except Exception as e_matnr:
                _log(f"  [ERROR] MATNR: {e_matnr}")
                errores_zfer.append(f"{zfer}: error al entrar MATNR")
                continue

            # Verificar si SAP mostro error (material no existe)
            sbar_txt, sbar_tipo = _leer_sbar(ses2)
            if sbar_tipo == "E" or "no existe" in sbar_txt.lower() or "no está activado" in sbar_txt.lower():
                _log(f"  [SKIP] SAP error para {zfer}: {sbar_txt}")
                errores_zfer.append(f"{zfer}: {sbar_txt}")
                # Limpiar campo para el siguiente
                try:
                    ses2.findById(_MATNR).text = ""
                except Exception: pass
                continue

            # ADATU[10,0] = fecha hoy, Enter
            try:
                adatu = ses2.findById(f"{_TBL}/ctxtMKAL_EXPAND-ADATU[10,0]")
                adatu.text = hoy
                adatu.setFocus()
                adatu.caretPosition = 2
                ses2.findById("wnd[0]").sendVKey(0)
                _esperar_ocupado(ses2, 1.5)
                _log(f"  ADATU={hoy} OK")
            except Exception as e_adatu:
                _log(f"  [WARN] ADATU: {e_adatu}")

            # Popup confirm (btnSPOP-OPTION1) si aparece
            try:
                ses2.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                time.sleep(0.5)
                _log(f"  Popup SPOP confirmado")
            except Exception:
                pass  # No siempre aparece

            # PLNNR[16,0] = HR, Enter
            try:
                plnnr = ses2.findById(f"{_TBL}/ctxtMKAL_EXPAND-PLNNR[16,0]")
                plnnr.text = hr_id
                plnnr.setFocus()
                plnnr.caretPosition = len(hr_id)
                ses2.findById("wnd[0]").sendVKey(0)
                _esperar_ocupado(ses2, 1.5)
                _log(f"  PLNNR={hr_id} OK")
            except Exception as e_plnnr:
                _log(f"  [WARN] PLNNR: {e_plnnr}")

            # Popup post-PLNNR: "¿Guardar cambios de hoja de ruta?" → OPTION1=Sí
            for _ in range(3):
                try:
                    wnd1 = ses2.findById("wnd[1]")
                    txt_p = ""
                    try: txt_p = str(wnd1.text or "").strip()
                    except Exception: pass
                    if txt_p:
                        _log(f"  Popup post-PLNNR: '{txt_p}'")
                    try:
                        ses2.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                    except Exception:
                        wnd1.sendVKey(0)
                    time.sleep(0.4)
                except Exception:
                    break

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

        try:
            ses2.findById("wnd[0]/tbar[0]/btn[3]").press()
            _esperar_ocupado(ses2, T_RAPIDO)
        except Exception: pass

        sbar_txt, sbar_tipo = _leer_sbar(ses2)
        _log(f"Pre-guardado: '{sbar_txt}' tipo={sbar_tipo}")
        if sbar_tipo == "E":
            return {"ok": False, "error": f"Error C223: {sbar_txt}",
                    "errores_zfer": errores_zfer, "detalles": detalles}

        _log("Guardando btn[11]...")
        ses2.findById("wnd[0]/tbar[0]/btn[11]").press()
        _esperar_ocupado(ses2, T_LENTO)
        
        # Manejar popup de confirmacion de guardado si aparece
        for _ in range(3):
            try:
                wnd1 = ses2.findById("wnd[1]")
                txt_pop = ""
                try: txt_pop = str(wnd1.text or "").strip()
                except Exception: pass
                _log(f"Popup post-guardado: '{txt_pop}'")
                # Intentar btnSPOP-OPTION1 primero (Si/confirmar), si no Enter
                try:
                    ses2.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                except Exception:
                    wnd1.sendVKey(0)
                time.sleep(0.5)
            except Exception:
                break  # no hay popup

        msg_final, tipo_final = _leer_sbar(ses2)
        _log(f"Guardado: '{msg_final}' tipo={tipo_final}")

        if tipo_final == "E":
            return {"ok": False, "error": f"Error guardando: {msg_final}",
                    "errores_zfer": errores_zfer, "detalles": detalles}

        msg_ok = msg_final or "C223 OK"
        if errores_zfer:
            msg_ok += f" | {len(errores_zfer)} ZFERs omitidos: " + "; ".join(errores_zfer[:3])
        return {"ok": True, "mensaje": msg_ok, "error": "",
                "errores_zfer": errores_zfer, "detalles": detalles}

    except Exception as e:
        _log(f"Excepcion: {e}")
        return {"ok": False, "error": str(e), "errores_zfer": errores_zfer, "detalles": detalles}

    finally:
        try:
            ses2.findById("wnd[0]/tbar[0]/okcd").text = "/i"
            ses2.findById("wnd[0]").sendVKey(0)
        except Exception:
            try: ses2.findById("wnd[0]").close()
            except Exception: pass
        _log("Sesion auxiliar C223 cerrada.")


def zinpg0004_actualizar(zfers: list = None) -> dict:
    """
    ZINGP0004 sesion auxiliar — flujo VBS:
    WERKS=CO01, VERID=5000
    Abre multi-value (btn%_SO_MATNR_%_APP_%-VALU_PUSH)
    Llena ZFERs en tabla tabpSIVA (ctxtRSCSEL_255-SLOW_I[1,vis_row])
    btn[0] -> btn[8] (confirmar) -> F8 -> btn[20] (actualizar)
    """
    detalles = []
    def _log(msg):
        print(f"[ZINGP0004] {msg}")
        detalles.append(msg)

    _log(f"Iniciando | {len(zfers) if zfers else 'todos'} ZFERs")
    ses0, ses2, conn, err = _conectar_sap_nueva_sesion()
    if ses2 is None:
        _log(f"ERROR sesion auxiliar: {err}")
        return {"ok": False, "error": f"No se pudo crear sesion auxiliar: {err}", "detalles": detalles}

    _TBL_MV = ("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
               "ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE")

    _log(f"ZINGP0004: {len(zfers) if zfers else 'todos'} ZFERs | CO01 / 5000")

    try:
        # Navegar a ZINGP0004
        nav_ok = False
        for _n in range(15):
            try:
                try: ses2.findById("wnd[1]").sendVKey(0)
                except Exception: pass
                ses2.findById("wnd[0]/tbar[0]/okcd").text = "ZINGP0004"
                ses2.findById("wnd[0]").sendVKey(0)
                _esperar_ocupado(ses2, T_MEDIO)
                nav_ok = True
                _log(f"Nav OK (intento {_n+1})")
                break
            except Exception as e_nav:
                _log(f"[WAIT] {_n+1}: {e_nav}")
                time.sleep(1.0)
        if not nav_ok:
            return {"ok": False, "error": "No se pudo navegar a ZINGP0004", "detalles": detalles}

        for _ in range(3):
            try: ses2.findById("wnd[1]").sendVKey(0); time.sleep(0.3)
            except Exception: break

        # Campos fijos — retry con espera porque ZINGP0004 puede tardar en cargar
        _log(">> Esperando que ZINGP0004 cargue...")
        werks_ok = False
        for _w in range(20):
            try:
                ses2.findById("wnd[0]/usr/ctxtPA_WERKS").text = "CO01"
                werks_ok = True
                _log(f">> ctxtPA_WERKS OK (intento {_w+1})")
                break
            except Exception as e_w:
                try: ses2.findById("wnd[1]").sendVKey(0)
                except Exception: pass
                time.sleep(0.8)
        if not werks_ok:
            return {"ok": False, "error": "ctxtPA_WERKS no accesible en ZINGP0004", "detalles": detalles}
        _log(">> ctxtPA_VERID...")
        ses2.findById("wnd[0]/usr/ctxtPA_VERID").text = "5000"
        _log(">> ctxtSO_MATNR-HIGH focus...")
        ses2.findById("wnd[0]/usr/ctxtSO_MATNR-HIGH").setFocus()
        ses2.findById("wnd[0]/usr/ctxtSO_MATNR-HIGH").caretPosition = 0
        _log(">> Abriendo multi-value...")

        # Abrir multi-value
        ses2.findById("wnd[0]/usr/btn%_SO_MATNR_%_APP_%-VALU_PUSH").press()
        _esperar_ocupado(ses2, T_MEDIO)

        if zfers:
            # Leer filas visibles de la tabla
            try:
                vis_mv = ses2.findById(_TBL_MV).VisibleRowCount
            except Exception:
                vis_mv = 8
            _log(f"Tabla multi-value: vis={vis_mv}, ZFERs={len(zfers)}")

            # Llenar ZFERs uno a uno con scroll
            for i, zfer in enumerate(zfers):
                vis_row = i % vis_mv
                # Scroll cuando se llena la pagina visible
                if i > 0 and vis_row == 0:
                    try:
                        ses2.findById(_TBL_MV).verticalScrollbar.position = i
                        time.sleep(0.1)
                    except Exception as e_sc:
                        _log(f"[WARN] scroll {i}: {e_sc}")
                try:
                    ses2.findById(f"{_TBL_MV}/ctxtRSCSEL_255-SLOW_I[1,{vis_row}]").text = zfer.strip()
                except Exception as e_cell:
                    _log(f"[WARN] celda {i} ({zfer}): {e_cell}")

            _log(f"{len(zfers)} ZFERs escritos")
        else:
            # Sin filtro = todos (limpiar con btn[24] si existe)
            _log("Sin filtro de material — ejecutando para todos")
            try:
                ses2.findById("wnd[1]/tbar[0]/btn[24]").press()
                _esperar_ocupado(ses2, T_RAPIDO)
            except Exception:
                pass

        # btn[0] (confirmar seleccion verde) → btn[8] (ejecutar popup)
        _log(">> btn[0] confirmar seleccion...")
        ses2.findById("wnd[1]/tbar[0]/btn[0]").press()
        _esperar_ocupado(ses2, T_RAPIDO)

        _log(">> btn[8] ejecutar seleccion...")
        ses2.findById("wnd[1]/tbar[0]/btn[8]").press()
        _esperar_ocupado(ses2, T_MEDIO)

        # Verificar que wnd[1] ya se cerro
        try:
            ses2.findById("wnd[1]")
            # Si sigue abierto, cerrarlo
            ses2.findById("wnd[1]").sendVKey(0)
            _esperar_ocupado(ses2, T_RAPIDO)
        except Exception:
            pass  # ya cerro = correcto

        # F8 Ejecutar (en pantalla principal)
        _log(">> F8 Ejecutar...")
        ses2.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Esperar procesamiento (puede tardar)
        resultado = _esperar_con_polling(ses2, max_seg=300, intervalo=2.0)
        _log(f"Post-F8: {resultado}")

        if resultado.startswith("POPUP_ERROR"):
            msg_err = resultado.replace("POPUP_ERROR: ", "")
            try: ses2.findById("wnd[1]").sendVKey(12)
            except Exception: pass
            return {"ok": False, "error": f"Error ZINGP0004: {msg_err}", "detalles": detalles}

        # btn[20] Actualizar
        _log(">> btn[20] Actualizar...")
        try:
            ses2.findById("wnd[0]/tbar[1]/btn[20]").press()
            _esperar_ocupado(ses2, T_MEDIO)
        except Exception as e20:
            _log(f"[WARN] btn[20]: {e20}")

        msg_final, tipo_final = _leer_sbar(ses2)
        _log(f"Resultado: '{msg_final}' tipo={tipo_final}")

        if tipo_final == "E":
            return {"ok": False, "error": f"Error ZINGP0004: {msg_final}", "detalles": detalles}

        return {"ok": True, "mensaje": msg_final or "ZINGP0004 OK", "error": "", "detalles": detalles}

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
        _log("Sesion auxiliar ZINGP0004 cerrada.")

