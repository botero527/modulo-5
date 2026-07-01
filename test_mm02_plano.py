"""
test_mm02_plano.py — Prueba aislada de lectura/escritura de plano en MM02 tab ZU04.

Flujo completo (igual al flujo real en sap_auto.py):
  1. Lee DOKNR del ZFER_BASE desde SAP (tab ZU04)
     Fallback: si SAP falla → busca en BD por MATERIAL (sin SP o con SP según modo)
  2. Busca en BD (ODATA_ZFER_RUTAS_JPG) el plano nuevo:
     - modo SIN SP (--sinsp, default):  busca plano sin " SP" al final
     - modo CON SP (--consp):            busca plano con " SP" al final
  3. Si encuentra → escribe en ZFER_NUEVO y guarda

Uso:
    py test_mm02_plano.py 700170478 700182985          (sin SP, ambos ZFERs)
    py test_mm02_plano.py 700170478 700182985 --consp  (con SP)
    py test_mm02_plano.py 700170478                    (solo leer, no escribe)
"""

import sys, time, re
import win32com.client
import pyodbc

# ── Args ──────────────────────────────────────────────────────────────────────
args = sys.argv[1:]
ZFER_BASE    = args[0] if len(args) > 0 else "700170478"
ZFER_NUEVO   = args[1] if len(args) > 1 and not args[1].startswith("--") else None
CON_SP       = "--consp" in args

print("=" * 60)
print("TEST MM02 PLANO")
print(f"  ZFER base  : {ZFER_BASE}")
print(f"  ZFER nuevo : {ZFER_NUEVO or '(solo lectura)'}")
print(f"  Modo       : {'CON SP' if CON_SP else 'SIN SP'}")
print("=" * 60)

# ── Timings ───────────────────────────────────────────────────────────────────
T_RAPIDO = 0.5
T_MEDIO  = 1.5

def esperar(t):
    time.sleep(t)

# ── BD connection string (igual que sap_auto.py) ──────────────────────────────
_DB_SAP_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=agpcolsap.database.windows.net;"
    "DATABASE=DB_COL_SAP;"
    "UID=Viewer;"
    "PWD=AgpconsCol2023;"
    "Encrypt=yes;TrustServerCertificate=no;Connection Timeout=20;"
)

# ── Helpers popup ─────────────────────────────────────────────────────────────
def popup_texto(session):
    try:    return session.findById("wnd[1]").text or "(sin título)"
    except: return None

def confirmar_popup(session, paso=""):
    txt = popup_texto(session)
    if txt is None:
        return False
    print(f"  [POPUP-{paso}] '{txt}' → confirmando Sí/OK")
    try:
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        print(f"    → btnSPOP-OPTION1 (Sí)")
    except Exception:
        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            print(f"    → tbar[0]/btn[0]")
        except Exception:
            session.findById("wnd[1]").sendVKey(0)
            print(f"    → sendVKey(0)")
    esperar(T_RAPIDO)
    return True

def confirmar_todos(session, paso="", max_n=4):
    n = 0
    for _ in range(max_n):
        if confirmar_popup(session, paso):
            n += 1
        else:
            break
    if n == 0:
        print(f"  [OK-{paso}] sin popup")
    return n

# ── Helpers BD ────────────────────────────────────────────────────────────────
def _plano_base(doknr):
    """Quita versión y SP del final: 'M1344 000 001 A SP' → 'M1344 000 001'"""
    return re.sub(r'(\s+[A-Za-z]{1,3})+$', '', doknr.strip()).strip()

def buscar_plano_bd(doknr_leido, con_sp):
    """
    Busca en ODATA_ZFER_RUTAS_JPG el plano nuevo.
    con_sp=False → sin SP  (flujo con→sin acero)
    con_sp=True  → con SP  (flujo sin→con acero)
    """
    base = _plano_base(doknr_leido)
    print(f"  [BD] base de búsqueda: '{base}' (de DOKNR '{doknr_leido}')")
    if not base:
        return None, "DOKNR sin base reconocible"
    try:
        cn  = pyodbc.connect(_DB_SAP_STR, autocommit=True)
        cur = cn.cursor()
        if con_sp:
            cur.execute(
                "SELECT TOP 1 DOCUMENTO, PLANO, ULTIMA_MOD "
                "FROM dbo.ODATA_ZFER_RUTAS_JPG "
                "WHERE DOCUMENTO LIKE ? AND DOCUMENTO LIKE '% SP' "
                "ORDER BY ULTIMA_MOD DESC",
                f"%{base}%"
            )
        else:
            cur.execute(
                "SELECT TOP 1 DOCUMENTO, PLANO, ULTIMA_MOD "
                "FROM dbo.ODATA_ZFER_RUTAS_JPG "
                "WHERE DOCUMENTO LIKE ? "
                "  AND DOCUMENTO NOT LIKE '% SP' "
                "  AND DOCUMENTO NOT LIKE '% L[0-9]%' "
                "ORDER BY ULTIMA_MOD DESC",
                f"%{base}%"
            )
        row = cur.fetchone()
        cn.close()
        if not row or not row[0]:
            tipo = "con SP" if con_sp else "sin SP"
            return None, f"⚠ No se encontró plano {tipo} para base '{base}'"
        return str(row[0]).strip(), f"DOCUMENTO='{row[0].strip()}' | PLANO='{row[1]}' | MOD={row[2]}"
    except Exception as e:
        return None, f"Error BD: {e}"

def buscar_doknr_por_material_bd(zfer, con_sp):
    """Fallback: busca el DOKNR directamente por MATERIAL en BD."""
    try:
        cn  = pyodbc.connect(_DB_SAP_STR, autocommit=True)
        cur = cn.cursor()
        if con_sp:
            cur.execute(
                "SELECT TOP 1 DOCUMENTO FROM dbo.ODATA_ZFER_RUTAS_JPG "
                "WHERE MATERIAL=? AND CENTRO='CO01' AND DOCUMENTO LIKE '% SP' "
                "ORDER BY ULTIMA_MOD DESC", zfer
            )
        else:
            cur.execute(
                "SELECT TOP 1 DOCUMENTO FROM dbo.ODATA_ZFER_RUTAS_JPG "
                "WHERE MATERIAL=? AND CENTRO='CO01' AND DOCUMENTO NOT LIKE '% SP' "
                "ORDER BY ULTIMA_MOD DESC", zfer
            )
        row = cur.fetchone()
        cn.close()
        if not row or not row[0]:
            return None, f"Sin plano en BD para MATERIAL={zfer}"
        return str(row[0]).strip(), f"DOKNR por material BD: '{row[0].strip()}'"
    except Exception as e:
        return None, f"Error BD: {e}"

# ── IDs SAP ───────────────────────────────────────────────────────────────────
_ID_TCODE  = "wnd[0]/tbar[0]/okcd"
_ID_MATNR  = "wnd[0]/usr/ctxtRMMG1-MATNR"
_subZU04   = ("wnd[0]/usr/tabsTABSPR1/tabpZU04"
              "/ssubTABFRA1:SAPLMGMM:2110"
              "/subSUB2:SAPLMGD1:3400"
              "/subDOCU:SAPLCV140:0204")
_grid_docu = _subZU04 + "/subDOC_ALV:SAPLCV140:0206/cntlALV_CUST_DOC/shellcont/shell"

def navegar_mm02_zu04(session, zfer, label=""):
    """
    Navega a MM02, entra el material, va a tab ZU04.
    Retorna True si llegó al grid de documentos, False si falló.
    """
    print(f"\n{'─'*50}")
    print(f"NAVEGANDO MM02 ZU04 — {label} ({zfer})")
    print(f"{'─'*50}")

    # Paso A: /NMM02
    print("A) /NMM02 → sendVKey(0)")
    session.findById(_ID_TCODE).text = "/NMM02"
    session.findById("wnd[0]").sendVKey(0)
    esperar(T_MEDIO)
    confirmar_todos(session, "A_NMM02")

    # Paso B: entrar material
    print(f"B) material={zfer} → sendVKey(0)")
    session.findById(_ID_MATNR).text = zfer
    session.findById(_ID_MATNR).caretPosition = len(zfer)
    session.findById("wnd[0]").sendVKey(0)
    esperar(T_MEDIO)
    confirmar_todos(session, "B_material")

    # Paso C: btn[30] Datos adicionales
    print("C) btn[30] (Datos adicionales)")
    try:
        session.findById("wnd[0]/tbar[1]/btn[30]").press()
        esperar(T_MEDIO)
        confirmar_todos(session, "C_btn30")
    except Exception as e:
        print(f"  [ERROR C] btn[30]: {e}")
        return False

    # Paso D: tab ZU04
    print("D) tabpZU04.select")
    try:
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU04").select()
        esperar(T_MEDIO)
        confirmar_todos(session, "D_tabZU04")
    except Exception as e:
        print(f"  [ERROR D] tabpZU04: {e}")
        return False

    # Paso E: radio GF_ALLE
    print("E) radGF_ALLE (opcional)")
    try:
        session.findById(_subZU04 + "/subBUTTON:SAPLCV140:0203/radGF_ALLE").setFocus()
        session.findById(_subZU04 + "/subBUTTON:SAPLCV140:0203/radGF_ALLE").select()
        esperar(T_RAPIDO)
        print("  [OK] radGF_ALLE")
    except Exception as e:
        print(f"  [WARN] radGF_ALLE no disponible: {e}")

    return True

# ── CONECTAR SAP ──────────────────────────────────────────────────────────────
try:
    sap_gui = win32com.client.GetObject("SAPGUI")
    app     = sap_gui.GetScriptingEngine
    conn    = app.Children(0)
    session = conn.Children(0)
    print("[OK] SAP conectado\n")
except Exception as e:
    print(f"[ERROR] No se pudo conectar a SAP: {e}")
    sys.exit(1)

# ═══════════════════════════════════════════════════════
# BLOQUE 1 — Leer DOKNR del ZFER BASE
# ═══════════════════════════════════════════════════════
doknr_base = ""

ok_sap = navegar_mm02_zu04(session, ZFER_BASE, "LECTURA BASE")
if ok_sap:
    print("F) Leer DOKNR del grid")
    try:
        doknr_base = str(session.findById(_grid_docu).getCellValue(0, "DOKNR") or "").strip()
        print(f"  [OK] DOKNR SAP = '{doknr_base}'")
    except Exception as e:
        print(f"  [ERROR] No pudo leer grid: {e}")
        txt_popup = popup_texto(session)
        if txt_popup:
            print(f"  [INFO] popup activo: '{txt_popup}'")

# Fallback BD si SAP no dio nada
if not doknr_base:
    print("\n[FALLBACK] Buscando DOKNR en BD por MATERIAL...")
    doknr_base, msg_fb = buscar_doknr_por_material_bd(ZFER_BASE, CON_SP)
    print(f"  {msg_fb}")

print(f"\n>>> DOKNR BASE final: '{doknr_base}'")

# ═══════════════════════════════════════════════════════
# BLOQUE 2 — Buscar plano nuevo en BD
# ═══════════════════════════════════════════════════════
plano_nuevo = None
if doknr_base:
    print("\n" + "═"*50)
    print("BÚSQUEDA EN BD")
    print("═"*50)
    plano_nuevo, msg_bd = buscar_plano_bd(doknr_base, CON_SP)
    if plano_nuevo:
        print(f"  [OK] {msg_bd}")
    else:
        print(f"  [WARN] {msg_bd}")
else:
    print("\n[SKIP] Sin DOKNR base — no se busca en BD")

print(f"\n>>> PLANO NUEVO: '{plano_nuevo}'")

# ═══════════════════════════════════════════════════════
# BLOQUE 3 — Escribir plano en ZFER NUEVO
# ═══════════════════════════════════════════════════════
if ZFER_NUEVO and plano_nuevo:
    print("\n" + "═"*50)
    print(f"ESCRITURA en {ZFER_NUEVO}")
    print("═"*50)

    ok_sap2 = navegar_mm02_zu04(session, ZFER_NUEVO, "ESCRITURA NUEVO")
    if ok_sap2:
        print(f"F) Escribir DOKNR='{plano_nuevo}' en grid")
        try:
            session.findById(_grid_docu).modifyCell(0, "DOKNR", plano_nuevo)
            session.findById(_grid_docu).currentCellColumn = "DOKNR"
            esperar(T_RAPIDO)
            # Guardar Ctrl+S (btn[11])
            print("G) Guardar (Ctrl+S / btn[11])")
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            esperar(T_MEDIO)
            confirmar_todos(session, "guardar")
            print(f"  [OK] Guardado")
        except Exception as e:
            print(f"  [ERROR] escritura: {e}")
            txt_popup = popup_texto(session)
            if txt_popup:
                print(f"  [INFO] popup activo: '{txt_popup}'")
elif ZFER_NUEVO and not plano_nuevo:
    print(f"\n[SKIP] No se escribe en {ZFER_NUEVO} — sin plano nuevo en BD")
else:
    print("\n[SKIP] No se especificó ZFER_NUEVO")

# ═══════════════════════════════════════════════════════
print("\n" + "="*60)
print("RESUMEN")
print(f"  DOKNR leído de SAP/BD : '{doknr_base}'")
print(f"  Plano nuevo (BD)      : '{plano_nuevo}'")
print(f"  Escrito en            : '{ZFER_NUEVO or 'N/A'}'")
print("="*60)
