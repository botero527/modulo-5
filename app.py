"""
app.py — MODULO 5 AGP Glass
Sistema de consulta y reporte de ZFERs (Colombia CO01)
"""
from flask import Flask, render_template, request, redirect, url_for
import pyodbc

app = Flask(__name__)

# ── Configuración BD ──────────────────────────────────────────────────────────
DB_SAP = {
    "server":   "agpcolsap.database.windows.net",
    "database": "DB_COL_SAP",
    "driver":   "ODBC Driver 17 for SQL Server",
    "user":     "Viewer",
    "password": "AgpconsCol2023",
}

def _conn_str():
    return (
        f"DRIVER={{{DB_SAP['driver']}}};"
        f"SERVER={DB_SAP['server']};"
        f"DATABASE={DB_SAP['database']};"
        f"UID={DB_SAP['user']};"
        f"PWD={DB_SAP['password']};"
        "Encrypt=yes;TrustServerCertificate=no;Connection Timeout=20;"
    )

def get_conn():
    return pyodbc.connect(_conn_str(), autocommit=True)


# ── Catálogos ─────────────────────────────────────────────────────────────────
PIEZAS = {
    "000": "Parabrisas",
    "001": "Lateral Delantero Izquierdo", "002": "Lateral Delantero Derecho",
    "003": "Lateral Trasero Izquierdo",   "004": "Lateral Trasero Derecho",
    "005": "Ventilete Trasero Izquierdo", "006": "Ventilete Trasero Derecho",
    "007": "Cabina Trasera Izquierda",    "008": "Cabina Trasera Derecha",
    "009": "Posterior",                   "010": "Techo Solar Delantero",
    "011": "Lateral Extendido Izquierdo", "012": "Lateral Extendido Derecho",
    "013": "Posterior Izquierdo",         "014": "Posterior Derecho",
    "015": "Claraboya Izquierda",         "016": "Claraboya Derecha",
    "017": "Mirilla",                     "018": "Probeta",
    "019": "Ventilete Delantero Izquierdo","020": "Ventilete Delantero Derecho",
    "021": "Cabina Delantera Izquierda",  "022": "Cabina Delantera Derecha",
    "023": "Cabina Superior Izquierda",   "024": "Cabina Superior Derecha",
    "025": "Techo Solar B",               "026": "Parabrisas Derecho",
    "027": "Parabrisas Izquierdo",        "028": "Lateral Secundario Derecho",
    "029": "Lateral Secundario Izquierdo","030": "Partición",
    "031": "Arquitectura",                "034": "Porthole 1",
    "035": "Porthole 2",                  "036": "Porthole 3",
    "037": "Porthole 4",                  "040": "Pummel",
    "085": "Posterior Secundario",        "087": "Techo Solar Céntrico",
    "088": "Techo Solar D",               "090": "Techo Solar Panorámico",
    "091": "Probeta 2",  "092": "Probeta 3", "093": "Probeta Especial",
    "094": "Probeta 4",  "095": "Kit Opaco", "096": "Probeta 5",
    "097": "Probeta 6",
    "110": "Techo Solar A — Paquete",     "125": "Techo Solar B — Paquete",
    "187": "Techo Solar C — Paquete",     "190": "Techo Solar Panorámico — Paquete",
}
for _i in range(1, 20):
    PIEZAS[f"{40+_i:03d}"] = f"Pieza Especial {_i}"
for _i in range(1, 11):
    PIEZAS[f"{59+_i:03d}"] = f"Vidrio Especial {_i}"
for _i, _n in enumerate([25, 26, 27, 28], 70):
    PIEZAS[f"{_i:03d}"] = f"Pieza Plana Especial {_n}"
for _i in range(80, 87):
    PIEZAS[f"{_i:03d}"] = "Vidrio Especial Laminado"

COLORES = {
    "NA": "No Aplica",       "00": "Blanco",
    "01": "Green Light",     "02": "Bronze Light",
    "03": "Azul",            "04": "Gray Light",
    "05": "Gray Light PC",   "06": "Gray Light Glass",
    "07": "Verde",           "08": "Bronze Medium",
    "09": "Gray Medium",     "10": "Gray Medium PC",
    "11": "Bronze Dark",     "12": "Gray Dark",
    "13": "Gray Dark Glass", "14": "Parsol Gray",
    "15": "Privacy",         "16": "Clear",
    "17": "Solar Green",     "18": "Gray Medium Glass",
    "19": "Gray Light Automotive",
    "20": "Gray Medium Automotive + PC", 
    "21": "Gray Dark Automotive + PC",
    "22": "G2 Gray Medium Automotive",
    "23": "G2 Gray Dark Automotive",
}

FRANJAS = {
    "00": "Sin Franja", "01": "Franja Azul",
    "02": "Franja Verde","03": "Franja Gris",
    "NA": "No Aplica",
}

PAISES = {
    "AE":"Emiratos Árabes Unidos","AF":"Afganistán","AR":"Argentina",
    "AT":"Austria","AU":"Australia","AX":"Islas de Åland","BE":"Bélgica",
    "BH":"Baréin","BO":"Bolivia","BR":"Brasil","BY":"Bielorrusia",
    "CA":"Canadá","CH":"Suiza","CL":"Chile","CN":"China","CO":"Colombia",
    "CR":"Costa Rica","CZ":"República Checa","DE":"Alemania","DK":"Dinamarca",
    "DM":"Dominica","DO":"República Dominicana","EC":"Ecuador","EG":"Egipto",
    "ES":"España","FI":"Finlandia","FR":"Francia","GB":"United Kingdom",
    "GR":"Grecia","GT":"Guatemala","HK":"Hong Kong","HN":"Honduras",
    "HR":"Croacia","HT":"Haití","ID":"Indonesia","IL":"Israel","IN":"India",
    "IQ":"Iraq","IT":"Italia","JE":"Jersey","JO":"Jordania","JP":"Japón",
    "KE":"Kenia","KR":"Corea del Sur","LB":"Líbano","MA":"Marruecos",
    "MX":"México","MY":"Malasia","NG":"Nigeria","NL":"Holanda","NO":"Noruega",
    "OM":"Omán","PA":"Panamá","PE":"Perú","PG":"Papúa Nueva Guinea",
    "PH":"Filipinas","PK":"Pakistán","PL":"Polonia","PR":"Puerto Rico",
    "PT":"Portugal","PY":"Paraguay","QA":"Qatar","RO":"Rumanía","RS":"Serbia",
    "SA":"Arabia Saudí","SE":"Suecia","SG":"Singapur","SK":"Eslovaquia",
    "SV":"El Salvador","TH":"Tailandia","TR":"Turquía","TW":"Taiwán",
    "US":"Estados Unidos","UY":"Uruguay","VE":"Venezuela","YE":"Yemen",
    "ZA":"Sudáfrica",
}

ATNAM_LABELS = {
    "Z_VEHICLE_MODEL":          "Modelo Vehículo",
    "Z_SUBPRODUCT":             "Subproducto",
    "Z_FORMULA_CODE":           "Fórmula",
    "Z_COLOR":                  "Color",
    "Z_PIECE_TYPE":             "Tipo de Pieza",
    "Z_SHADE_BAND":             "Franja",
    "Z_AGP_LEVEL":              "Nivel AGP",
    "Z_BEHAVIOR_DIFFERENTIALS": "Differentials",
    "Z_COMMERCIAL_THICKNESS":   "Espesor Comercial",
    "Z_AGP_VERSION":            "Versión AGP",
    "Z_AGP_PARTNUMBER":         "Partnumber AGP",
}


def _decode_route(route: str) -> str:
    """Intenta decodificar código de ruta SAP a nombre de país."""
    if not route:
        return "Sin ruta"
    r = route.strip().upper()
    if r in PAISES:
        return PAISES[r]
    # Formato "XX-YY": intentar prefijo y sufijo
    if "-" in r:
        partes = r.split("-")
        for p in reversed(partes):   # sufijo primero
            if p in PAISES:
                return PAISES[p]
    # Primeros 2 chars
    if len(r) >= 2 and r[:2] in PAISES:
        return PAISES[r[:2]]
    return route


# ── Queries ───────────────────────────────────────────────────────────────────

def q_zfer_head(material: str):
    """Tabla 1: ODATA_ZFER_HEAD — info básica del ZFER."""
    try:
        conn = get_conn()
        cur  = conn.cursor()
        cur.execute("""
            SELECT MATERIAL, CENTRO, TEXTO_BREVE_MATERIAL, STATUS,
                   ZFOR, GRUPO_ARTICULOS, CREADO_EL, ULTIMA_MOD
            FROM   dbo.ODATA_ZFER_HEAD
            WHERE  MATERIAL    = ?
              AND  CENTRO      = 'CO01'
              AND  UPPER(ISNULL(STATUS,'')) != 'ZZ'
        """, (material,))
        row = cur.fetchone()
        cols = [c[0] for c in cur.description]
        conn.close()
        return dict(zip(cols, row)) if row else None
    except Exception as e:
        return {"_error": str(e)}


def q_atributos(material: str) -> dict:
    """Tabla 2: ODATA_ZFER_CLASS_001 — atributos de clasificación."""
    try:
        conn = get_conn()
        cur  = conn.cursor()
        cur.execute("""
            SELECT ATNAM, ATWRT
            FROM   dbo.ODATA_ZFER_CLASS_001
            WHERE  MATERIAL = ?
              AND  CENTRO   = 'CO01'
              AND  ATNAM IN (
                'Z_AGP_LEVEL','Z_BEHAVIOR_DIFFERENTIALS','Z_VEHICLE_MODEL',
                'Z_AGP_PARTNUMBER','Z_SUBPRODUCT','Z_COLOR','Z_FORMULA_CODE',
                'Z_COMMERCIAL_THICKNESS','Z_AGP_VERSION','Z_PIECE_TYPE','Z_SHADE_BAND'
              )
        """, (material,))
        rows = cur.fetchall()
        conn.close()
        return {r[0]: (r[1] or "").strip() for r in rows}
    except Exception as e:
        return {"_error": str(e)}


def q_entregas(material: str) -> list:
    """Tabla 3: ODATA_ZCDS_Entregas_Pos_CO — números de entrega (ntgew > 0)."""
    try:
        conn = get_conn()
        cur  = conn.cursor()
        cur.execute("""
            SELECT DISTINCT entrega
            FROM   dbo.ODATA_ZCDS_Entregas_Pos_CO
            WHERE  matnr = ?
              AND  TRY_CAST(ntgew AS FLOAT) > 0
        """, (material,))
        rows = cur.fetchall()
        conn.close()
        return [str(r[0]) for r in rows if r[0] is not None]
    except Exception:
        return []


def _parsear_partnumber(pn: str) -> dict | None:
    """Parsea '1490_008_L23-26_12_000' → {vehiculo, version, formula, color, pieza}."""
    if not pn:
        return None
    parts = pn.strip().split("_")
    if len(parts) != 5:
        return None
    return {"vehiculo": parts[0], "version": parts[1], "formula": parts[2],
            "color": parts[3], "pieza": parts[4]}


def q_variantes_por_pn(vehiculo: str, version: str, formula: str, pieza: str) -> list:
    """
    Busca ZFERs activos (no ZZ) en CO01 cuyo PARTNUMBER comparte vehiculo+version+
    formula+pieza con cualquier color. Usa LIKE con ESCAPE para buscar en ODATA_ZFER_CLASS_001.
    """
    try:
        conn = get_conn()
        cur  = conn.cursor()
        # ESCAPE '!' — usar ! como caracter de escape (evita problemas con \ en pyodbc)
        def _esc(s):
            return s.replace("!", "!!").replace("%", "!%").replace("_", "!_")
        pattern = "!_".join([_esc(vehiculo), _esc(version), _esc(formula), "%", _esc(pieza)])

        cur.execute("""
            SELECT c.MATERIAL, c.ATWRT AS partnumber
            FROM   dbo.ODATA_ZFER_CLASS_001 c
            JOIN   dbo.ODATA_ZFER_HEAD h
                ON h.MATERIAL = c.MATERIAL AND h.CENTRO = 'CO01'
            WHERE  c.CENTRO = 'CO01'
              AND  c.ATNAM  = 'Z_AGP_PARTNUMBER'
              AND  c.ATWRT  LIKE ? ESCAPE '!'
              AND  UPPER(ISNULL(h.STATUS,'')) != 'ZZ'
            ORDER BY c.MATERIAL
        """, (pattern,))

        materiales_pn = {r[0]: r[1] for r in cur.fetchall()}
        if not materiales_pn:
            conn.close()
            return []

        mats = list(materiales_pn.keys())
        ph   = ",".join(["?"] * len(mats))

        # Atributos de color y franja
        cur.execute(f"""
            SELECT MATERIAL, ATNAM, ATWRT
            FROM   dbo.ODATA_ZFER_CLASS_001
            WHERE  CENTRO = 'CO01' AND MATERIAL IN ({ph})
              AND  ATNAM IN ('Z_COLOR', 'Z_SHADE_BAND')
        """, mats)
        pivot = {}
        for mat, atnam, atwrt in cur.fetchall():
            pivot.setdefault(mat, {})[atnam] = (atwrt or "").strip()

        # Status y descripción
        cur.execute(f"""
            SELECT MATERIAL, STATUS, TEXTO_BREVE_MATERIAL
            FROM   dbo.ODATA_ZFER_HEAD
            WHERE  CENTRO = 'CO01' AND MATERIAL IN ({ph})
        """, mats)
        head_d = {r[0]: {"status": (r[1] or "").strip(), "texto": (r[2] or "").strip()}
                  for r in cur.fetchall()}
        conn.close()

        resultado = []
        for mat in sorted(mats):
            d = pivot.get(mat, {})
            h = head_d.get(mat, {})
            color_raw = d.get("Z_COLOR", "")
            resultado.append({
                "material":     mat,
                "partnumber":   materiales_pn[mat],
                "color_raw":    color_raw,
                "color_nombre": COLORES.get(color_raw, color_raw) if color_raw else "—",
                "franja_raw":   d.get("Z_SHADE_BAND", ""),
                "status":       h.get("status", ""),
                "texto":        h.get("texto",  ""),
            })
        return resultado
    except Exception as e:
        return [{"_error": str(e)}]


def q_zplas_compatibles(formula_code: str, piece_type: str,
                        shade_band: str = "", differentials_base: str = "") -> list:
    """
    Busca ZPLAs en ODATA_ZPLA_CLASS_001 (CO01, TIPO_MAT=ZPLA) compatibles con
    la fórmula, tipo de pieza, franja y diferencial del ZFER base.
    Los atributos Z_PIECE_TYPE y Z_BEHAVIOR_DIFFERENTIALS son multi-valor (comas).
    """
    if not formula_code or not piece_type:
        return []
    try:
        conn = get_conn()
        cur  = conn.cursor()
        cur.execute("""
            SELECT
                MATERIAL,
                MAX(CASE WHEN ATNAM = 'Z_FORMULA_CODE'           THEN ATWRT ELSE NULL END) AS formula,
                MAX(CASE WHEN ATNAM = 'Z_COLOR'                  THEN ATWRT ELSE NULL END) AS color,
                MAX(CASE WHEN ATNAM = 'Z_PIECE_TYPE'             THEN ATWRT ELSE NULL END) AS piece_types,
                MAX(CASE WHEN ATNAM = 'Z_SHADE_BAND'             THEN ATWRT ELSE NULL END) AS shade_band,
                MAX(CASE WHEN ATNAM = 'Z_BEHAVIOR_DIFFERENTIALS' THEN ATWRT ELSE NULL END) AS differentials,
                MAX(CASE WHEN ATNAM = 'Z_AGP_LEVEL'              THEN ATWRT ELSE NULL END) AS level
            FROM dbo.ODATA_ZPLA_CLASS_001
            WHERE CENTRO   = 'CO01'
              AND TIPO_MAT = 'ZPLA'
            GROUP BY MATERIAL
            HAVING MAX(CASE WHEN ATNAM = 'Z_FORMULA_CODE' THEN ATWRT ELSE NULL END) = ?
        """, (formula_code,))
        rows = cur.fetchall()
        conn.close()

        # Diferencial(es) del ZFER base como set para comparar
        base_diffs = {d.strip() for d in differentials_base.split(",") if d.strip()}

        resultado = []
        for mat, formula, color, piece_types_str, zpla_shade, differentials, level in rows:
            if not color:
                continue
            # Z_PIECE_TYPE multi-valor: verificar que el tipo de pieza base esté incluido
            pieces = [p.strip() for p in (piece_types_str or "").split(",") if p.strip()]
            if piece_type not in pieces:
                continue
            # Franja: si el ZFER base tiene franja específica, el ZPLA debe coincidir
            if shade_band and shade_band not in ("00", ""):
                if (zpla_shade or "00") not in (shade_band, "00"):
                    continue
            # Diferencial: si el ZFER base tiene diferencial definido, el ZPLA debe contenerlo
            if base_diffs:
                zpla_diffs = {d.strip() for d in (differentials or "").split(",") if d.strip()}
                if not base_diffs.intersection(zpla_diffs):
                    continue
            resultado.append({
                "material":      mat,
                "color":         color.strip(),
                "color_nombre":  COLORES.get(color.strip(), color.strip()),
                "shade_band":    zpla_shade or "00",
                "differentials": differentials or "",
                "level":         level or "",
            })
        return sorted(resultado, key=lambda x: x["color"])
    except Exception as e:
        return [{"_error": str(e)}]


def q_mercados(entregas: list) -> list:
    """Tabla 4: ODATA_ZCDS_Entregas_Head_CO — conteo por route/mercado."""
    if not entregas:
        return []
    try:
        conn = get_conn()
        cur  = conn.cursor()
        ph   = ",".join(["?"] * len(entregas))
        cur.execute(f"""
            SELECT   route, COUNT(*) AS total
            FROM     dbo.ODATA_ZCDS_Entregas_Head_CO
            WHERE    entrega IN ({ph})
              AND    ISNULL(route,'') != ''
            GROUP BY route
            ORDER BY total DESC
        """, entregas)
        rows = cur.fetchall()
        conn.close()
        return [
            {"route": r[0], "pais": _decode_route(r[0]), "total": r[1]}
            for r in rows
        ]
    except Exception:
        return []


# ── Rutas Flask ───────────────────────────────────────────────────────────────

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        zfer = request.form.get("zfer", "").strip()
        if zfer:
            return redirect(url_for("detalle_zfer", material=zfer))
    return render_template("index.html", error=None)


@app.route("/zfer/<material>")
def detalle_zfer(material: str):
    material = material.strip()

    head = q_zfer_head(material)
    if head is None:
        return render_template("index.html",
            error=f"ZFER '{material}' no encontrado o STATUS = ZZ (inactivo).")
    if "_error" in head:
        return render_template("index.html",
            error=f"Error de conexión BD: {head['_error']}")

    attrs    = q_atributos(material)
    entregas = q_entregas(material)
    mercados = q_mercados(entregas)
 
    # Construir lista de atributos para mostrar (en orden definido)
    attrs_display = []
    for atnam, label in ATNAM_LABELS.items():
        val = attrs.get(atnam, "")
        if not val:
            continue
        decoded = val
        if atnam == "Z_COLOR":
            decoded = f"{val} — {COLORES.get(val, val)}"
        elif atnam == "Z_PIECE_TYPE":
            decoded = f"{val} — {PIEZAS.get(val, val)}"
        elif atnam == "Z_SHADE_BAND":
            decoded = f"{val} — {FRANJAS.get(val, val)}"
        attrs_display.append({
            "atnam":   atnam,
            "label":   label,
            "raw":     val,
            "decoded": decoded,
        })

    total_entregas = sum(m["total"] for m in mercados)
    # Top 15 para el gráfico; el resto en la tabla
    mercados_chart = mercados[:15]

    return render_template("zfer.html",
        material       = material,
        head           = head,
        attrs_display  = attrs_display,
        entregas_n     = len(entregas),
        mercados       = mercados,
        mercados_chart = mercados_chart,
        total_entregas = total_entregas,
    )


@app.route("/combinaciones/<material>")
def combinaciones(material: str):
    material = material.strip()

    head = q_zfer_head(material)
    if head is None:
        return render_template("index.html",
            error=f"ZFER '{material}' no encontrado o STATUS = ZZ (inactivo).")
    if "_error" in head:
        return render_template("index.html",
            error=f"Error de conexión BD: {head['_error']}")

    attrs = q_atributos(material)

    formula_code  = attrs.get("Z_FORMULA_CODE",         "")
    piece_type    = attrs.get("Z_PIECE_TYPE",            "")
    color_base    = attrs.get("Z_COLOR",                 "")
    shade_band    = attrs.get("Z_SHADE_BAND",            "00") or "00"
    partnumber    = attrs.get("Z_AGP_PARTNUMBER",        "")
    vehicle_model = attrs.get("Z_VEHICLE_MODEL",         "")
    thickness     = attrs.get("Z_COMMERCIAL_THICKNESS",  "")
    differentials = attrs.get("Z_BEHAVIOR_DIFFERENTIALS","")

    pn_parsed = _parsear_partnumber(partnumber)

    # Buscar variantes existentes vía PARTNUMBER (método preciso)
    if pn_parsed:
        variantes = q_variantes_por_pn(
            pn_parsed["vehiculo"], pn_parsed["version"],
            pn_parsed["formula"],  pn_parsed["pieza"]
        )
    else:
        variantes = []

    if variantes and "_error" in variantes[0]:
        return render_template("index.html",
            error=f"Error BD variantes: {variantes[0]['_error']}")

    # Buscar ZPLAs compatibles (fórmula + tipo pieza + franja + diferencial)
    zplas = q_zplas_compatibles(formula_code, piece_type, shade_band, differentials)
    if zplas and "_error" in zplas[0]:
        zplas = []

    # Mapa color → ZFER existente
    colores_con_zfer = {v["color_raw"]: v for v in variantes if v.get("color_raw")}
    # Mapa color → lista de ZPLAs (puede haber varios por color con distintos diferenciales)
    colores_con_zpla: dict = {}
    for z in zplas:
        colores_con_zpla.setdefault(z["color"], []).append(z)

    # Matriz completa: un item por color del catálogo (excepto NA)
    matrix = []
    for cod, nombre in COLORES.items():
        if cod == "NA":
            continue
        zfer_v    = colores_con_zfer.get(cod)
        zpla_list = colores_con_zpla.get(cod, [])
        if zfer_v:
            estado = "EXISTE"
        elif zpla_list:
            estado = "DISPONIBLE"
        else:
            estado = "SIN_ZPLA"
        matrix.append({
            "color_codigo":  cod,
            "color_nombre":  nombre,
            "estado":        estado,
            "zfer":          zfer_v["material"]           if zfer_v    else "",
            "zfer_texto":    zfer_v["texto"]              if zfer_v    else "",
            "zfer_pn":       zfer_v["partnumber"]         if zfer_v    else "",
            "zpla":          zpla_list[0]["material"]     if zpla_list else "",
            "zpla_count":    len(zpla_list),
            "zpla_list":     [z["material"] for z in zpla_list],
            "es_base":       cod == color_base,
        })

    n_existe     = sum(1 for c in matrix if c["estado"] == "EXISTE")
    n_disponible = sum(1 for c in matrix if c["estado"] == "DISPONIBLE")
    n_sin_zpla   = sum(1 for c in matrix if c["estado"] == "SIN_ZPLA")

    # Patrón LIKE usado para la búsqueda (para mostrar en UI)
    if pn_parsed:
        pn_pattern_ui = "_".join([
            pn_parsed["vehiculo"], pn_parsed["version"],
            pn_parsed["formula"], "**", pn_parsed["pieza"]
        ])
    else:
        pn_pattern_ui = ""

    return render_template("combinaciones.html",
        material       = material,
        head           = head,
        vehicle_model  = vehicle_model,
        formula_code   = formula_code,
        piece_type     = piece_type,
        piece_nombre   = PIEZAS.get(piece_type, piece_type),
        color_base     = color_base,
        shade_band     = shade_band,
        thickness      = thickness,
        differentials  = differentials,
        partnumber     = partnumber,
        pn_parsed      = pn_parsed,
        pn_pattern_ui  = pn_pattern_ui,
        variantes      = variantes,
        zplas          = zplas,
        matrix         = matrix,
        n_existe       = n_existe,
        n_disponible   = n_disponible,
        n_sin_zpla     = n_sin_zpla,
    )


if __name__ == "__main__":
    print("\n  AGP Intelligence — MODULO 5")
    print("  Abre tu navegador en: http://localhost:5000\n")
    app.run(debug=True, host="0.0.0.0", port=5000)
