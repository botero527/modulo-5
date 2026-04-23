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

DIFERENCIALES = {
    "01": "SOLAR PLUS",
    "02": "LIGHT WEIGHT",
    "03": "MULTI HIT",
    "04": "SUN ADVANCED",
    "05": "EXTREME PROTECT",
    "06": "STEEL PLUS",
    "07": "TNT",
    "08": "TNT FLEX",
    "09": "SUN BAND",
    "10": "GUNPORT",
    "11": "VARIO PLUS",
    "12": "AGP DURA P",
    "13": "AGP DURA NPC",
    "14": "AGP DURA G",
    "15": "HIGH PERFORMANCE",
    "16": "FRAMES",
    "17": "CLAMP",
    "18": "METALLIC SUPPORT FOR MIRROR",
    "19": "HEATING - METALLIC COATING",
    "20": "HEATING - WIRED - HEATPLEX",
    "21": "ANTIREFLECTIVE",
    "22": "SILVER PASTE",
    "23": "N.A",
    "24": "ENCAPSULATED - FRAMES",
}

SUBPRODUCTOS = {
    "B1":"B33","B2":"iB33","B3":"STANDARD","B4":"AGP PREMIUM","B5":"AGP TITANIUM",
    "B6":"OEM","B7":"ARCHITECTURAL BRG","B8":"B33 ESPECIAL","B9":"iB33 ESPECIAL",
    "B10":"STANDAR 40mm","B11":"STANDAR 45mm","B12":"3KL","B13":"3KL DURA P",
    "B14":"Antitheft","B15":"Impenetra Plus","B16":"Impenetra Plus DURA P",
    "B17":"NBR15000 II-A","B18":"iB33X",
    "D1":"LAND","D2":"NAVY","D3":"ARCHITECTURAL DEFENSE",
    "EA1":"STANDARD LAMINATED GLASS ARG","EA2":"STANDARD TOUGHENED GLASS ARG",
    "EA3":"MONOLITIC GLASS","EAV1":"AVO ONLY",
    "EL1":"ULT LAMINATED GLASS","EL2":"ULT LAMINATED GLASS W/AVO",
    "EL3":"STANDARD LAMINATED GLASS","EL4":"STAND LAMINATED GLASS W/AVO",
    "EL5":"SUP ULT LAMINATED GLASS","EL6":"SUP ULT LAMINATED GLASS W/AVO",
    "ET1":"ULTRALITE TOUGHNED GLASS","ET2":"ULTRALITE TOUGHNED GLASS W/AVO",
    "ET3":"STANDARD TOUGHENED  GLASS","ET4":"STAND TOUGHENED  GLASS  W/AVO",
    "M1":"BR6 Stoof","M2":"BR7 Stoof Opción 1","M3":"BR7 Stoof Opción 2 (con Heating)",
    "M4":"BR7 Farmingtons","M5":"Light Weigh DURA + SRF 14mm",
    "M6":"Estándar 18mm + SRF 14mm","M7":"VPAM 3 + SRF 14mm",
    "M8":"N5 Plasan Combinado","M9":"Marine NB",
    "M10":"Light Weight 28 y 30mm y VPAM 3","M11":"Estándar 42mm y VPAM 3","M12":"BR7 Ang 27G",
    "P1":"BR5 North Glass","P2":"Estándar 45mm North Glass","P3":"Nivel 4 Plasán",
    "S1":"Samples R&D",
    "X1":"21mm MH (POS LW 19mm & SRF 18mm)","X2":"21mm MH (SRF 18mm)",
    "X3":"B33 17mm","X4":"B33 17mm DURA P","X5":"B33 23mm  DURA",
    "X6":"B33 23mm (SRF 14mm)","X7":"B33 23mm (SRF 18mm)","X8":"B33 23mm DURA P",
    "X9":"B33 30mm - DURA","X10":"B33 43mm","X11":"B33 43mm DURA P",
    "X12":"BMW OEM VPAM 6","X13":"BR5 Tinted Galron","X14":"BR7 Stoof",
    "X15":"Estándar 18mm","X16":"Estándar 18mm DURA P","X17":"Estándar 21mm",
    "X18":"Estándar 21mm DURA P","X19":"Estándar 32mm","X20":"Estándar 32mm DURA P",
    "X21":"Estándar 33mm","X22":"Estándar 39mm","X23":"Estándar 40mm",
    "X24":"Estándar 40mm con Acero E","X25":"Estándar 40mm DURA P",
    "X26":"Estándar 42mm","X27":"Estándar 45mm","X28":"Estándar 45mm DURA P",
    "X29":"Estándar 48mm","X30":"Estándar 56mm","X31":"Estándar 56mm DURA P",
    "X32":"Estándar 58mm","X33":"Estándar 58mm DURA P","X34":"Estándar 60mm",
    "X35":"Estándar 60mm DURA P","X36":"Estándar 73mm","X37":"Estándar 76mm",
    "X38":"Estándar 76mm DURA P","X39":"Estándar 79mm","X40":"Estándar 79mm DURA P",
    "X41":"Estándar 82mm DURA P","X42":"Estándar 88mm","X43":"Estándar 88mm DURA P",
    "X44":"Estándar 97mm","X45":"Estándar 97mm DURA P",
    "X46":"Estándar BR3 17mm","X47":"Estándar BR3 18mm",
    "X48":"Estándar BR3 18mm DURA P","X49":"Estándar BR3 20mm DURA P",
    "X50":"Estándar Stop Gun 13mm","X51":"Estándar Stop Gun 14mm",
    "X52":"Estándar VPAM 3 15mm","X53":"Estándar VPAM 3 15mm DURA","X54":"GL43-01",
    "X55":"Light Weight 110mm DURA P","X56":"Light Weight 115mm DURA P",
    "X57":"Light Weight 19mm","X58":"Light Weight 19mm DURA","X59":"Light Weight 28mm",
    "X60":"Light Weight 28mm VSAG12","X61":"Light Weight 30mm",
    "X62":"Light Weight 30mm DURA P","X63":"Light Weight 36mm",
    "X64":"Light Weight 36mm DURA P","X65":"Light Weight 50mm",
    "X66":"Light Weight 50mm DURA P","X67":"Light Weight 50mm NORTH GLASS",
    "X68":"Light Weight 52mm DURA P","X69":"Light Weight 62mm",
    "X70":"Light Weight 62mm DURA P","X71":"Light Weight 69mm DURA P",
    "X72":"LW 19mm (LT´s 21mm & SRF 18mm)","X73":"LW 19mm (SRF Laminado)",
    "X74":"LW 69mm (PBS 50mm)","X75":"Marine Estándar 29mm",
    "X76":"Marine VPAM CL9 FRIGATTE 66 mm","X77":"Matine Estándar 40mm",
    "X78":"Mix 21mm STD y LW 19mm","X79":"MultiHit 21mm","X80":"MultiHit 21mm DURA P",
    "X81":"MultiHit 32mm","X82":"MultiHit 42mm","X83":"MultiHit 42mm DURA P",
    "X84":"NP58-2","X85":"NPC 85mm","X86":"PE NIJ III 38mm Blinsecurity",
    "X87":"PE STANAG 1 65mm TATRA","X88":"PE STANAG 1 Rheinmetall",
    "X89":"PE STANAG 2 60mm DURA P NIMR","X90":"PE WBS Rheinmetall",
    "X91":"Stop Gun 13mm DURA P","X92":"VPAM 3 15mm DURA P",
    "X93":"Estándar BR3 20mm","X94":"Estándar 24mm","X95":"Estándar 31mm",
    "X96":"N5 WBS Plasan 36mm","X97":"Estándar 40mm (Outer Glass 6mm)",
    "X98":"N5 WBS Plasan 43mm","X99":"Estándar 44mm","X100":"Estándar 45mm USA",
    "X101":"Estándar 47mm","X102":"BR7 Ang 55G Stoof","X103":"VPAM 9 Ang 55G Stoof",
    "X104":"Estándar 70mm","X105":"Estándar 72mm","X106":"Light Weight 66mm DURA P",
    "X107":"Estándar 71mm DURA P","X108":"Estándar 22mm","X109":"Marine NB 124",
    "X110":"Marine NB 155","X111":"Marine NB 124-1","X112":"Marine NB 103-2",
    "X113":"WBS 3 + 2","X114":"UL10 61mm",
    "X115":"Light Weight 30mm VPAM 6 OuterGlass 6mm","X116":"Estándar 48mm DURA P",
    "X117":"Estándar 67mm DURA P","X118":"Estándar 67mm","X119":"iB33 PLUS (FIJAS LW)",
    "X120":"AGP HEAT","X121":"Light Weight 66mm","X122":"L28CG y L28SCG",
    "X123":"3KL DUPA  P","X124":"B33 ESPECIAL VOLVO","X125":"iB33 NG",
    "X128":"B33 GEN2","X130":"Sunroof VPAM CL2 16mm DURA NPC",
    "X131":"Sunroof VPAM CL3 18mm DURA NPC","X132":"Estándar VPAM 2 11mm",
    "X139":"GL25-1","X140":"B33 EXPORTACIÓN LW","X141":"Ultra-Lightweight 12mm",
    "X142":"Envostar","X143":"iB33 G6","X145":"iB33 G6 EXPORTAÇÃO",
    "X146":"Estándar 30mm","X147":"Estándar VPAM CL2 17mm DURA NPC",
    "X149":"Marine NB 20mm","X150":"B33 23mm","X151":"B33 28mm",
    "X152":"LAMINADO 11mm","X153":"Estándar 50mm PBS",
    "X154":"Light Weight 18mm ARGENTINA","X155":"Estándar 32mm GALRON",
    "X156":"Light Weight 19mm LATAM","X157":"WBS 28mm DURA P ENVOSTAR",
    "X158":"Estándar 28mm CAM","X159":"Light Weight 70mm","X160":"Estándar 45mm URO",
    "X161":"Estándar 83mm","X162":"Estándar 69mm TPS MÉXICO",
    "X163":"Estándar 80mm DURA P","X164":"Estándar 102 DURA P NMIR",
    "X165":"Estándar 74mm DURA P PBS TENCATE",
    "X166":"Estándar 38mm DURA P NORTHGLASS","X167":"Estándar 38mm DURA P",
    "X168":"Estándar 43mm","X169":"Estándar 44mm DURA P USA",
    "X170":"Estándar 42mm NORTHGLASS","X171":"Estándar 88mm URO",
    "X172":"Estándar 55mm PBS ANGULO","X173":"Estándar 66mm PBS ANGULO",
    "X174":"Estándar 70mm STOOF","X175":"Estándar 82mm",
    "X176":"Estándar 29mm DURA P","X177":"Estándar 45mm ALEMANIA",
    "X178":"Estándar 44mm DURA P NORTHGLASS","X179":"Estándar 145mm DURA P",
    "X180":"Estándar 26mm DURA P MARINE","X181":"Estándar 38mm NORTHGLASS",
    "X182":"Estándar 85mm PBS","X183":"Estándar 86mm RHEINMETAL",
    "X184":"Estándar 61mm DURA P TENCATE","X185":"WBS 33mm DURA P ENVOSTAR",
    "X186":"Estándar 47mm DURA P","X187":"Estándar 43mm DURA P",
    "X189":"WBS 35mm PLASAN","X190":"Estándar 32mm DURA P NORTHGLASS",
    "X191":"Estándar 32mm NO EUROPA","X192":"Estándar 70mm DURA P",
    "X193":"Estándar 30mm DURA P","X194":"Estándar 50mm DURA P JANKEL",
    "X195":"Stop Gun 17mm DURA P NORTHGLASS","X196":"Estándar 84mm DURA P",
    "X197":"Exclusivo USA 49mm DOS","X198":"Estándar 86mm",
    "X199":"Estándar 60mm SENTINEL","X200":"Estándar 65mm",
    "X201":"Estándar 97mm CAMBLI","X202":"Estándar 46mm DURA P NORTHGLASS",
    "X203":"Estándar 62mm","X204":"Estándar 81mm MEXICO","X205":"Estándar 51mm GREIT",
    "X206":"Estándar 54mm GREIT","X207":"Estándar 33mm PLASAN",
    "X208":"Estándar 74mm DURA P TENCTE","X209":"Multihit 114mm DURA P NIMR",
    "X210":"Multihit 110mm DURA P NIMR","X211":"Estándar 42mm LICITACIÓN",
    "X212":"Arquitectónico 26mm DURA P","X213":"Light Weight 71mm",
    "X214":"WBS 7mm PLASAN","X215":"WBS 5mm NIMR","X216":"Exclusivo USA 52mm",
    "X217":"Estándar 22mm DURA P BOON EDAM",
    "X218":"Doble cara impacto 27mm DURA P BOON EDAM",
    "X219":"Estándar 22mm DURA P NORTHGLASS","X220":"WBS 21mm DURA P PMMA",
    "X221":"WBS 26mm  DURA P PMMA","X222":"FGR 47mm DURA NPC NAVANTIA",
    "X223":"Estándar 65mm PBS TNTF","X224":"WBS  35mm MIRLL AEROSPACE DURA P",
    "X225":"Estándar 47mm DEFENTURE","X226":"Light Weight 65mm DURA P",
    "X227":"Estándar 83mm DURA P AEROSPACE","X228":"WBS 8mm TECNOGETAFE",
    "X231":"Estándar 95mm","X233":"Multihit 43mm","X237":"OSOP","X241":"Estándar 51mm",
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
            SELECT ATNAM,
                   CASE WHEN ATNAM = 'Z_COMMERCIAL_THICKNESS' THEN CAST(ATFLV AS VARCHAR(50)) ELSE ATWRT END AS valor
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
        return {r[0]: str(r[1]).strip() if r[1] is not None else "" for r in rows}
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

#AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
#necesitaba expresarme
#sigamos 
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
            pivot.setdefault(mat, {})[atnam] = str(atwrt).strip() if atwrt is not None else ""

        # Status y descripción
        cur.execute(f"""
            SELECT MATERIAL, STATUS, TEXTO_BREVE_MATERIAL
            FROM   dbo.ODATA_ZFER_HEAD
            WHERE  CENTRO = 'CO01' AND MATERIAL IN ({ph})
        """, mats)
        head_d = {r[0]: {"status": str(r[1]).strip() if r[1] is not None else "",
                          "texto":  str(r[2]).strip() if r[2] is not None else ""}
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
        # Pre-filtra por fórmula ANTES del GROUP BY para no agrupar toda la tabla
        cur.execute("""
            SELECT
                c.MATERIAL,
                MAX(CASE WHEN c.ATNAM = 'Z_COLOR'                  THEN c.ATWRT ELSE NULL END) AS color,
                MAX(CASE WHEN c.ATNAM = 'Z_PIECE_TYPE'             THEN c.ATWRT ELSE NULL END) AS piece_types,
                MAX(CASE WHEN c.ATNAM = 'Z_SHADE_BAND'             THEN c.ATWRT ELSE NULL END) AS shade_band,
                MAX(CASE WHEN c.ATNAM = 'Z_BEHAVIOR_DIFFERENTIALS' THEN c.ATWRT ELSE NULL END) AS differentials,
                MAX(CASE WHEN c.ATNAM = 'Z_AGP_LEVEL'              THEN c.ATWRT ELSE NULL END) AS level
            FROM dbo.ODATA_ZPLA_CLASS_001 c
            JOIN dbo.ODATA_ZPLA_HEAD h
              ON h.MATERIAL = c.MATERIAL AND h.CENTRO = 'CO01'
            WHERE c.CENTRO   = 'CO01'
              AND c.TIPO_MAT = 'ZPLA'
              AND UPPER(ISNULL(h.STATUS, '')) != 'ZZ'
              AND c.MATERIAL IN (
                SELECT MATERIAL FROM dbo.ODATA_ZPLA_CLASS_001
                WHERE CENTRO = 'CO01' AND TIPO_MAT = 'ZPLA'
                  AND ATNAM  = 'Z_FORMULA_CODE' AND ATWRT = ?
              )
            GROUP BY c.MATERIAL
        """, (formula_code,))
        rows = cur.fetchall()
        conn.close()

        # Diferencial(es) del ZFER base como set para comparar
        base_diffs = {d.strip() for d in differentials_base.split(",") if d.strip()}

        resultado = []
        # Query devuelve 6 cols: MATERIAL, color, piece_types, shade_band, differentials, level
        for mat, color, piece_types_str, zpla_shade, differentials, level in rows:
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
            # Diferencial: el ZPLA debe contener AL MENOS UNO de los diferenciales del ZFER base
            # Si el ZFER base no tiene diferencial definido, no filtrar
            if base_diffs:
                zpla_diffs = {d.strip() for d in (differentials or "").split(",") if d.strip()}
                if zpla_diffs and not base_diffs.intersection(zpla_diffs):
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


def q_explorar(vehiculo="", formula="", pieza="", color="", version="", nivel="",
               cod_vehiculo="", zfers_lista: list = None) -> list:
    """
    Busca ZFERs activos (no ZZ) en CO01 según filtros opcionales (LIKE parcial).
    Si se pasa zfers_lista, busca exactamente esos ZFERs y los enriquece con atributos.
    Retorna lista de dicts con los atributos clave de cada ZFER.
    Máximo 300 resultados.
    """
    def _esc(s):
        return s.replace("!", "!!").replace("%", "!%").replace("_", "!_")

    try:
        conn = get_conn()
        cur  = conn.cursor()

        if zfers_lista:
            # Búsqueda directa por lista de materiales
            ph = ",".join(["?"] * len(zfers_lista))
            cur.execute(f"""
                SELECT MATERIAL FROM dbo.ODATA_ZFER_HEAD
                WHERE  MATERIAL IN ({ph}) AND CENTRO = 'CO01'
                  AND  UPPER(ISNULL(STATUS,'')) != 'ZZ'
            """, zfers_lista)
            materiales = [str(r[0]) for r in cur.fetchall()]
        else:
            # Búsqueda por filtros con INTERSECT dinámico
            filtros = [
                ("Z_VEHICLE_MODEL", vehiculo.strip()),
                ("Z_FORMULA_CODE",  formula.strip()),
                ("Z_PIECE_TYPE",    pieza.strip()),
                ("Z_COLOR",         color.strip()),
                ("Z_AGP_VERSION",   version.strip()),
                ("Z_AGP_LEVEL",     nivel.strip()),
            ]
            activos = [(a, v) for a, v in filtros if v]

            # Un solo scan con OR + GROUP BY/HAVING en lugar de N INTERSECTs
            or_parts, params = [], []
            for atnam, val in activos:
                or_parts.append("(c.ATNAM=? AND c.ATWRT LIKE ? ESCAPE '!')")
                params.extend([atnam, f"%{_esc(val)}%"])
            # Código vehículo: prefijo del PARTNUMBER (ej: "1715" → "1715_...")
            if cod_vehiculo.strip():
                or_parts.append("(c.ATNAM='Z_AGP_PARTNUMBER' AND c.ATWRT LIKE ? ESCAPE '!')")
                params.append(f"{_esc(cod_vehiculo.strip())}!_%")

            if not or_parts:
                conn.close()
                return []

            n = len(activos) + (1 if cod_vehiculo.strip() else 0)
            cur.execute(f"""
                SELECT TOP 300 c.MATERIAL
                FROM dbo.ODATA_ZFER_CLASS_001 c
                JOIN dbo.ODATA_ZFER_HEAD h
                  ON h.MATERIAL = c.MATERIAL AND h.CENTRO = 'CO01'
                WHERE c.CENTRO = 'CO01'
                  AND UPPER(ISNULL(h.STATUS,'')) != 'ZZ'
                  AND ({" OR ".join(or_parts)})
                GROUP BY c.MATERIAL
                HAVING COUNT(DISTINCT c.ATNAM) >= {n}
                ORDER BY c.MATERIAL
            """, params)
            materiales = [str(r[0]) for r in cur.fetchall()]

        if not materiales:
            conn.close()
            return []

        ph = ",".join(["?"] * len(materiales))

        # Atributos clave para mostrar en tabla
        cur.execute(f"""
            SELECT MATERIAL, ATNAM, ATWRT
            FROM   dbo.ODATA_ZFER_CLASS_001
            WHERE  CENTRO = 'CO01' AND MATERIAL IN ({ph})
              AND  ATNAM IN (
                'Z_VEHICLE_MODEL','Z_FORMULA_CODE','Z_COLOR',
                'Z_PIECE_TYPE','Z_AGP_VERSION','Z_AGP_PARTNUMBER',
                'Z_SHADE_BAND','Z_BEHAVIOR_DIFFERENTIALS','Z_AGP_LEVEL'
              )
        """, materiales)
        pivot = {}
        for mat, atnam, atwrt in cur.fetchall():
            pivot.setdefault(str(mat), {})[atnam] = str(atwrt).strip() if atwrt is not None else ""

        # Cabecera (status, descripción, ZFOR)
        cur.execute(f"""
            SELECT MATERIAL, STATUS, TEXTO_BREVE_MATERIAL, ZFOR
            FROM   dbo.ODATA_ZFER_HEAD
            WHERE  CENTRO = 'CO01' AND MATERIAL IN ({ph})
        """, materiales)
        head_d = {str(r[0]): {"status": str(r[1]).strip() if r[1] is not None else "",
                          "texto":  str(r[2]).strip() if r[2] is not None else "",
                          "zfor":   str(r[3]).strip() if r[3] is not None else ""}
                  for r in cur.fetchall()}
        conn.close()

        resultado = []
        for mat in sorted(materiales):
            d = pivot.get(mat, {})
            h = head_d.get(mat, {})
            color_raw  = d.get("Z_COLOR", "")
            pieza_raw  = d.get("Z_PIECE_TYPE", "")
            resultado.append({
                "material":      mat,
                "texto":         h.get("texto", ""),
                "status":        h.get("status", ""),
                "zfor":          h.get("zfor", ""),
                "vehiculo":      d.get("Z_VEHICLE_MODEL", ""),
                "formula":       d.get("Z_FORMULA_CODE", ""),
                "color_raw":     color_raw,
                "color_nombre":  COLORES.get(color_raw, color_raw),
                "pieza_raw":     pieza_raw,
                "pieza_nombre":  PIEZAS.get(pieza_raw, pieza_raw),
                "version":       d.get("Z_AGP_VERSION", ""),
                "partnumber":    d.get("Z_AGP_PARTNUMBER", ""),
                "shade_band":    d.get("Z_SHADE_BAND", ""),
                "differentials": d.get("Z_BEHAVIOR_DIFFERENTIALS", ""),
                "nivel":         d.get("Z_AGP_LEVEL", ""),
            })
        return resultado
    except Exception as e:
        return [{"_error": str(e)}]


def q_valores_distintos(atnam: str) -> list:
    """Devuelve los 200 valores ATWRT distintos más frecuentes para un ATNAM en CO01."""
    try:
        conn = get_conn()
        cur  = conn.cursor()
        cur.execute("""
            SELECT TOP 200 ATWRT, COUNT(*) AS n
            FROM   dbo.ODATA_ZFER_CLASS_001
            WHERE  CENTRO = 'CO01' AND ATNAM = ?
              AND  ISNULL(ATWRT,'') != ''
            GROUP BY ATWRT
            ORDER BY n DESC
        """, (atnam,))
        rows = cur.fetchall()
        conn.close()
        return [r[0] for r in rows]
    except Exception:
        return []


# ── Rutas Flask ───────────────────────────────────────────────────────────────

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        raw = request.form.get("zfer", "").strip()
        if not raw:
            return render_template("index.html", error=None)
        # Si hay comas → multi-ZFER → explorar
        zfers = [z.strip() for z in raw.replace(";", ",").split(",") if z.strip()]
        if len(zfers) > 1:
            return redirect(url_for("explorar") + "?zfers=" + ",".join(zfers))
        return redirect(url_for("detalle_zfer", material=zfers[0]))
    return render_template("index.html", error=None)


@app.route("/explorar")
def explorar():
    vehiculo = request.args.get("vehiculo", "").strip()
    formula  = request.args.get("formula",  "").strip()
    pieza    = request.args.get("pieza",    "").strip()
    color    = request.args.get("color",    "").strip()
    version  = request.args.get("version",  "").strip()
    nivel        = request.args.get("nivel",        "").strip()
    cod_vehiculo = request.args.get("cod_vehiculo", "").strip()
    zfers_qs     = request.args.get("zfers",        "").strip()

    zfers_lista = [z.strip() for z in zfers_qs.split(",") if z.strip()] if zfers_qs else []

    hay_filtros = any([vehiculo, formula, pieza, color, version, nivel, cod_vehiculo]) or bool(zfers_lista)
    resultados  = []
    error       = None

    if hay_filtros:
        resultados = q_explorar(vehiculo, formula, pieza, color, version, nivel, cod_vehiculo, zfers_lista or None)
        if resultados and "_error" in resultados[0]:
            error      = resultados[0]["_error"]
            resultados = []

    # Autocomplete: solo carga hints cuando el usuario ya busca (evita 2 queries extra en carga inicial)
    vehiculos_hints = q_valores_distintos("Z_VEHICLE_MODEL") if hay_filtros else []
    formulas_hints  = q_valores_distintos("Z_FORMULA_CODE")  if hay_filtros else []

    return render_template("explorar.html",
        vehiculo        = vehiculo,
        formula         = formula,
        pieza           = pieza,
        color           = color,
        version         = version,
        nivel           = nivel,
        cod_vehiculo    = cod_vehiculo,
        zfers_qs        = zfers_qs,
        resultados      = resultados,
        error           = error,
        hay_filtros     = hay_filtros,
        modo_lista      = bool(zfers_lista),
        vehiculos_hints = vehiculos_hints,
        formulas_hints  = formulas_hints,
        COLORES         = COLORES,
        PIEZAS          = PIEZAS,
        FRANJAS         = FRANJAS,
    )


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
        DIFERENCIALES  = DIFERENCIALES,
        SUBPRODUCTOS   = SUBPRODUCTOS,
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
#aa hahah  hdh hadh sajddh shjdh  hjsdkjh  PARA LIBERAR ESTRES HPTAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA

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
