"""
autocolor.py — Auto Color: cola paralela de cambio masivo de color por ZFER.

Ejecucion unicamente manual (sin timer/scheduler).
Se integra con app.py via init() + register_blueprint().

Dependencias inyectadas desde app.py:
  - _get_conn_local, _sap_bloque_lock y sus helpers
  - _COLORES_ACTIVOS, COLORES, q_atributos, q_variantes_por_pn, q_zplas_compatibles
  - _parsear_partnumber, _hr_buscar_candidata
"""

import threading
import importlib
import datetime
from concurrent.futures import ThreadPoolExecutor
from functools import wraps

from flask import Blueprint, request, jsonify, session, redirect, url_for

bp = Blueprint("auto_color", __name__, url_prefix="/api/auto_color")

# ── Dependencias inyectadas por init() ─────────────────────────────────────────
_get_conn_local             = None
_sap_bloque_lock            = None
_sap_lock_ocupado_por_otro  = None
_sap_lock_insertar_items    = None
_sap_lock_keepalive         = None
_sap_lock_keepalive_loop    = None
_sap_lock_limpiar_proyecto  = None
_COLORES_ACTIVOS: set       = set()
_COLORES: dict              = {}
_q_atributos                = None
_q_variantes_por_pn         = None
_q_zplas_compatibles        = None
_parsear_partnumber         = None
_hr_buscar_candidata        = None   # misma funcion que usa la cola normal

_dt = datetime.datetime

# ── Estado interno ─────────────────────────────────────────────────────────────
_ac_disparados: set = set()   # bloque_ids en ejecucion actualmente


# ── Inicializacion ─────────────────────────────────────────────────────────────

def init(
    get_conn_local,
    sap_bloque_lock,
    sap_lock_ocupado_por_otro,
    sap_lock_insertar_items,
    sap_lock_keepalive,
    sap_lock_keepalive_loop,
    sap_lock_limpiar_proyecto,
    colores_activos: set,
    colores_dict: dict,
    q_atributos,
    q_variantes_por_pn,
    q_zplas_compatibles,
    parsear_partnumber,
    hr_buscar_candidata=None,
):
    """Inyecta dependencias compartidas desde app.py. Sin scheduler — ejecucion manual."""
    global _get_conn_local, _sap_bloque_lock
    global _sap_lock_ocupado_por_otro, _sap_lock_insertar_items
    global _sap_lock_keepalive, _sap_lock_keepalive_loop, _sap_lock_limpiar_proyecto
    global _COLORES_ACTIVOS, _COLORES
    global _q_atributos, _q_variantes_por_pn, _q_zplas_compatibles
    global _parsear_partnumber, _hr_buscar_candidata

    _get_conn_local             = get_conn_local
    _sap_bloque_lock            = sap_bloque_lock
    _sap_lock_ocupado_por_otro  = sap_lock_ocupado_por_otro
    _sap_lock_insertar_items    = sap_lock_insertar_items
    _sap_lock_keepalive         = sap_lock_keepalive
    _sap_lock_keepalive_loop    = sap_lock_keepalive_loop
    _sap_lock_limpiar_proyecto  = sap_lock_limpiar_proyecto
    _COLORES_ACTIVOS            = colores_activos
    _COLORES                    = colores_dict
    _q_atributos                = q_atributos
    _q_variantes_por_pn         = q_variantes_por_pn
    _q_zplas_compatibles        = q_zplas_compatibles
    _parsear_partnumber         = parsear_partnumber
    _hr_buscar_candidata        = hr_buscar_candidata

    # Recuperar bloques/items que quedaron EJECUTANDO por crash del servidor
    try:
        with get_conn_local() as cn:
            cur = cn.cursor()
            cur.execute("""
                UPDATE itg.M5_AutoColor_BLOQUES
                SET estado='ERROR'
                WHERE estado='EJECUTANDO'
            """)
            cur.execute("""
                UPDATE itg.M5_AutoColor_COLA
                SET estado='PENDIENTE'
                WHERE estado='EJECUTANDO'
            """)
            rows = cur.rowcount
        if rows:
            print(f"[AUTOCOLOR] recuperacion arranque: {rows} item(s) EJECUTANDO → PENDIENTE")
    except Exception as e:
        print(f"[AUTOCOLOR] advertencia recuperacion arranque: {e}")

    print("[AUTOCOLOR] modulo iniciado — solo ejecucion manual")


def get_migration_sqls() -> list:
    """DDL para las 3 tablas de Auto Color (idempotente, IF NOT EXISTS)."""
    return [
        """IF OBJECT_ID('itg.M5_AutoColorZfer','U') IS NULL
           CREATE TABLE itg.M5_AutoColorZfer (
               id INT IDENTITY(1,1) PRIMARY KEY,
               zfer_base NVARCHAR(20) NOT NULL,
               estado NVARCHAR(20) NOT NULL DEFAULT 'PENDIENTE',
               cargado_el DATETIME NOT NULL DEFAULT GETDATE()
           )""",
        """IF OBJECT_ID('itg.M5_AutoColor_BLOQUES','U') IS NULL
           CREATE TABLE itg.M5_AutoColor_BLOQUES (
               id INT IDENTITY(1,1) PRIMARY KEY,
               bloque_num INT NOT NULL DEFAULT 1,
               estado NVARCHAR(20) NOT NULL DEFAULT 'PENDIENTE',
               cambiar_hr BIT NOT NULL DEFAULT 0,
               creado_el DATETIME NOT NULL DEFAULT GETDATE(),
               ejecutado_el DATETIME NULL,
               ok_count INT NOT NULL DEFAULT 0,
               error_count INT NOT NULL DEFAULT 0
           )""",
        """IF OBJECT_ID('itg.M5_AutoColor_COLA','U') IS NULL
           CREATE TABLE itg.M5_AutoColor_COLA (
               id INT IDENTITY(1,1) PRIMARY KEY,
               bloque_id INT NOT NULL,
               autocolor_zfer_id INT NULL,
               zfer_base NVARCHAR(20) NOT NULL,
               color NVARCHAR(10) NULL,
               color_nombre NVARCHAR(100) NULL,
               zpla NVARCHAR(20) NULL,
               franja NVARCHAR(5) NOT NULL DEFAULT '00',
               pn_base NVARCHAR(50) NULL,
               nivel NVARCHAR(10) NULL,
               tipo_pieza NVARCHAR(10) NULL,
               cambiar_hr BIT NOT NULL DEFAULT 0,
               estado NVARCHAR(20) NOT NULL DEFAULT 'PENDIENTE',
               zfer_nuevo NVARCHAR(20) NULL,
               error_msg NVARCHAR(500) NULL,
               creado_el DATETIME NOT NULL DEFAULT GETDATE(),
               ejecutado_el DATETIME NULL
           )""",
    ]


# ── Auth helper ────────────────────────────────────────────────────────────────

def _login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("usuario"):
            return redirect(url_for("login", next=request.path))
        return f(*args, **kwargs)
    return wrapper


# ── Logica de colores disponibles ─────────────────────────────────────────────

def _colores_disponibles_para_zfer(zfer_base: str) -> tuple:
    """
    Retorna (colores_list, attrs_dict).
    colores_list = colores sin ZFER existente que tienen ZPLA compatible.
    En error retorna (None, error_str).
    """
    attrs = _q_atributos(zfer_base)
    if "_error" in attrs:
        return None, attrs["_error"]

    formula_code  = attrs.get("Z_FORMULA_CODE", "")
    piece_type    = attrs.get("Z_PIECE_TYPE",   "")
    shade_band    = attrs.get("Z_SHADE_BAND",   "00") or "00"
    differentials = attrs.get("Z_BEHAVIOR_DIFFERENTIALS", "")
    tiene_acero   = "06" in {d.strip() for d in differentials.split(",") if d.strip()}
    pn_parsed     = _parsear_partnumber(attrs.get("Z_AGP_PARTNUMBER", ""))

    with ThreadPoolExecutor(max_workers=2) as ex:
        fut_var  = ex.submit(
            _q_variantes_por_pn,
            pn_parsed["vehiculo"], pn_parsed["version"],
            pn_parsed["formula"],  pn_parsed["pieza"]
        ) if pn_parsed else None
        fut_zpla = ex.submit(
            _q_zplas_compatibles, formula_code, piece_type, shade_band, differentials, tiene_acero
        )

    variantes = (fut_var.result() if fut_var else []) or []
    zplas     = fut_zpla.result() or []
    if variantes and "_error" in variantes[0]: variantes = []
    if zplas     and "_error" in zplas[0]:     zplas     = []

    colores_con_zfer = {v["color_raw"]: v for v in variantes if v.get("color_raw")}
    colores_con_zpla: dict = {}
    for z in zplas:
        colores_con_zpla.setdefault(z["color"], []).append(z)

    colores = []
    for cod in _COLORES_ACTIVOS:
        if cod in colores_con_zfer:
            continue
        zpla_list = colores_con_zpla.get(cod, [])
        if not zpla_list:
            continue
        colores.append({
            "color_codigo": cod,
            "color_nombre": _COLORES.get(cod, cod),
            "zpla":         zpla_list[0]["material"],
        })
    colores.sort(key=lambda x: x["color_codigo"])
    return colores, attrs


# ── Ejecucion SAP ──────────────────────────────────────────────────────────────

def _autocolor_ejecutar_bloque(bloque_id: int, ejecutado_por: str = "manual"):
    """Ejecuta todos los items PENDIENTE del bloque via procesar_combinacion (COLOR)."""
    ok_n, err_n = 0, 0

    if not _sap_bloque_lock.acquire(blocking=False):
        print(f"[AUTOCOLOR] bloque {bloque_id}: SAP ocupado por otro bloque")
        _ac_disparados.discard(bloque_id)
        return

    try:
        ocupado_por = _sap_lock_ocupado_por_otro()
        if ocupado_por:
            print(f"[AUTOCOLOR] bloque {bloque_id}: SAP ocupado por '{ocupado_por}'")
            _ac_disparados.discard(bloque_id)
            return

        with _get_conn_local() as cn:
            cur = cn.cursor()
            cur.execute(
                "UPDATE itg.M5_AutoColor_BLOQUES SET estado='EJECUTANDO' WHERE id=?", bloque_id
            )
            # Leer cambiar_hr a nivel de bloque
            row_b = cur.execute(
                "SELECT cambiar_hr FROM itg.M5_AutoColor_BLOQUES WHERE id=?", bloque_id
            ).fetchone()
            bloque_cambiar_hr = bool(row_b[0]) if row_b else False

            cur.execute("""
                SELECT id, zfer_base, color, color_nombre, zpla, franja,
                       pn_base, nivel, tipo_pieza
                FROM itg.M5_AutoColor_COLA
                WHERE bloque_id=? AND estado='PENDIENTE'
            """, bloque_id)
            rows = cur.fetchall()

        if not rows:
            with _get_conn_local() as cn:
                cn.cursor().execute(
                    "UPDATE itg.M5_AutoColor_BLOQUES SET estado='COMPLETADO', ejecutado_el=GETDATE() WHERE id=?",
                    bloque_id
                )
            _ac_disparados.discard(bloque_id)
            return

        cola = [
            {
                "_id":          r[0],
                "zfer":         r[1],
                "color":        r[2],
                "color_nombre": r[3],
                "zpla":         r[4],
                "franja":       r[5] or "00",
                "pn_base":      r[6] or "",
                "nivel":        r[7] or "",
                "tipo_pieza":   r[8] or "",
                "cambiar_hr":   bloque_cambiar_hr,
                "_cola_id":     r[0],
                "_bloque_id":   bloque_id,
                "tipo":         "COLOR",
            }
            for r in rows
        ]

        _sap_lock_insertar_items(cola)
        _sap_lock_stop = threading.Event()
        threading.Thread(
            target=_sap_lock_keepalive_loop, args=(_sap_lock_stop,), daemon=True
        ).start()

        try:
            sap        = importlib.import_module("sap_auto")
            proc_color = getattr(sap, "procesar_combinacion", None)

            for item in cola:
                with _get_conn_local() as cn:
                    cn.cursor().execute(
                        "UPDATE itg.M5_AutoColor_COLA SET estado='EJECUTANDO' WHERE id=?",
                        item["_id"]
                    )

                estado_item = "ERROR"
                zfer_nuevo  = ""
                msg_err     = ""
                try:
                    if proc_color:
                        res = proc_color(
                            zfer_base=item["zfer"],
                            color_codigo=item["color"],
                            color_nombre=item["color_nombre"],
                            zpla=item["zpla"],
                            franja=item["franja"],
                            pn_base=item["pn_base"],
                            nivel=item["nivel"],
                            tipo_pieza=item["tipo_pieza"],
                        )
                        if res and getattr(res, "estado", "") == "OK":
                            estado_item = "OK"
                            zfer_nuevo  = getattr(res, "zfer_nuevo", "") or ""
                        else:
                            msg_err = getattr(res, "error", "Sin resultado") if res else "Sin resultado"
                    else:
                        msg_err = "procesar_combinacion no encontrada en sap_auto"
                except Exception as ex:
                    msg_err = str(ex)
                    print(f"[AUTOCOLOR] error item {item['zfer']}/{item['color']}: {ex}")

                # Cambio de HR si corresponde
                msg_hr = ""
                if estado_item == "OK" and zfer_nuevo and item["cambiar_hr"] and _hr_buscar_candidata:
                    try:
                        hr_id, hr_desc, hr_err = _hr_buscar_candidata(item["zfer"], zfer_nuevo)
                        if hr_err:
                            msg_hr = f"HR-WARN: {hr_err}"
                            print(f"[AUTOCOLOR] {msg_hr}")
                        else:
                            print(f"[AUTOCOLOR] HR asignada {hr_id} para {zfer_nuevo}")
                    except Exception as hr_ex:
                        msg_hr = f"HR-ERR: {hr_ex}"
                        print(f"[AUTOCOLOR] {msg_hr}")

                error_final = " | ".join(filter(None, [msg_err, msg_hr])) or None

                with _get_conn_local() as cn:
                    cn.cursor().execute("""
                        UPDATE itg.M5_AutoColor_COLA
                        SET estado=?, zfer_nuevo=?, error_msg=?, ejecutado_el=GETDATE()
                        WHERE id=?
                    """, estado_item,
                        zfer_nuevo[:20] if zfer_nuevo else None,
                        error_final[:500] if error_final else None,
                        item["_id"])

                _sap_lock_keepalive()
                if estado_item == "OK":
                    ok_n  += 1
                else:
                    err_n += 1

        except Exception as sap_ex:
            print(f"[AUTOCOLOR] error cargando sap_auto: {sap_ex}")
        finally:
            _sap_lock_stop.set()
            _sap_lock_limpiar_proyecto()

    except Exception as ex:
        print(f"[AUTOCOLOR] error general bloque {bloque_id}: {ex}")
    finally:
        _sap_bloque_lock.release()
        _ac_disparados.discard(bloque_id)

    try:
        with _get_conn_local() as cn:
            cn.cursor().execute("""
                UPDATE itg.M5_AutoColor_BLOQUES
                SET estado='COMPLETADO', ejecutado_el=GETDATE(), ok_count=?, error_count=?
                WHERE id=?
            """, ok_n, err_n, bloque_id)
        print(f"[AUTOCOLOR] bloque {bloque_id} terminado — OK={ok_n} ERR={err_n}")
    except Exception as e:
        print(f"[AUTOCOLOR] error cerrando bloque {bloque_id}: {e}")


# ── API Routes ─────────────────────────────────────────────────────────────────

@bp.route("/zfers")
@_login_required
def api_ac_zfers():
    try:
        with _get_conn_local() as cn:
            rows = cn.cursor().execute("""
                SELECT id, zfer_base, estado, cargado_el FROM itg.M5_AutoColorZfer
                ORDER BY cargado_el DESC
            """).fetchall()

        resultado = []
        for r in rows:
            zfer_id, zfer_base, estado, cargado_el = r[0], r[1], r[2], r[3]
            colores = []
            try:
                colores, _ = _colores_disponibles_para_zfer(zfer_base)
                if colores is None:
                    colores = []
            except Exception:
                pass
            resultado.append({
                "id":         zfer_id,
                "zfer_base":  zfer_base,
                "estado":     estado,
                "cargado_el": cargado_el.strftime("%d/%m/%Y %H:%M") if cargado_el else "",
                "colores":    colores,
                "n_colores":  len(colores),
            })
        return jsonify({"ok": True, "zfers": resultado})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 200


@bp.route("/zfer", methods=["POST"])
@_login_required
def api_ac_agregar_zfer():
    body = request.get_json(force=True) or {}
    zfer = str(body.get("zfer", "")).strip()[:20]
    if not zfer:
        return jsonify({"ok": False, "error": "ZFER requerido"}), 200
    try:
        with _get_conn_local() as cn:
            cur = cn.cursor()
            existe = cur.execute("""
                SELECT TOP 1 id FROM itg.M5_AutoColorZfer
                WHERE zfer_base=? AND estado IN ('PENDIENTE','EN_COLA')
            """, zfer).fetchone()
            if existe:
                return jsonify({"ok": False, "error": f"{zfer} ya esta en la lista"}), 200
            cur.execute(
                "INSERT INTO itg.M5_AutoColorZfer (zfer_base, estado) VALUES (?, 'PENDIENTE')",
                zfer
            )
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 200


@bp.route("/zfer/<int:zfer_id>", methods=["DELETE"])
@_login_required
def api_ac_eliminar_zfer(zfer_id: int):
    try:
        with _get_conn_local() as cn:
            cn.cursor().execute("DELETE FROM itg.M5_AutoColorZfer WHERE id=?", zfer_id)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 200


@bp.route("/enviar/<int:zfer_id>", methods=["POST"])
@_login_required
def api_ac_enviar(zfer_id: int):
    """
    Genera items en M5_AutoColor_COLA.
    Recibe hr_por_color: {color_codigo: bool} — uno por color.
    El bloque se marca con cambiar_hr=1 si al menos un color lo requiere.
    """
    body        = request.get_json(force=True) or {}
    usuario     = str(body.get("usuario", "web"))[:50]
    hr_por_color = body.get("hr_por_color") or {}   # {"19": True, "21": False, ...}

    try:
        with _get_conn_local() as cn:
            row = cn.cursor().execute(
                "SELECT zfer_base FROM itg.M5_AutoColorZfer WHERE id=?", zfer_id
            ).fetchone()
        if not row:
            return jsonify({"ok": False, "error": "ZFER no encontrado"}), 200
        zfer_base = row[0]

        colores, attrs = _colores_disponibles_para_zfer(zfer_base)
        if colores is None:
            return jsonify({"ok": False, "error": f"No se pudo leer atributos: {attrs}"}), 200
        if not colores:
            return jsonify({"ok": False, "error": "No hay colores disponibles para homologar"}), 200

        shade_band = attrs.get("Z_SHADE_BAND", "00") or "00"
        partnumber = attrs.get("Z_AGP_PARTNUMBER", "")
        nivel      = attrs.get("Z_AGP_LEVEL", "")
        piece_type = attrs.get("Z_PIECE_TYPE", "")

        # El bloque lleva cambiar_hr=1 si al menos un color lo pide
        alguno_con_hr = any(hr_por_color.get(c["color_codigo"], False) for c in colores)

        with _get_conn_local() as cn:
            cur = cn.cursor()
            cur.execute("SELECT ISNULL(MAX(bloque_num),0)+1 FROM itg.M5_AutoColor_BLOQUES")
            nuevo_num = cur.fetchone()[0]
            cur.execute("""
                INSERT INTO itg.M5_AutoColor_BLOQUES (bloque_num, cambiar_hr)
                OUTPUT INSERTED.id, INSERTED.bloque_num
                VALUES (?, ?)
            """, nuevo_num, 1 if alguno_con_hr else 0)
            br = cur.fetchone()
            bloque_id, bloque_num = br[0], br[1]

            for it in colores:
                item_hr = 1 if hr_por_color.get(it["color_codigo"], False) else 0
                cur.execute("""
                    INSERT INTO itg.M5_AutoColor_COLA
                    (bloque_id, autocolor_zfer_id, zfer_base, color, color_nombre,
                     zpla, franja, pn_base, nivel, tipo_pieza, cambiar_hr)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?)
                """, bloque_id, zfer_id, zfer_base,
                    it["color_codigo"], it["color_nombre"], it["zpla"],
                    shade_band, partnumber[:50], nivel[:10], piece_type[:10],
                    item_hr)

            cur.execute(
                "UPDATE itg.M5_AutoColorZfer SET estado='EN_COLA' WHERE id=?", zfer_id
            )

        return jsonify({
            "ok": True, "n_agregados": len(colores),
            "bloque_id": bloque_id, "bloque_num": bloque_num,
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 200


@bp.route("/bloques")
@_login_required
def api_ac_bloques():
    try:
        with _get_conn_local() as cn:
            cur = cn.cursor()
            cur.execute("""
                SELECT b.id, b.bloque_num, b.estado, b.cambiar_hr,
                       b.creado_el, b.ejecutado_el, b.ok_count, b.error_count,
                       COUNT(c.id) AS total,
                       SUM(CASE WHEN c.estado IN ('PENDIENTE','EJECUTANDO') THEN 1 ELSE 0 END) AS pendientes
                FROM itg.M5_AutoColor_BLOQUES b
                LEFT JOIN itg.M5_AutoColor_COLA c ON c.bloque_id = b.id
                GROUP BY b.id, b.bloque_num, b.estado, b.cambiar_hr,
                         b.creado_el, b.ejecutado_el, b.ok_count, b.error_count
                ORDER BY b.id DESC
            """)
            bloques_rows = cur.fetchall()

        bloques = []
        for r in bloques_rows:
            bid = r[0]
            with _get_conn_local() as cn:
                items_rows = cn.cursor().execute("""
                    SELECT id, zfer_base, color, color_nombre, zpla,
                           estado, zfer_nuevo, error_msg, ejecutado_el
                    FROM itg.M5_AutoColor_COLA WHERE bloque_id=? ORDER BY id
                """, bid).fetchall()
            items = [{
                "id":           ir[0],
                "zfer_base":    ir[1],
                "color":        ir[2],
                "color_nombre": ir[3],
                "zpla":         ir[4],
                "estado":       ir[5],
                "zfer_nuevo":   ir[6] or "",
                "error_msg":    ir[7] or "",
                "ejecutado_el": ir[8].strftime("%H:%M:%S") if ir[8] else "",
            } for ir in items_rows]
            bloques.append({
                "id":           r[0],
                "bloque_num":   r[1],
                "estado":       r[2],
                "cambiar_hr":   bool(r[3]),
                "ok_count":     r[6],
                "error_count":  r[7],
                "total":        r[8],
                "pendientes":   r[9],
                "items":        items,
            })
        return jsonify({"ok": True, "bloques": bloques})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 200


@bp.route("/ejecutar/<int:bloque_id>", methods=["POST"])
@_login_required
def api_ac_ejecutar(bloque_id: int):
    body    = request.get_json(force=True) or {}
    usuario = str(body.get("usuario", "web"))[:50]
    try:
        with _get_conn_local() as cn:
            row = cn.cursor().execute(
                "SELECT estado FROM itg.M5_AutoColor_BLOQUES WHERE id=?", bloque_id
            ).fetchone()
        if not row:
            return jsonify({"ok": False, "error": "Bloque no encontrado"}), 200
        if row[0] == "EJECUTANDO":
            return jsonify({"ok": False, "error": "Ya en ejecucion"}), 200

        ocupado_por = _sap_lock_ocupado_por_otro()
        if ocupado_por:
            return jsonify({"ok": False, "error": f"SAP ocupado por '{ocupado_por}'"}), 200
        if not _sap_bloque_lock.acquire(blocking=False):
            return jsonify({"ok": False, "error": "Otro bloque esta usando SAP ahora mismo"}), 200
        _sap_bloque_lock.release()

        if bloque_id not in _ac_disparados:
            _ac_disparados.add(bloque_id)
            threading.Thread(
                target=_autocolor_ejecutar_bloque, args=(bloque_id, usuario), daemon=True
            ).start()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 200


@bp.route("/bloque/<int:bloque_id>/reporte")
@_login_required
def api_ac_reporte(bloque_id: int):
    try:
        with _get_conn_local() as cn:
            cur = cn.cursor()
            b = cur.execute("""
                SELECT bloque_num, estado, cambiar_hr, creado_el, ejecutado_el,
                       ok_count, error_count
                FROM itg.M5_AutoColor_BLOQUES WHERE id=?
            """, bloque_id).fetchone()
            if not b:
                return jsonify({"ok": False, "error": "Bloque no encontrado"}), 200
            items = cur.execute("""
                SELECT zfer_base, color, color_nombre, zpla, franja,
                       estado, zfer_nuevo, error_msg, ejecutado_el
                FROM itg.M5_AutoColor_COLA WHERE bloque_id=? ORDER BY id
            """, bloque_id).fetchall()

        duracion_seg = None
        if b[3] and b[4]:
            duracion_seg = int((b[4] - b[3]).total_seconds())

        return jsonify({
            "ok": True,
            "bloque": {
                "id":           bloque_id,
                "num":          b[0],
                "estado":       b[1],
                "cambiar_hr":   bool(b[2]),
                "creado_el":    b[3].strftime("%d/%m/%Y %H:%M:%S") if b[3] else "",
                "ejecutado_el": b[4].strftime("%d/%m/%Y %H:%M:%S") if b[4] else "",
                "ok_count":     b[5],
                "error_count":  b[6],
                "total":        len(items),
                "duracion_seg": duracion_seg,
            },
            "items": [{
                "zfer_base":    r[0],
                "color":        r[1],
                "color_nombre": r[2],
                "zpla":         r[3] or "",
                "franja":       r[4] or "",
                "estado":       r[5],
                "zfer_nuevo":   r[6] or "",
                "error_msg":    r[7] or "",
                "ejecutado_el": r[8].strftime("%H:%M:%S") if r[8] else "",
            } for r in items],
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 200


@bp.route("/bloque/<int:bloque_id>/reset", methods=["POST"])
@_login_required
def api_ac_reset_bloque(bloque_id: int):
    """Resetea un bloque colgado en EJECUTANDO → ERROR, items EJECUTANDO → PENDIENTE."""
    try:
        with _get_conn_local() as cn:
            cur = cn.cursor()
            cur.execute("""
                UPDATE itg.M5_AutoColor_BLOQUES SET estado='ERROR'
                WHERE id=? AND estado='EJECUTANDO'
            """, bloque_id)
            cur.execute("""
                UPDATE itg.M5_AutoColor_COLA SET estado='PENDIENTE'
                WHERE bloque_id=? AND estado='EJECUTANDO'
            """, bloque_id)
        _ac_disparados.discard(bloque_id)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 200


@bp.route("/bloque/<int:bloque_id>", methods=["DELETE"])
@_login_required
def api_ac_borrar_bloque(bloque_id: int):
    try:
        with _get_conn_local() as cn:
            cur = cn.cursor()
            row = cur.execute(
                "SELECT estado FROM itg.M5_AutoColor_BLOQUES WHERE id=?", bloque_id
            ).fetchone()
            if row and row[0] == "EJECUTANDO":
                return jsonify({"ok": False, "error": "No se puede borrar un bloque en ejecucion"}), 200
            cur.execute("DELETE FROM itg.M5_AutoColor_COLA WHERE bloque_id=?", bloque_id)
            cur.execute("DELETE FROM itg.M5_AutoColor_BLOQUES WHERE id=?", bloque_id)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 200
