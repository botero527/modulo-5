"""
db_local.py — Helper para BD local SQL Server Express
Maneja bloqueos (M5_Bloqueos) y log de ejecuciones (M5_LogEjecucion).
"""
import pyodbc
from datetime import datetime
from typing import Optional

_CONN_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    r"SERVER=localhost\SQLEXPRESS;"
    "DATABASE=MODULO_5;"
    "Trusted_Connection=yes;"
)

def _conn():
    return pyodbc.connect(_CONN_STR, autocommit=True)


# ── Bloqueos ───────────────────────────────────────────────────────────────────

def bloqueo_existe(zfer: str, color_codigo: str) -> Optional[dict]:
    """Devuelve el bloqueo activo para zfer+color, o None."""
    try:
        cn = _conn()
        cur = cn.cursor()
        cur.execute(
            "SELECT id, motivo, bloqueado_por, fecha_bloqueo "
            "FROM dbo.M5_Bloqueos "
            "WHERE pedido_origen=? AND color_codigo=? AND activo=1",
            (zfer, color_codigo)
        )
        row = cur.fetchone()
        cn.close()
        if row:
            return {"id": row[0], "motivo": row[1], "bloqueado_por": row[2], "fecha": str(row[3])}
        return None
    except Exception:
        return None


def bloqueos_para_zfer(zfer: str) -> list:
    """Devuelve todos los bloqueos activos para un ZFER base."""
    try:
        cn = _conn()
        cur = cn.cursor()
        cur.execute(
            "SELECT id, color_codigo, motivo, bloqueado_por, fecha_bloqueo "
            "FROM dbo.M5_Bloqueos WHERE pedido_origen=? AND activo=1 ORDER BY fecha_bloqueo DESC",
            (zfer,)
        )
        rows = cur.fetchall()
        cn.close()
        return [{"id": r[0], "color": r[1], "motivo": r[2], "por": r[3], "fecha": str(r[4])} for r in rows]
    except Exception:
        return []


def bloquear(zfer: str, color_codigo: str, formula: str, tipo_pieza: str,
             acero_variante: str, motivo: str, bloqueado_por: str = "web") -> bool:
    """Inserta o reactiva un bloqueo. Devuelve True si OK."""
    try:
        cn = _conn()
        cur = cn.cursor()
        # Verificar si ya existe (inactivo) para reactivar
        cur.execute(
            "SELECT id FROM dbo.M5_Bloqueos WHERE pedido_origen=? AND color_codigo=?",
            (zfer, color_codigo)
        )
        row = cur.fetchone()
        if row:
            cur.execute(
                "UPDATE dbo.M5_Bloqueos SET activo=1, motivo=?, bloqueado_por=?, fecha_bloqueo=GETDATE() WHERE id=?",
                (motivo, bloqueado_por, row[0])
            )
        else:
            cur.execute(
                "INSERT INTO dbo.M5_Bloqueos (pedido_origen, tipo_pieza, formula, acero_variante, color_codigo, motivo, bloqueado_por) "
                "VALUES (?, ?, ?, ?, ?, ?, ?)",
                (zfer, tipo_pieza, formula, acero_variante, color_codigo, motivo, bloqueado_por)
            )
        cn.close()
        return True
    except Exception as e:
        print(f"[db_local] Error bloquear: {e}")
        return False


def desbloquear(zfer: str, color_codigo: str) -> bool:
    """Desactiva un bloqueo."""
    try:
        cn = _conn()
        cur = cn.cursor()
        cur.execute(
            "UPDATE dbo.M5_Bloqueos SET activo=0 WHERE pedido_origen=? AND color_codigo=? AND activo=1",
            (zfer, color_codigo)
        )
        cn.close()
        return True
    except Exception as e:
        print(f"[db_local] Error desbloquear: {e}")
        return False


# ── Log de ejecuciones ─────────────────────────────────────────────────────────

def log_inicio(batch_id: str, zfer: str, tipo_pieza: str, formula: str,
               color_codigo: str, acero: str = "") -> Optional[int]:
    try:
        cn = _conn()
        cur = cn.cursor()
        cur.execute(
            "INSERT INTO dbo.M5_LogEjecucion "
            "(batch_id, pedido_origen, tipo_pieza, formula, color_codigo, acero_variante, estado, fecha_inicio) "
            "OUTPUT INSERTED.id "
            "VALUES (?, ?, ?, ?, ?, ?, 'EN_PROCESO', GETDATE())",
            (batch_id, zfer, tipo_pieza, formula, color_codigo, acero)
        )
        row = cur.fetchone()
        cn.close()
        return row[0] if row else None
    except Exception as e:
        print(f"[db_local] Error log_inicio: {e}")
        return None


def log_fin(log_id: int, estado: str, detalle: str = "") -> None:
    try:
        cn = _conn()
        cur = cn.cursor()
        cur.execute(
            "UPDATE dbo.M5_LogEjecucion SET estado=?, detalle_error=?, fecha_fin=GETDATE() WHERE id=?",
            (estado, detalle or None, log_id)
        )
        cn.close()
    except Exception as e:
        print(f"[db_local] Error log_fin: {e}")
