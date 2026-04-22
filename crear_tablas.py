import pydoc
import pyodbc
import sys
#aqui meto la configuracion del servidor pa cualquier cosa que cambien la direccion o algo es solo cambiar la direccion de esta chimbada
DB_LOCAL = {
    "server": r"localhost\SQLEXPRESS",
    "database": "MODULO_5",
    "driver": "ODBC Driver 17 for SQL server",
}

CONNECTION_STRING = (
    f"DRIVER={{{DB_LOCAL['driver']}}};"
    f"SERVER={DB_LOCAL['server']};"
    f"DATABASE={DB_LOCAL['database']};"
    "Trusted_Connection=yes;"
)


#SENTENCIAS SQL

SQL_EXISTE_TABLA = """
SELECT COUNT(1)
FROM sys.tables
WHERE name = ?
 AND schema_id = SCHEMA_ID('dbo') 
"""

#aqui la primera tabla la m5_Bloqueos

SQL_CREAR_BLOQUEOS = """
CREATE TABLE dbo.M5_Bloqueos (
    id              INT IDENTITY(1,1)   NOT NULL,
    pedido_origen   NVARCHAR(50)        NOT NULL,
    tipo_pieza      NVARCHAR(20)        NOT NULL,
    formula         NVARCHAR(20)        NOT NULL,
    acero_variante  NVARCHAR(5)         NOT NULL,
    color_codigo    NVARCHAR(50)        NOT NULL DEFAULT '',
    motivo          NVARCHAR(500)       NOT NULL,
    bloqueado_por   NVARCHAR(100)       NOT NULL,
    fecha_bloqueo   DATETIME            NOT NULL DEFAULT GETDATE(),
    activo          BIT                 NOT NULL DEFAULT 1,
    CONSTRAINT PK_M5_Bloqueos PRIMARY KEY CLUSTERED (id ASC)
)
"""

# Migración idempotente: agrega color_codigo si la tabla ya existe sin ella
SQL_MIGRAR_COLOR_CODIGO = """
IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME = 'M5_Bloqueos' AND COLUMN_NAME = 'color_codigo'
)
BEGIN
    ALTER TABLE dbo.M5_Bloqueos
    ADD color_codigo NVARCHAR(50) NOT NULL DEFAULT ''
END
"""

SQL_INDICE_BLOQUEOS = """
CREATE NONCLUSTERED INDEX IX_M5_Bloqueos_Pedido_Activo
    ON dbo.M5_Bloqueos (pedido_origen, tipo_pieza, activo)
"""

#la otra tabla la M5_LogEjecucion
SQL_CREAR_LOG = """
CREATE TABLE dbo.M5_LogEjecucion (
    id              INT IDENTITY(1,1)   NOT NULL,
    batch_id        UNIQUEIDENTIFIER    NOT NULL,
    pedido_origen   NVARCHAR(50)        NOT NULL,
    tipo_pieza      NVARCHAR(100)      NOT NULL,
    formula         NVARCHAR(20)        NOT NULL,
    color_codigo    NVARCHAR(50)        NOT NULL,
    acero_variante  NVARCHAR(5)         NOT NULL,
    estado          NVARCHAR(20)        NOT NULL,
    detalle_error   NVARCHAR(MAX)       NULL,
    fecha_inicio    DATETIME            NOT NULL,
    fecha_fin       DATETIME            NULL,
    CONSTRAINT PK_M5_LogEjecucion PRIMARY KEY CLUSTERED (id ASC)
)
"""

SQL_INDICE_LOG_BATCH = """
CREATE NONCLUSTERED INDEX IX_M5_LogEjecucion_Batch_Estado
    ON dbo.M5_LogEjecucion (batch_id, estado)
"""

SQL_INDICE_LOG_PEDIDO = """
CREATE NONCLUSTERED INDEX IX_M5_LogEjecucion_Pedido
    ON dbo.M5_LogEjecucion (pedido_origen, fecha_inicio DESC)
"""

#FUNCIONESSSSSSSSSSSS

def tabla_existe(cursor, nombre_tabla:str) -> bool:
    """Devuelve true si la tabla ua exitewee"""
    cursor.execute(SQL_EXISTE_TABLA,(nombre_tabla,))
    return cursor.fetchone()[0] == 1

def crear_tabla_bloqueos(cursor) -> None:
    nombre = "M5_Bloqueos"
    if tabla_existe(cursor, nombre):
        print(f"  [OK] {nombre} ya existe — verificando migraciones...")
        cursor.execute(SQL_MIGRAR_COLOR_CODIGO)
        print(f"  [OK] Migracion color_codigo aplicada (o ya existia).")
        return
    print(f"  [→]  Creando {nombre}...")
    cursor.execute(SQL_CREAR_BLOQUEOS)
    print(f"  [→]  Creando índice en {nombre}...")
    cursor.execute(SQL_INDICE_BLOQUEOS)
    print(f"  [✓]  {nombre} creada.")


def crear_tabla_log(cursor) -> None:
    nombre = "M5_LogEjecucion"
    if tabla_existe(cursor, nombre):
        print(f"  [OK] {nombre} ya existe — sin cambios.")
        return
    print(f"  [→]  Creando {nombre}...")
    cursor.execute(SQL_CREAR_LOG)
    print(f"  [→]  Creando índices en {nombre}...")
    cursor.execute(SQL_INDICE_LOG_BATCH)
    cursor.execute(SQL_INDICE_LOG_PEDIDO)
    print(f"  [✓]  {nombre} creada.")


def verificar_columnas(cursor, nombre_tabla: str) -> None:
    """Imprime las columnas de la tabla para verificación visual."""
    cursor.execute("""
        SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE
        FROM   INFORMATION_SCHEMA.COLUMNS
        WHERE  TABLE_NAME = ?
        ORDER  BY ORDINAL_POSITION
    """, (nombre_tabla,))
    filas = cursor.fetchall()
    print(f"\n  Columnas de {nombre_tabla}:")
    print(f"  {'Columna':<25} {'Tipo':<20} {'Longitud':<10} {'Nulable'}")
    print(f"  {'-'*65}")
    for col, tipo, longitud, nulable in filas:
        print(f"  {col:<25} {tipo:<20} {str(longitud) if longitud else '—':<10} {nulable}")


def inicializar_tablas_modulo5() -> bool:
    """
    Conecta a la BD local y crea las tablas M5_Bloqueos y M5_LogEjecucion.

    Returns:
        True si todo salió bien, False si hubo error.
    """
    print("\n" + "="*60)
    print("  MÓDULO 5 — Inicialización de tablas (BD LOCAL)")
    print(f"  Servidor     : {DB_LOCAL['server']}")
    print(f"  Base de datos: {DB_LOCAL['database']}")
    print("="*60)

    try:
        print("\n  Conectando...")
        conn = pyodbc.connect(CONNECTION_STRING, autocommit=False)
        cursor = conn.cursor()
        print("  Conexión exitosa.\n")

        crear_tabla_bloqueos(cursor)
        crear_tabla_log(cursor)

        conn.commit()
        print("\n  Cambios confirmados (commit).")

        # Verificación visual de las columnas creadas
        verificar_columnas(cursor, "M5_Bloqueos")
        verificar_columnas(cursor, "M5_LogEjecucion")

        cursor.close()
        conn.close()

        print("\n" + "="*60)
        print("  ✓ Inicialización completada exitosamente.")
        print("="*60 + "\n")
        return True

    except pyodbc.Error as e:
        print(f"\n  [ERROR] Error de base de datos: {e}")
        print("\n  Posibles causas:")
        print(f"  → El servidor '{DB_LOCAL['server']}' no está corriendo.")
        print(f"     Verifica en SSMS que puedes conectarte con Windows Auth.")
        print(f"  → La BD '{DB_LOCAL['database']}' no existe en local.")
        print(f"     Créala en SSMS: clic derecho en Databases → New Database.")
        print(f"  → El ODBC Driver 17 no está instalado.")
        print(f"     Descárgalo de: aka.ms/odbc17")
        return False

    except Exception as e:
        print(f"\n  [ERROR] Error inesperado: {e}")
        return False


if __name__ == "__main__":
    exito = inicializar_tablas_modulo5()
    sys.exit(0 if exito else 1)