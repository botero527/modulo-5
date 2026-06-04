"""
Script de migracion: copia datos de localhost\SQLEXPRESS\MODULO_5
a agpcolombia.database.windows.net\AGP_Ingenieria
Ejecutar UNA SOLA VEZ: py migrar_datos.py
"""
import pyodbc

LOCAL = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    r"SERVER=localhost\SQLEXPRESS;"
    "DATABASE=MODULO_5;"
    "Trusted_Connection=yes;"
)

NUEVO = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=agpcolombia.database.windows.net;"
    "DATABASE=AGP_Ingenieria;"
    "UID=DevIngenieria;"
    "PWD=HiJE068i0LQVrwA;"
    "Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
)

# Ya migradas OK: M5_Bloques, M5_LogEjecuciones, M5_RutasZFER, jobs_gestor_auto, bom_zfer_gestor_auto
# Solo ejecutar las que fallaron:
tablas = [
    (
        "dbo.M5_Bloques",
        "itg.M5_BLOQUES",
        "SELECT bloque_num, hora_prog, timer_activo, estado, ejecutado_el, ISNULL(ok_count,0), ISNULL(error_count,0) FROM dbo.M5_Bloques",
        "bloque_num, hora_prog, timer_activo, estado, ejecutado_el, ok_count, error_count"
    ),
    (
        "dbo.M5_Cola",
        "itg.M5_COLA",
        "SELECT bloque_id, zfer_base, tipo, color, color_nombre, zpla, franja, pn_base, nivel, tipo_pieza, formula_nueva, descripcion, acero_dir, ISNULL(cambiar_hr,0), zhal, subproducto, plano_manual, estado, zfer_nuevo, NULL, error_msg, ejecutado_el FROM dbo.M5_Cola",
        "bloque_id, zfer_base, tipo, color, color_nombre, zpla, franja, pn_base, nivel, tipo_pieza, formula_nueva, descripcion, acero_dir, cambiar_hr, zhal, subproducto, plano_manual, estado, zfer_nuevo, zfor_nuevo, error_msg, ejecutado_el"
    ),
    (
        "dbo.M5_LogEjecuciones",
        "itg.M5_LOGEJECUCIONES",
        "SELECT bloque_id, zfer_base, tipo, color, color_nombre, zpla, franja, pn_base, nivel, tipo_pieza, formula_nueva, acero_dir, zfer_nuevo, estado, error_msg, ejecutado_el, ejecutado_por FROM dbo.M5_LogEjecuciones",
        "bloque_id, zfer_base, tipo, color, color_nombre, zpla, franja, pn_base, nivel, tipo_pieza, formula_nueva, acero_dir, zfer_nuevo, estado, error_msg, ejecutado_el, ejecutado_por"
    ),
    (
        "dbo.M5_LogEjecucion",
        "itg.M5_LOGEJECUCION",
        "SELECT batch_id, NULL, zfer_nuevo, NULL, zpla, tipo_pieza, NULL, NULL, color_codigo, estado, detalle_error, fecha_inicio, fecha_fin FROM dbo.M5_LogEjecucion",
        "batch_id, zfer_base, zfer_nuevo, zfor_nuevo, zpla, tipo_pieza, formula, acero, color_codigo, estado, detalle_error, fecha_inicio, fecha_fin"
    ),
    (
        "dbo.M5_RutasZFER",
        "itg.M5_RUTASZFER",
        "SELECT zfer, ruta, descripcion, modificado_el, ISNULL(tiene_simetria,0), zfer_simetrico, pieza_contraria FROM dbo.M5_RutasZFER",
        "zfer, ruta, descripcion, modificado_el, tiene_simetria, zfer_simetrico, pieza_contraria"
    ),
    (
        "dbo.M5_Bloqueos",
        "itg.M5_BLOQUEOS",
        "SELECT pedido_origen, tipo_pieza, formula, acero_variante, color_codigo, motivo, bloqueado_por, NULL, ISNULL(activo,1) FROM dbo.M5_Bloqueos",
        "pedido_origen, tipo_pieza, formula, acero_variante, color_codigo, motivo, bloqueado_por, fecha_bloqueo, activo"
    ),
    (
        "dbo.jobs_gestor_auto",
        "itg.M5_JOBSGESTORAUTO",
        "SELECT id_origen, vehiculo_nombre, version_vehiculo, vehiculo_codigo, pieza, simetria, zfer_simetria, zfer, zfor, zpla, ruta_3dm FROM dbo.jobs_gestor_auto",
        "id_origen, vehiculo_nombre, version_vehiculo, vehiculo_codigo, pieza, simetria, zfer_simetria, zfer, zfor, zpla, ruta_3dm"
    ),
    (
        "dbo.bom_zfer_gestor_auto",
        "itg.M5_BOMGESTORAUTO",
        "SELECT zfer, posicion, clase, descripcion FROM dbo.bom_zfer_gestor_auto",
        "zfer, posicion, clase, descripcion"
    ),
]

def migrar():
    print("Conectando a servidores...")
    cn_local = pyodbc.connect(LOCAL, autocommit=True)
    cn_nuevo = pyodbc.connect(NUEVO, autocommit=True)
    print("Conexiones OK\n")

    for tabla_local, tabla_nueva, query_src, cols_dst in tablas:
        try:
            cur_local = cn_local.cursor()
            cur_local.execute(query_src)
            rows = cur_local.fetchall()
            print(f"{tabla_local} -> {tabla_nueva}: {len(rows)} filas")

            if not rows:
                print(f"  (vacia, se omite)")
                continue

            n_cols = len(cols_dst.split(","))
            ph = ",".join(["?"] * n_cols)
            cur_nuevo = cn_nuevo.cursor()
            cur_nuevo.fast_executemany = True
            cur_nuevo.executemany(
                f"INSERT INTO {tabla_nueva} ({cols_dst}) VALUES ({ph})",
                [tuple(r) for r in rows]
            )
            print(f"  OK - {len(rows)} filas migradas")
        except Exception as e:
            print(f"  ERROR en {tabla_local}: {e}")

    cn_local.close()
    cn_nuevo.close()
    print("\nMigracion completada.")

if __name__ == "__main__":
    migrar()
