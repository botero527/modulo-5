-- ============================================================
-- PASO 2: Migrar datos de local a AGP_Ingenieria
-- Ejecutar en: localhost\SQLEXPRESS (servidor LOCAL)
-- con linked server o desde SSMS conectado al LOCAL
-- ============================================================
-- NOTA: Ajusta [AGPCOLOMBIA] si el linked server tiene otro nombre
-- O puedes exportar/importar con SSMS "Import/Export Data"
-- ============================================================

-- Si usas SSMS "Import Data" (recomendado para no tener linked server):
-- 1. Click derecho en AGP_Ingenieria → Tasks → Import Data
-- 2. Source: SQL Server Native Client / localhost\SQLEXPRESS / MODULO_5
-- 3. Destination: SQL Server Native Client / agpcolombia.database.windows.net / AGP_Ingenieria
-- 4. Mapear cada tabla local → nueva tabla itg.*
-- ============================================================

-- ALTERNATIVA: Script para ejecutar DESDE el servidor local
-- si tienes linked server apuntando a agpcolombia

-- M5_BLOQUES
INSERT INTO [agpcolombia].[AGP_Ingenieria].[itg].[M5_BLOQUES]
    (bloque_num, hora_prog, timer_activo, estado, ejecutado_el, ok_count, error_count)
SELECT bloque_num, hora_prog, timer_activo, estado, ejecutado_el,
       ISNULL(ok_count,0), ISNULL(error_count,0)
FROM [MODULO_5].[dbo].[M5_Bloques];
GO

-- M5_COLA (sin FK violation — bloques deben existir primero)
INSERT INTO [agpcolombia].[AGP_Ingenieria].[itg].[M5_COLA]
    (bloque_id, zfer_base, tipo, color, color_nombre, zpla, franja, pn_base, nivel,
     tipo_pieza, formula_nueva, descripcion, acero_dir, cambiar_hr, zhal, subproducto,
     plano_manual, estado, zfer_nuevo, zfor_nuevo, error_msg, ejecutado_el)
SELECT bloque_id, zfer_base, tipo, color, color_nombre, zpla, franja, pn_base, nivel,
       tipo_pieza, formula_nueva, descripcion, acero_dir, ISNULL(cambiar_hr,0), zhal, subproducto,
       plano_manual, estado, zfer_nuevo, zfor_nuevo, error_msg, ejecutado_el
FROM [MODULO_5].[dbo].[M5_Cola];
GO

-- M5_LOGEJECUCIONES
INSERT INTO [agpcolombia].[AGP_Ingenieria].[itg].[M5_LOGEJECUCIONES]
    (bloque_id, zfer_base, tipo, color, color_nombre, zpla, franja, pn_base, nivel,
     tipo_pieza, formula_nueva, acero_dir, zfer_nuevo, estado, error_msg, ejecutado_el, ejecutado_por)
SELECT bloque_id, zfer_base, tipo, color, color_nombre, zpla, franja, pn_base, nivel,
       tipo_pieza, formula_nueva, acero_dir, zfer_nuevo, estado, error_msg, ejecutado_el, ejecutado_por
FROM [MODULO_5].[dbo].[M5_LogEjecuciones];
GO

-- M5_LOGEJECUCION (legado)
INSERT INTO [agpcolombia].[AGP_Ingenieria].[itg].[M5_LOGEJECUCION]
    (batch_id, zfer_base, zfer_nuevo, zfor_nuevo, zpla, tipo_pieza, formula, acero,
     color_codigo, estado, detalle_error, fecha_inicio, fecha_fin)
SELECT batch_id, zfer_base, zfer_nuevo, zfor_nuevo, zpla, tipo_pieza, formula, acero,
       color_codigo, estado, detalle_error, fecha_inicio, fecha_fin
FROM [MODULO_5].[dbo].[M5_LogEjecucion];
GO

-- M5_RUTASZFER
INSERT INTO [agpcolombia].[AGP_Ingenieria].[itg].[M5_RUTASZFER]
    (zfer, ruta, descripcion, modificado_el, tiene_simetria, zfer_simetrico, pieza_contraria)
SELECT zfer, ruta, descripcion, modificado_el, ISNULL(tiene_simetria,0), zfer_simetrico, pieza_contraria
FROM [MODULO_5].[dbo].[M5_RutasZFER];
GO

-- M5_BLOQUEOS
INSERT INTO [agpcolombia].[AGP_Ingenieria].[itg].[M5_BLOQUEOS]
    (pedido_origen, tipo_pieza, formula, acero_variante, color_codigo, motivo,
     bloqueado_por, fecha_bloqueo, activo)
SELECT pedido_origen, tipo_pieza, formula, acero_variante, color_codigo, motivo,
       bloqueado_por, fecha_bloqueo, ISNULL(activo,1)
FROM [MODULO_5].[dbo].[M5_Bloqueos];
GO

-- M5_JOBSGESTORAUTO
INSERT INTO [agpcolombia].[AGP_Ingenieria].[itg].[M5_JOBSGESTORAUTO]
    (id_origen, vehiculo_nombre, version_vehiculo, vehiculo_codigo, pieza, simetria,
     zfer_simetria, zfer, zfor, zpla, ruta_3dm)
SELECT id_origen, vehiculo_nombre, version_vehiculo, vehiculo_codigo, pieza, simetria,
       zfer_simetria, zfer, zfor, zpla, ruta_3dm
FROM [MODULO_5].[dbo].[jobs_gestor_auto];
GO

-- M5_BOMGESTORAUTO
INSERT INTO [agpcolombia].[AGP_Ingenieria].[itg].[M5_BOMGESTORAUTO]
    (zfer, posicion, clase, descripcion)
SELECT zfer, posicion, clase, descripcion
FROM [MODULO_5].[dbo].[bom_zfer_gestor_auto];
GO

PRINT 'Migración de datos completada';
GO
