-- ============================================================
-- PASO 3: Verificar migración en AGP_Ingenieria
-- Ejecutar en: agpcolombia.database.windows.net / AGP_Ingenieria
-- ============================================================

SELECT 'M5_BLOQUES'       AS tabla, COUNT(*) AS filas FROM itg.M5_BLOQUES       UNION ALL
SELECT 'M5_COLA'          AS tabla, COUNT(*) AS filas FROM itg.M5_COLA           UNION ALL
SELECT 'M5_LOGEJECUCIONES'AS tabla, COUNT(*) AS filas FROM itg.M5_LOGEJECUCIONES UNION ALL
SELECT 'M5_LOGEJECUCION'  AS tabla, COUNT(*) AS filas FROM itg.M5_LOGEJECUCION   UNION ALL
SELECT 'M5_RUTASZFER'     AS tabla, COUNT(*) AS filas FROM itg.M5_RUTASZFER       UNION ALL
SELECT 'M5_BLOQUEOS'      AS tabla, COUNT(*) AS filas FROM itg.M5_BLOQUEOS        UNION ALL
SELECT 'M5_JOBSGESTORAUTO'AS tabla, COUNT(*) AS filas FROM itg.M5_JOBSGESTORAUTO  UNION ALL
SELECT 'M5_BOMGESTORAUTO' AS tabla, COUNT(*) AS filas FROM itg.M5_BOMGESTORAUTO;
GO
