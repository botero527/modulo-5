-- ============================================================
-- PASO 1: Crear schema + tablas en AGP_Ingenieria
-- Ejecutar en: agpcolombia.database.windows.net / AGP_Ingenieria
-- Usuario: DevIngenieria
-- ============================================================

-- 1. Crear schema
IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = 'itg')
    EXEC('CREATE SCHEMA itg');
GO

-- ============================================================
-- 2. M5_BLOQUES — Bloques de la cola de homologación
-- ============================================================
CREATE TABLE itg.M5_BLOQUES (
    id              INT IDENTITY(1,1) PRIMARY KEY,
    bloque_num      INT NOT NULL DEFAULT 1,
    hora_prog       DATETIME NULL,
    timer_activo    BIT NOT NULL DEFAULT 1,
    estado          VARCHAR(20) NOT NULL DEFAULT 'PENDIENTE',
    ejecutado_el    DATETIME NULL,
    ok_count        INT NULL DEFAULT 0,
    error_count     INT NULL DEFAULT 0
);
GO

-- ============================================================
-- 3. M5_COLA — Items pendientes de homologación
-- ============================================================
CREATE TABLE itg.M5_COLA (
    id              INT IDENTITY(1,1) PRIMARY KEY,
    bloque_id       INT NOT NULL,
    zfer_base       NVARCHAR(20)  NOT NULL,
    tipo            NVARCHAR(20)  NOT NULL DEFAULT 'COLOR',
    color           NVARCHAR(10)  NULL,
    color_nombre    NVARCHAR(100) NULL,
    zpla            NVARCHAR(20)  NULL,
    franja          NVARCHAR(5)   NULL DEFAULT '00',
    pn_base         NVARCHAR(50)  NULL,
    nivel           NVARCHAR(10)  NULL,
    tipo_pieza      NVARCHAR(10)  NULL,
    formula_nueva   NVARCHAR(30)  NULL,
    descripcion     NVARCHAR(200) NULL,
    acero_dir       NVARCHAR(10)  NULL,
    cambiar_hr      BIT NOT NULL DEFAULT 0,
    zhal            NVARCHAR(20)  NULL,
    subproducto     NVARCHAR(20)  NULL,
    plano_manual    NVARCHAR(100) NULL,
    estado          NVARCHAR(20)  NOT NULL DEFAULT 'PENDIENTE',
    zfer_nuevo      NVARCHAR(20)  NULL,
    zfor_nuevo      NVARCHAR(50)  NULL,
    error_msg       NVARCHAR(500) NULL,
    ejecutado_el    DATETIME NULL,
    CONSTRAINT FK_COLA_BLOQUE FOREIGN KEY (bloque_id) REFERENCES itg.M5_BLOQUES(id)
);
GO

CREATE INDEX IX_M5_COLA_BLOQUE ON itg.M5_COLA (bloque_id, estado);
GO

-- ============================================================
-- 4. M5_LOGEJECUCIONES — Historial permanente de homologaciones
-- ============================================================
CREATE TABLE itg.M5_LOGEJECUCIONES (
    id              INT IDENTITY(1,1) PRIMARY KEY,
    bloque_id       INT NULL,
    zfer_base       NVARCHAR(20)  NULL,
    tipo            NVARCHAR(20)  NULL,
    color           NVARCHAR(10)  NULL,
    color_nombre    NVARCHAR(100) NULL,
    zpla            NVARCHAR(20)  NULL,
    franja          NVARCHAR(5)   NULL,
    pn_base         NVARCHAR(50)  NULL,
    nivel           NVARCHAR(10)  NULL,
    tipo_pieza      NVARCHAR(10)  NULL,
    formula_nueva   NVARCHAR(30)  NULL,
    acero_dir       NVARCHAR(10)  NULL,
    zfer_nuevo      NVARCHAR(20)  NULL,
    estado          NVARCHAR(20)  NULL,
    error_msg       NVARCHAR(500) NULL,
    ejecutado_el    DATETIME NULL DEFAULT GETDATE(),
    ejecutado_por   NVARCHAR(100) NULL
);
GO

CREATE INDEX IX_M5_LOGEJECUCIONES_FECHA ON itg.M5_LOGEJECUCIONES (ejecutado_el DESC);
CREATE INDEX IX_M5_LOGEJECUCIONES_TIPO  ON itg.M5_LOGEJECUCIONES (tipo, estado);
GO

-- ============================================================
-- 5. M5_LOGEJECUCION — Log de ejecuciones (versión legada db_local.py)
-- ============================================================
CREATE TABLE itg.M5_LOGEJECUCION (
    id              INT IDENTITY(1,1) PRIMARY KEY,
    batch_id        VARCHAR(50)   NULL,
    zfer_base       NVARCHAR(20)  NULL,
    zfer_nuevo      NVARCHAR(20)  NULL,
    zfor_nuevo      NVARCHAR(50)  NULL,
    zpla            NVARCHAR(20)  NULL,
    tipo_pieza      NVARCHAR(10)  NULL,
    formula         NVARCHAR(30)  NULL,
    acero           NVARCHAR(30)  NULL,
    color_codigo    NVARCHAR(10)  NULL,
    estado          VARCHAR(10)   NULL,
    detalle_error   NVARCHAR(500) NULL,
    fecha_inicio    DATETIME NULL,
    fecha_fin       DATETIME NULL
);
GO

-- ============================================================
-- 6. M5_RUTASZFER — Rutas de archivos 3D por ZFER
-- ============================================================
CREATE TABLE itg.M5_RUTASZFER (
    zfer            NVARCHAR(20)  NOT NULL PRIMARY KEY,
    ruta            NVARCHAR(500) NULL,
    descripcion     NVARCHAR(200) NULL,
    modificado_el   DATETIME NULL DEFAULT GETDATE(),
    tiene_simetria  BIT NOT NULL DEFAULT 0,
    zfer_simetrico  NVARCHAR(20)  NULL,
    pieza_contraria NVARCHAR(10)  NULL
);
GO

-- ============================================================
-- 7. M5_BLOQUEOS — Combinaciones bloqueadas manualmente
-- ============================================================
CREATE TABLE itg.M5_BLOQUEOS (
    id              INT IDENTITY(1,1) PRIMARY KEY,
    pedido_origen   NVARCHAR(20)  NULL,
    tipo_pieza      NVARCHAR(10)  NULL,
    formula         NVARCHAR(30)  NULL,
    acero_variante  NVARCHAR(30)  NULL,
    color_codigo    NVARCHAR(10)  NULL,
    motivo          NVARCHAR(200) NULL,
    bloqueado_por   NVARCHAR(100) NULL,
    fecha_bloqueo   DATETIME NULL DEFAULT GETDATE(),
    activo          BIT NOT NULL DEFAULT 1
);
GO

-- ============================================================
-- 8. M5_JOBSGESTORAUTO — Jobs para el gestor automático
-- ============================================================
CREATE TABLE itg.M5_JOBSGESTORAUTO (
    id_origen           INT NOT NULL PRIMARY KEY,
    vehiculo_nombre     VARCHAR(150)  NOT NULL,
    version_vehiculo    VARCHAR(80)   NOT NULL,
    vehiculo_codigo     VARCHAR(20)   NOT NULL,
    pieza               VARCHAR(3)    NOT NULL,
    simetria            VARCHAR(2)    NOT NULL,
    zfer_simetria       VARCHAR(20)   NULL,
    zfer                VARCHAR(20)   NOT NULL,
    zfor                VARCHAR(50)   NULL,
    zpla                VARCHAR(50)   NULL,
    ruta_3dm            VARCHAR(500)  NOT NULL,
    CONSTRAINT CK_JOBS_SIMETRIA CHECK (simetria IN ('SI','NO')),
    CONSTRAINT CK_JOBS_PIEZA    CHECK (LEN(pieza) = 3),
    CONSTRAINT CK_JOBS_ZFER_SIM CHECK (
        (simetria='NO') OR
        (simetria='SI' AND zfer_simetria IS NOT NULL AND LEN(LTRIM(RTRIM(zfer_simetria)))>0)
    )
);
GO

-- ============================================================
-- 9. M5_BOMGESTORAUTO — BOM del gestor automático
-- ============================================================
CREATE TABLE itg.M5_BOMGESTORAUTO (
    id_bom      INT IDENTITY(1,1) PRIMARY KEY,
    zfer        VARCHAR(20)   NOT NULL,
    posicion    VARCHAR(10)   NOT NULL,
    clase       VARCHAR(100)  NOT NULL,
    descripcion VARCHAR(200)  NULL
);
GO

PRINT 'Schema itg y todas las tablas creadas correctamente en AGP_Ingenieria';
GO
