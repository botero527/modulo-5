# MODULO 5 — AGP Glass: Generador de Combinaciones + Automatizacion SAP
## Documentacion tecnica para agentes y desarrolladores

---

## 1. RESUMEN DEL PROYECTO

**Objetivo:** Automatizar la creacion masiva de variantes de vidrios blindados en SAP.
Dado un ZFER base (pieza de vidrio), el sistema genera todas las combinaciones posibles
de formula × acero × color, permite bloquear combinaciones no deseadas, y luego ejecuta
el proceso de homologacion en SAP (transaccion ZMME0001 "Cambio de color") para cada
combinacion activa.

**Empresa:** AGP Glass — Colombia (planta CO01)
**Entorno SAP de pruebas (QUAS):** usuario `PROGRAING`, password `AGPcol123*`
**BD Produccion:** `agpcol.database.windows.net` / `agpc-productivity` / user `Consulta` / pwd `@GPgl4$$2021`

---

## 2. ARQUITECTURA GENERAL

```
MODULO5.py          ← App principal Tkinter (3 pestanas)
│
├── Pestana 1: COMBINACIONES
│     Ejecuta COMBINADOR.py como subprocess
│     Lee ZFER base → consulta BD produccion → genera combinaciones5.xlsx
│
├── Pestana 2: BLOQUEOS
│     Carga VISTAAAA.py (VistaPreviaBloqueos)
│     Permite marcar combinaciones como bloqueadas o pendientes
│     Solo las "activas" (no bloqueadas, no pendientes) van a SAP
│
└── Pestana 3: SAP AUTOMATIZACION
      Filtra items activos por formula base (busca en BD)
      Lanza hilo → AutomatizadorSAP (SAP_AUTOMATIZADOR.py)
      Genera reporte Excel + JSON checkpoint por item
```

---

## 3. ARCHIVOS DEL PROYECTO

| Archivo | Rol |
|---|---|
| `MODULO5.py` | App principal unificada (Tkinter 3 tabs) |
| `SAP_AUTOMATIZADOR.py` | Motor de automatizacion SAP via GUI Scripting |
| `COMBINADOR.py` | Genera combinaciones formula×acero×color desde BD |
| `VISTAAAA.py` | Vista previa de bloqueos (tabla interactiva Tkinter) |
| `combinaciones5.xlsx` | Output del COMBINADOR (input de Bloqueos y SAP) |
| `reporte_sap_YYYYMMDD_HHMMSS.xlsx` | Reporte final por batch (4 hojas) |
| `progreso_XXXXXXXX.json` | Checkpoint JSON actualizado despues de cada item |

---

## 4. FLUJO COMPLETO PASO A PASO

### Pestaña 1 — Generar Combinaciones
1. Usuario ingresa ZFER base (ej: `700179044`)
2. MODULO5.py lanza `COMBINADOR.py` como subprocess con env var `M5_ZFER_BASE`
3. COMBINADOR consulta BD produccion: busca todas las formulas, aceros y colores
   validos para ese ZFER
4. Genera `combinaciones5.xlsx` con columnas: `zfer_origen`, `formula`, `acero`,
   `color`, `mercado`, `cod_pieza`, `tipo_pieza`, etc.
5. El mercado se detecta automaticamente desde la BD (no hay selector manual)

### Pestaña 2 — Bloqueos
1. Carga `combinaciones5.xlsx` en tabla interactiva
2. Usuario puede marcar combinaciones como BLOQUEADA o PENDIENTE
3. Las "activas" (sin marcar) son las que van a SAP
4. Boton "Ir a SAP" habilita la pestaña 3

### Pestaña 3 — Automatizacion SAP
Proceso completo:

**Pre-vuelo (hilo UI):**
1. Obtiene items activos de VistaPreviaBloqueos
2. Consulta BD produccion para formula del ZFER base:
   - Primero: `ZFER_Characteristics_Genesis` WHERE `SpecID = ZFER` → columna `FormulaCode`
   - Si no: `TCAL_CALENDARIO_COLOMBIA_DIRECT` WHERE `ZFER = ZFER` → columna `Formula`
   - Si no encuentra en ninguna: **BLOQUEA** y muestra error al usuario
3. Filtra: misma formula que base → `items_sap` (van a SAP)
4. Filtra: formula diferente → `items_solo` (solo reporte, pendiente cambio formula)
5. Muestra dialogo de confirmacion con numeros reales
6. Lanza hilo worker

**Hilo worker — por cada combinacion en `items_sap`:**

```
PASO 1: leer_clasificacion_zfer(zfer_base)
        → MM02 → tab Clasificacion → tab PIEZA
        → lee PARTNUMBER, COLOR, FRANJA (codigo SAP: "00","01","02","03","NA")
        → p_franj = clasif.franja if clasif.franja else "00"

PASO 2: zmme0001_ejecutar(zfer_base, p_color, p_franj)
        → Navega a ZMME0001
        → Selecciona Homologar (radRB5)
        → Llena P_MATER-LOW = zfer_base
        → Llena P_CENTER = CO01
        → Selecciona Cambio de Color (radRB3_A1)
        → Llena P_COLOR = numero del color (ej: "19")
        → Llena P_FRANJ = codigo franja (ej: "00")
        → F4 en P_ZPLA → selecciona fila 0 → lee ZPLA seleccionado
        → F8 (Ejecutar)
        → Lee grid resultado: ZFER_NUEVO, ZFOR_NUEVO
        → Retorna (zfer_nuevo, zfor_nuevo, zpla)

PASO 3: zppr0020_esperar_fases(zfer_nuevo)
        → Navega a ZPPR0020
        → Llena Mod.por = PROGRAING, Centro = CO01
        → F8 (Ejecutar)
        → Polling cada 30 segundos, maximo 10 minutos:
            - Lee grid ALV buscando fila donde ZFER = zfer_nuevo
            - Si alguna fase tiene "E" → error, aborta
            - Si fase 8+ tiene "S" → OK, continua
            - F9 para refrescar entre intentos
        → Retorna {ok, zpla, fase_error, detalle, fases}

PASO 4: Volver a ZMME0001 con ZFER_NUEVO
        → /NZMME0001
        → Re-establece todos los campos (Homologar, Centro, Color, Franja, ZPLA)
          por si SAP reseteo la pantalla al navegar con /N
        → Cambia P_MATER-LOW = zfer_nuevo
        → zmme0001_leer_posiciones_popup():
            - Presiona Comparar BOM (btnBUTTON1)
            - Lee popup wnd[1]: tabla tblZMME0001T_COMP → columna POSNR
            - Cierra popup
            - Retorna lista de posiciones (ej: ["0458"])
        → zmme0001_agregar_filas_bom(posiciones, zpla):
            - Por cada posicion:
              * Presiona Insert (btnT_LISTA_MATERIA_INSERT)
              * Llena POSNR en columna 0
              * Consulta ODATA_ZPLA_BOM en BD → obtiene CLASE (Clave Destino)
              * Llena CLASE_DESTINO en columna 3
        → zmme0001_segunda_comparar_y_copy():
            - Segunda Comparar BOM (btnBUTTON1)
            - Verifica popup: si tipo "E" → error
            - Cierra popup
            - Presiona COPY_ITEM (btnCOPY_ITEM)

PASO 5: mm02_actualizar_partnumber(zfer_nuevo, nuevo_partnumber)
        → Lee PARTNUMBER del ZFER base desde MM02
        → Construye nuevo PARTNUMBER reemplazando segmento de color (indice 3)
          Patron: {codigo}_{seq}_{formula}_{color_num}_{version}
          Ej: "1407_000_L40-2_01_002" con p_color="19" → "1407_000_L40-2_19_002"
        → MM02 del ZFER_NUEVO → tab Clasificacion → tab PIEZA → actualiza fila 0
        → Guarda (btn[0] dos veces) + confirma dialogo si aparece
        → Idem para ZFOR_NUEVO si existe
```

---

## 5. CONFIGURACION DE BASE DE DATOS

### BD Produccion (Azure SQL)
```python
DB_PROD = {
    "server":   "agpcol.database.windows.net",
    "database": "agpc-productivity",
    "driver":   "ODBC Driver 17 for SQL Server",
    "user":     "Consulta",
    "password": "@GPgl4$$2021",
}
```

### Tablas clave en BD Produccion

| Tabla / Vista | Uso | Columnas clave |
|---|---|---|
| `ZFER_Characteristics_Genesis` | Formula del ZFER base | `SpecID` (= ZFER), `FormulaCode` |
| `TCAL_CALENDARIO_COLOMBIA_DIRECT` | Formula del ZFER base (fallback) | `ZFER`, `Formula`, `Mercado`, `Color` |
| `ODATA_ZPLA_BOM` | Clase Destino por posicion BOM | `MATERIAL` (= ZPLA), `POSICION`, `CLASE` |
| `VW_AppEnvolvente_LandMacro` | **DEPRECADA** — reemplazada por las dos tablas de arriba | — |

**IMPORTANTE:** La logica de busqueda de formula es:
1. Buscar en `ZFER_Characteristics_Genesis` (WHERE SpecID = ZFER)
2. Si no encuentra → buscar en `TCAL_CALENDARIO_COLOMBIA_DIRECT` (WHERE ZFER = ZFER)
3. Si no encuentra en ninguna → bloquear ejecucion SAP y mostrar error

### BD Local (SQL Server Express) — Log de ejecuciones
```python
DB_LOCAL = {
    "server":   r"localhost\SQLEXPRESS",
    "database": "MODULO_5",
    "driver":   "ODBC Driver 17 for SQL Server",
}
```
Tabla: `dbo.M5_LogEjecucion` — registra cada resultado procesado.
Si no existe la BD local, los errores de log se ignoran (solo warning en consola).

---

## 6. CAMPOS SAP — IDs CONFIRMADOS POR VBS

Todos los IDs de controles SAP fueron extraidos de grabaciones VBS del SAP GUI Recorder.

### ZMME0001 — Pantalla principal
```
wnd[0]/tbar[0]/okcd              → campo T-Code
wnd[0]/usr/radRB5                → radio Homologar
wnd[0]/usr/ctxtP_MATER-LOW      → campo Material ZFER (directo, sin popup)
wnd[0]/usr/ctxtP_CENTER         → campo Centro (CO01)
wnd[0]/usr/radRB3_A1            → radio Cambio de Color
wnd[0]/usr/ctxtP_COLOR          → campo Color (numero, ej: "19")
wnd[0]/usr/ctxtP_FRANJ          → campo Franja (codigo: "00","01","02","03","NA")
wnd[0]/usr/ctxtP_ZPLA           → campo ZPLA referencia
wnd[1]/usr/cntlLO_CONTAINER0500/shellcont/shell  → popup F4 lista de ZPLAs
wnd[0]/tbar[1]/btn[8]           → boton Ejecutar (F8)
wnd[0]/usr/cntlGRID1/shellcont/shell  → grid resultado despues de F8
wnd[0]/usr/btnBUTTON1           → boton Comparar BOM
```

### ZMME0001 — Tabla inferior (paso 4, despues de Comparar BOM)
```
wnd[0]/usr/tabsTABSTRIP_MAX/tabpPUSH1/ssub%_SUBSCREEN_MAX:ZMME0001:0200/btnT_LISTA_MATERIA_INSERT
wnd[0]/usr/tabsTABSTRIP_MAX/tabpPUSH1/ssub%_SUBSCREEN_MAX:ZMME0001:0200/tblZMME0001T_LISTA_MATERIA
  → /txtWA_LISTA-POSNR[0,{fila}]           columna 0: numero de posicion
  → /ctxtWA_LISTA-CLASE_DESTINO[3,{fila}]  columna 3: clave destino
wnd[0]/usr/tabsTABSTRIP_MAX/tabpPUSH1/ssub%_SUBSCREEN_MAX:ZMME0001:0200/btnCOPY_ITEM
```

### ZPPR0020 — Reporte de fases
```
wnd[0]/usr/txtS_USER-LOW         → campo Modificado por (= "PROGRAING")
wnd[0]/usr/ctxtS_WERKS-LOW      → campo Centro (= "CO01")
wnd[0]/tbar[1]/btn[8]           → Ejecutar (F8)
[grid ALV — nombre NO confirmado, el codigo intenta:]
  wnd[0]/usr/cntlGRID/shellcont/shell
  wnd[0]/usr/cntlGRID1/shellcont/shell
  wnd[0]/usr/cntlEUGRID/shellcont/shell
  wnd[0]/usr/cntlZPPR_GRID/shellcont/shell
[columnas del grid — nombres NO confirmados, el codigo intenta variantes:]
  ZFER | MATNR_ZFER | ZFER_NEW | MAT_ZFER | MATNR   → numero ZFER nuevo
  ZPLA | MATNR_ZPLA | ZPLA_NEW | MAT_ZPLA           → numero ZPLA
  FASE1..FASE15 | FASE_01..FASE_15 | F01..F15       → estado de cada fase
```

**PENDIENTE CONFIRMAR:** Los nombres exactos de columnas del grid de ZPPR0020.
Cuando el codigo encuentra la fila, imprime:
`ZPPR0020: fila X encontrada — ZPLA=... fases={...}`
Si `fases={}` es porque los nombres de columna no coinciden.

### MM02 — Clasificacion PIEZA
```
wnd[0]/usr/ctxtRMMG1-MATNR      → campo Material
wnd[0]/usr/tabsTABSPR1/tabpSP03 → tab Clasificacion
wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB4  → tab PIEZA
wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB4
  /ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S
  /ctxtRCTMS-MWERT[1,0]   → PARTNUMBER AGP (fila 0)
  /ctxtRCTMS-MWERT[1,1]   → COLOR (fila 1)
  /ctxtRCTMS-MWERT[1,2]   → FRANJA (fila 2, codigo SAP: "00","01","02","03","NA")
wnd[0]/tbar[0]/btn[0]           → Guardar
wnd[1]/usr/btnSPOP-OPTION1      → Confirmar dialogo de guardar (si aparece)
```

---

## 7. LOGICA DE FRANJA

La FRANJA determina el parametro `P_FRANJ` que se envia a ZMME0001:

| Valor en MM02 PIEZA | Significado | P_FRANJ enviado |
|---|---|---|
| `""` o `"00"` | Sin Franja | `"00"` |
| `"01"` | Franja Azul | `"01"` |
| `"02"` | Franja Verde | `"02"` |
| `"03"` | Franja Gris | `"03"` |
| `"NA"` | No Aplica | `"NA"` |

**IMPORTANTE:** Se lee la FRANJA del ZFER BASE (no del nuevo), se usa el mismo codigo
para todas las combinaciones del batch. La franja no cambia cuando cambia el color.

---

## 8. LOGICA DE COLOR

El campo `P_COLOR` en ZMME0001 recibe el **numero** del color, extraido del nombre:
```
"19-Gray Light Automotive" → "19"
"21-Gray Dark Automotive"  → "21"
"G2 Gray Medium"           → "" (no tiene numero, se advierte en log)
```
La funcion `_extraer_numero_color(color: str) -> str` implementa esta logica.

---

## 9. LOGICA DE PARTNUMBER

Patron del PARTNUMBER AGP:
```
{codigo_pedido}_{secuencia}_{formula_code}_{color_num}_{version}
Ejemplo: "1407_000_L40-2_01_002"
                             ^^ indice [3] = color_num → se reemplaza con p_color
```
Para construir el PARTNUMBER del ZFER nuevo, se toma el del ZFER base y se
reemplaza el segmento [3] con el nuevo numero de color.

---

## 10. REPORTE EXCEL (4 hojas)

Generado al final del batch en `reporte_sap_YYYYMMDD_HHMMSS.xlsx`:

| Hoja | Contenido |
|---|---|
| `RESUMEN` | Info del batch, totales, timing, operador |
| `PROCESADOS_SAP` | Todos los items enviados a SAP (OK en verde, ERROR en rojo) |
| `SOLO_REPORTE` | Items no procesados (formula diferente) en amarillo |
| `ERRORES` | Solo los errores, para revision rapida tecnico SAP |

Ademas se genera `progreso_XXXXXXXX.json` que se actualiza despues de CADA item
(checkpoint de recuperacion en caso de interrupcion).

---

## 11. ESTRUCTURA DE DATOS PRINCIPALES

### ItemColor (de VISTAAAA.py)
```python
item.zfer_origen  # ZFER base (str)
item.formula      # Codigo de formula (str, ej: "L19-13")
item.acero        # Variante de acero (str)
item.color        # Nombre completo del color (str, ej: "19-Gray Light Automotive")
item.mercado      # Mercado (str, ej: "COLOMBIA", "CentroAmerica")
item.bloqueado    # bool
item.pendiente    # bool
item.cod_pieza    # Codigo tipo pieza (str)
item.tipo_pieza   # Nombre tipo pieza (str)
```

### ResultadoCombinacion (de SAP_AUTOMATIZADOR.py)
```python
res.batch_id       # UUID del lote
res.zfer_base      # ZFER de entrada
res.zfer_nuevo     # ZFER creado por SAP
res.zfor_nuevo     # ZFOR creado por SAP (puede ser "")
res.posiciones_bom # Lista de posiciones BOM procesadas (ej: ["0458"])
res.estado         # "OK" | "ERROR" | "PENDIENTE"
res.error          # Mensaje de error (si aplica)
res.duracion_seg   # Tiempo total del paso en segundos
```

---

## 12. CONFIGURACION Y CONSTANTES

```python
# SAP_AUTOMATIZADOR.py
T_RAPIDO = 0.4   # segundos entre clicks simples
T_MEDIO  = 1.2   # despues de navegacion / F4
T_LENTO  = 2.5   # despues de ejecutar transacciones pesadas

_SAP_USER = "PROGRAING"   # usuario SAP para ZPPR0020 (campo Mod.por)
BASE_DIR  = r"C:\Users\abotero\OneDrive - AGP GROUP\Documentos\MODULO_5"
```

---

## 13. PENDIENTES / UNKNOWNS

Los siguientes puntos NO han sido confirmados con pruebas reales en SAP:

1. **Nombres de columnas del grid ZPPR0020** — el codigo intenta multiples variantes
   pero los nombres exactos dependen de la configuracion del sistema. Si `fases={}`
   en el log, hay que inspeccionar el grid con SAP GUI Inspector y ajustar.

2. **Si COPY_ITEM es suficiente o necesita "Ejecutar BOM" adicional** — el VBS
   termina en `btnCOPY_ITEM`. El usuario menciono un paso de "Ejecutar BOM" separado
   con nodos F00038 y ZINGP0003. Pendiente confirmar si es necesario.

3. **ZMMR0005** — el usuario menciono verificar la lista de materiales generada.
   No esta implementado aun.

4. **Popup de Comparar BOM — columna POSNR** — el codigo intenta leer la columna "POSNR"
   del popup. Si el nombre real es diferente, las posiciones quedaran vacias.

5. **Estado "S" en ZPPR0020** — el codigo espera que Fase 8+ tenga valor "S" (de "Success"
   o "Satisfactorio"). Si el sistema usa otro caracter, ajustar en `zppr0020_esperar_fases`.

---

## 14. COMO EJECUTAR

```bash
# Requisitos
pip install pyodbc openpyxl Pillow pywin32

# Ejecutar la app principal
py MODULO5.py

# Probar solo la conexion SAP (sin UI)
py SAP_AUTOMATIZADOR.py

# Probar solo el generador de combinaciones
set M5_ZFER_BASE=700179044
py COMBINADOR.py
```

**Pre-requisitos para SAP:**
- SAP GUI abierto con sesion activa (entorno QUAS para pruebas)
- Scripting habilitado: tuerca → Options → Accessibility & Scripting → Enable Scripting ✓
- El usuario SAP debe tener acceso a: MM02, ZMME0001, ZPPR0020

---

## 15. HISTORIAL DE DECISIONES TECNICAS

- **Mercado automatico:** El COMBINADOR consulta el mercado desde la BD, no hay selector
  manual en la UI. Se elimino el radio button de mercado.

- **Filtro por formula en UI thread:** La consulta a BD para detectar la formula base
  ocurre en el hilo principal de Tkinter (antes del dialogo de confirmacion) para
  poder mostrar numeros reales (X van a SAP, Y solo reporte) al usuario.

- **Dos tablas de formula:** `VW_AppEnvolvente_LandMacro` fue deprecada. Ahora se
  busca primero en `ZFER_Characteristics_Genesis` y luego en
  `TCAL_CALENDARIO_COLOMBIA_DIRECT`. Si no se encuentra en ninguna, se bloquea.

- **Re-establecer campos en paso 4:** Al navegar de vuelta a ZMME0001 con `/N` desde
  ZPPR0020, SAP puede resetear los campos. El codigo re-establece todos los campos
  (Homologar, Centro, Color, Franja, ZPLA) antes de cambiar el material a ZFER_NUEVO.

- **Franja como codigo directo:** La FRANJA se lee como codigo SAP ("00","01","02","03","NA")
  y se pasa tal cual a P_FRANJ. No se mapea de texto a codigo.
