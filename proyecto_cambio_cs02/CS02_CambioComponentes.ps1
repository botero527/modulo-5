# CS02 - Cambio Masivo de Componentes
# PowerShell + Windows Forms
# Uso: clic derecho -> "Ejecutar con PowerShell"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ── COLORES ──────────────────────────────────────────────────────────────────
$BG        = [System.Drawing.ColorTranslator]::FromHtml("#1e2230")
$BG2       = [System.Drawing.ColorTranslator]::FromHtml("#252a3d")
$BG3       = [System.Drawing.ColorTranslator]::FromHtml("#0f1117")
$BORDER    = [System.Drawing.ColorTranslator]::FromHtml("#3a3f55")
$FG        = [System.Drawing.ColorTranslator]::FromHtml("#e0e0e0")
$FG_DIM    = [System.Drawing.ColorTranslator]::FromHtml("#8892b0")
$C_OK      = [System.Drawing.ColorTranslator]::FromHtml("#50fa7b")
$C_ALERT   = [System.Drawing.ColorTranslator]::FromHtml("#ffb86c")
$C_ERROR   = [System.Drawing.ColorTranslator]::FromHtml("#ff5555")
$C_INFO    = [System.Drawing.ColorTranslator]::FromHtml("#8be9fd")
$C_HDR     = [System.Drawing.ColorTranslator]::FromHtml("#bd93f9")
$C_ACCENT  = [System.Drawing.ColorTranslator]::FromHtml("#5c6ef8")
$C_STOP    = [System.Drawing.ColorTranslator]::FromHtml("#e05c6a")
$FONT_MONO = New-Object System.Drawing.Font("Consolas", 9)
$FONT_UI   = New-Object System.Drawing.Font("Segoe UI", 9)
$FONT_LBL  = New-Object System.Drawing.Font("Segoe UI", 8)
$FONT_BTN  = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

# ── ESTADO GLOBAL ─────────────────────────────────────────────────────────────
$script:Detener = $false
$script:LogFile = ""
$script:nOk     = 0
$script:nAlert  = 0
$script:nError  = 0
$script:nTotal  = 0

# ── FORMULARIO PRINCIPAL ──────────────────────────────────────────────────────
$Form = New-Object System.Windows.Forms.Form
$Form.Text            = "CS02 - Cambio Masivo de Componentes"
$Form.Size            = New-Object System.Drawing.Size(800, 720)
$Form.StartPosition   = "CenterScreen"
$Form.BackColor       = $BG
$Form.ForeColor       = $FG
$Form.Font            = $FONT_UI
$Form.FormBorderStyle = "FixedSingle"
$Form.MaximizeBox     = $false

function New-Label($text, $x, $y, $w, $h, $font=$FONT_UI, $color=$FG) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text      = $text
    $l.Location  = New-Object System.Drawing.Point($x, $y)
    $l.Size      = New-Object System.Drawing.Size($w, $h)
    $l.ForeColor = $color
    $l.Font      = $font
    $l.BackColor = [System.Drawing.Color]::Transparent
    return $l
}

function New-Panel($x, $y, $w, $h) {
    $p = New-Object System.Windows.Forms.Panel
    $p.Location  = New-Object System.Drawing.Point($x, $y)
    $p.Size      = New-Object System.Drawing.Size($w, $h)
    $p.BackColor = $BG2
    return $p
}

# ── ENCABEZADO ────────────────────────────────────────────────────────────────
$lblTitulo = New-Label "⚙  CS02 — Cambio Masivo de Componentes" 14 12 600 26 `
    (New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)) $FG
$lblSub    = New-Label "Planta CO01  |  Utilización 1  |  Todas las alternativas" 14 40 600 18 `
    $FONT_LBL $FG_DIM

# ── PANEL DATOS ───────────────────────────────────────────────────────────────
$pnlDatos = New-Panel 10 68 768 175
$lblDatosTit = New-Label "DATOS A PROCESAR" 10 8 300 16 $FONT_LBL $FG_DIM
$lblHint     = New-Label "En Excel copia las columnas MAT_CONFIG, POSICION y COMPONENTE (sin encabezado) y pega con Ctrl+V." `
    10 28 748 16 $FONT_LBL $FG_DIM

$txtDatos = New-Object System.Windows.Forms.RichTextBox
$txtDatos.Location     = New-Object System.Drawing.Point(10, 48)
$txtDatos.Size         = New-Object System.Drawing.Size(748, 115)
$txtDatos.BackColor    = $BG3
$txtDatos.ForeColor    = $FG
$txtDatos.Font         = $FONT_MONO
$txtDatos.BorderStyle  = "None"
$txtDatos.ScrollBars   = "Vertical"

$pnlDatos.Controls.AddRange(@($lblDatosTit, $lblHint, $txtDatos))

# ── PANEL PROGRESO ────────────────────────────────────────────────────────────
$pnlProg = New-Panel 10 252 768 95
$lblProgTit = New-Label "PROGRESO" 10 8 200 16 $FONT_LBL $FG_DIM

# Cajas de estadísticas
function New-StatBox($x, $label, $color) {
    $p = New-Object System.Windows.Forms.Panel
    $p.Location  = New-Object System.Drawing.Point($x, 28)
    $p.Size      = New-Object System.Drawing.Size(175, 52)
    $p.BackColor = $BG3
    $num = New-Object System.Windows.Forms.Label
    $num.Text      = "0"
    $num.Location  = New-Object System.Drawing.Point(0, 6)
    $num.Size      = New-Object System.Drawing.Size(175, 28)
    $num.TextAlign = "MiddleCenter"
    $num.Font      = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $num.ForeColor = $color
    $num.BackColor = [System.Drawing.Color]::Transparent
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $label
    $lbl.Location  = New-Object System.Drawing.Point(0, 34)
    $lbl.Size      = New-Object System.Drawing.Size(175, 14)
    $lbl.TextAlign = "MiddleCenter"
    $lbl.Font      = $FONT_LBL
    $lbl.ForeColor = $FG_DIM
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    $p.Controls.AddRange(@($num, $lbl))
    return $p, $num
}

$boxTotal, $numTotal = New-StatBox 10   "Total"       $FG
$boxOk,    $numOk    = New-StatBox 193  "Guardados OK" $C_OK
$boxAlert, $numAlert = New-StatBox 376  "Alertas"      $C_ALERT
$boxError, $numError = New-StatBox 559  "Errores"      $C_ERROR

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 83)
$progressBar.Size     = New-Object System.Drawing.Size(748, 6)
$progressBar.Style    = "Continuous"
$progressBar.BackColor = $BG3
$progressBar.ForeColor = $C_ACCENT
$progressBar.Minimum  = 0
$progressBar.Maximum  = 100
$progressBar.Value    = 0

$lblStatus = New-Label "Listo." 10 85 748 16 $FONT_LBL $FG_DIM

$pnlProg.Controls.AddRange(@($lblProgTit, $boxTotal, $boxOk, $boxAlert, $boxError, $progressBar, $lblStatus))

# ── PANEL LOG ─────────────────────────────────────────────────────────────────
$pnlLog = New-Panel 10 356 768 280
$lblLogTit = New-Label "LOG EN TIEMPO REAL" 10 8 300 16 $FONT_LBL $FG_DIM

$btnLimpiarLog = New-Object System.Windows.Forms.Button
$btnLimpiarLog.Text      = "Limpiar"
$btnLimpiarLog.Location  = New-Object System.Drawing.Point(580, 4)
$btnLimpiarLog.Size      = New-Object System.Drawing.Size(70, 22)
$btnLimpiarLog.BackColor = $BG3
$btnLimpiarLog.ForeColor = $FG_DIM
$btnLimpiarLog.FlatStyle = "Flat"
$btnLimpiarLog.Font      = $FONT_LBL
$btnLimpiarLog.FlatAppearance.BorderColor = $BORDER

$btnAbrirLog = New-Object System.Windows.Forms.Button
$btnAbrirLog.Text      = "Abrir log .txt"
$btnAbrirLog.Location  = New-Object System.Drawing.Point(654, 4)
$btnAbrirLog.Size      = New-Object System.Drawing.Size(100, 22)
$btnAbrirLog.BackColor = $BG3
$btnAbrirLog.ForeColor = $FG_DIM
$btnAbrirLog.FlatStyle = "Flat"
$btnAbrirLog.Font      = $FONT_LBL
$btnAbrirLog.FlatAppearance.BorderColor = $BORDER

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location    = New-Object System.Drawing.Point(10, 28)
$logBox.Size        = New-Object System.Drawing.Size(748, 246)
$logBox.BackColor   = $BG3
$logBox.ForeColor   = $FG
$logBox.Font        = $FONT_MONO
$logBox.ReadOnly    = $true
$logBox.BorderStyle = "None"
$logBox.ScrollBars  = "Vertical"
$logBox.WordWrap    = $false

$pnlLog.Controls.AddRange(@($lblLogTit, $btnLimpiarLog, $btnAbrirLog, $logBox))

# ── BOTONES ACCION ────────────────────────────────────────────────────────────
$lblEstado = New-Label "Sin datos." 14 648 460 20 $FONT_LBL $FG_DIM

$btnIniciar = New-Object System.Windows.Forms.Button
$btnIniciar.Text      = "▶  Iniciar"
$btnIniciar.Location  = New-Object System.Drawing.Point(618, 642)
$btnIniciar.Size      = New-Object System.Drawing.Size(120, 34)
$btnIniciar.BackColor = $C_ACCENT
$btnIniciar.ForeColor = [System.Drawing.Color]::White
$btnIniciar.FlatStyle = "Flat"
$btnIniciar.Font      = $FONT_BTN
$btnIniciar.FlatAppearance.BorderSize = 0

$btnDetener = New-Object System.Windows.Forms.Button
$btnDetener.Text      = "■  Detener"
$btnDetener.Location  = New-Object System.Drawing.Point(744, 642)
$btnDetener.Size      = New-Object System.Drawing.Size(32, 34)   # oculto, se muestra al correr
$btnDetener.BackColor = $C_STOP
$btnDetener.ForeColor = [System.Drawing.Color]::White
$btnDetener.FlatStyle = "Flat"
$btnDetener.Font      = $FONT_BTN
$btnDetener.Enabled   = $false
$btnDetener.Visible   = $false
$btnDetener.FlatAppearance.BorderSize = 0

$Form.Controls.AddRange(@(
    $lblTitulo, $lblSub,
    $pnlDatos, $pnlProg, $pnlLog,
    $lblEstado, $btnIniciar, $btnDetener
))

# ── FUNCIONES UI ──────────────────────────────────────────────────────────────
function Log-UI {
    param([string]$tipo, [string]$msg)
    $hora = (Get-Date).ToString("HH:mm:ss")
    $pre  = switch ($tipo) {
        "OK"    { "[OK]    " }
        "ALERT" { "[ALERTA]" }
        "ERR"   { "[ERROR] " }
        "INFO"  { "[info]  " }
        "HDR"   { ""         }
        default { "[skip]  " }
    }
    $color = switch ($tipo) {
        "OK"    { $C_OK    }
        "ALERT" { $C_ALERT }
        "ERR"   { $C_ERROR }
        "INFO"  { $C_INFO  }
        "HDR"   { $C_HDR   }
        default { $FG_DIM  }
    }
    $linea = "$hora  $pre  $msg"
    $logBox.SelectionStart  = $logBox.TextLength
    $logBox.SelectionLength = 0
    $logBox.SelectionColor  = $color
    $logBox.AppendText("$linea`n")
    $logBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()

    if ($script:LogFile -ne "") {
        Add-Content -Path $script:LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $tipo | $msg"
    }
}

function Update-Stats {
    $numTotal.Text = $script:nTotal
    $numOk.Text    = $script:nOk
    $numAlert.Text = $script:nAlert
    $numError.Text = $script:nError
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-Status($s) {
    $lblStatus.Text = $s
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-Progreso($actual, $total) {
    if ($total -gt 0) {
        $pct = [int](($actual / $total) * 100)
        $progressBar.Value = [Math]::Min($pct, 100)
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Pad-Pos($pos) {
    try { return ([int]$pos).ToString("0000") }
    catch { return $pos }
}

# ── EVENTOS BOTONES ───────────────────────────────────────────────────────────
$btnLimpiarLog.Add_Click({ $logBox.Clear() })

$btnAbrirLog.Add_Click({
    if ($script:LogFile -ne "" -and (Test-Path $script:LogFile)) {
        Start-Process notepad.exe $script:LogFile
    } else {
        [System.Windows.Forms.MessageBox]::Show("Ejecute el proceso primero.", "Info") | Out-Null
    }
})

$btnDetener.Add_Click({
    $script:Detener = $true
    Log-UI "ALERT" "Deteniendo... espere que termine la operacion actual."
})

$btnIniciar.Add_Click({
    # ── Parsear datos pegados ──────────────────────────────────────────────
    $texto = $txtDatos.Text.Trim()
    if ($texto -eq "") {
        [System.Windows.Forms.MessageBox]::Show(
            "Pegue los datos primero.`nColumnas: MAT_CONFIG, POSICION, COMPONENTE (sin encabezado).",
            "Sin datos", "OK", "Warning") | Out-Null
        return
    }

    $filas = @()
    foreach ($linea in $texto -split "`n") {
        $l = $linea.Trim().TrimEnd("`r")
        if ($l -eq "") { continue }
        $cols = $l -split "`t"
        if ($cols.Count -lt 3) { continue }
        $m = $cols[0].Trim(); $p = $cols[1].Trim(); $c = $cols[2].Trim()
        if ($m -ne "" -and $p -ne "" -and $c -ne "") {
            $filas += [pscustomobject]@{ Mat=$m; Pos=$p; Comp=$c }
        }
    }

    if ($filas.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No se encontraron filas válidas.`nAsegúrese de copiar SIN encabezado, 3 columnas separadas por tabulador.",
            "Sin datos", "OK", "Warning") | Out-Null
        return
    }

    # ── Reset estado ──────────────────────────────────────────────────────
    $script:Detener = $false
    $script:nOk = 0; $script:nAlert = 0; $script:nError = 0
    $script:nTotal = $filas.Count
    Update-Stats
    $progressBar.Value = 0

    $btnIniciar.Enabled = $false
    $txtDatos.Enabled   = $false
    $btnDetener.Enabled = $true
    $btnDetener.Visible = $true
    $btnDetener.Location = New-Object System.Drawing.Point(626, 642)
    $btnIniciar.Location = New-Object System.Drawing.Point(500, 642)

    # Log file y archivos temp
    if (-not (Test-Path "C:\Temp")) { New-Item -ItemType Directory -Path "C:\Temp" | Out-Null }
    $script:LogFile = "C:\Temp\CS02_" + (Get-Date -Format "yyyyMMdd_HHmm") + ".txt"
    $inFile  = "C:\Temp\cs02_input.txt"
    $outFile = "C:\Temp\cs02_output.txt"

    # Escribir archivo de entrada para el worker VBS
    $lineasIn = $filas | ForEach-Object { "$($_.Mat)|$($_.Pos)|$($_.Comp)" }
    [System.IO.File]::WriteAllLines($inFile, $lineasIn, [System.Text.Encoding]::UTF8)

    # Borrar archivo de salida anterior
    if (Test-Path $outFile) { Remove-Item $outFile -Force }

    Log-UI "HDR" "==========================================="
    Log-UI "HDR" "  INICIO $(Get-Date)  |  $($filas.Count) filas"
    Log-UI "HDR" "==========================================="

    # Ruta al worker VBS (misma carpeta que este script)
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    if (-not $scriptDir) { $scriptDir = $PSScriptRoot }
    if (-not $scriptDir) { $scriptDir = "." }
    $workerVbs = Join-Path $scriptDir "CS02_Worker.vbs"

    if (-not (Test-Path $workerVbs)) {
        Log-UI "ERR" "No se encontro CS02_Worker.vbs en: $scriptDir"
        Finalizar-UI; return
    }

    # Lanzar worker en segundo plano
    $proc = Start-Process -FilePath "cscript.exe" `
        -ArgumentList "//NoLogo `"$workerVbs`"" `
        -WindowStyle Hidden -PassThru

    Log-UI "INFO" "Worker SAP iniciado (PID $($proc.Id))"

    # ── Timer para monitorear salida en tiempo real ────────────────────────
    $script:outLineIdx = 0
    $script:procEnded  = $false

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 350

    $timer.Add_Tick({
        # Leer nuevas lineas del archivo de salida
        if (Test-Path $outFile) {
            try {
                $allLines = [System.IO.File]::ReadAllLines($outFile, [System.Text.Encoding]::UTF8)
            } catch { $allLines = @() }

            while ($script:outLineIdx -lt $allLines.Count) {
                $raw = $allLines[$script:outLineIdx]
                $script:outLineIdx++
                if ($raw.Trim() -eq "") { continue }

                $p = $raw -split '\|', 5
                $status = $p[0]
                $mat    = if ($p.Count -gt 1) { $p[1] } else { "" }
                $pos    = if ($p.Count -gt 2) { $p[2] } else { "" }
                $comp   = if ($p.Count -gt 3) { $p[3] } else { "" }
                $msg    = if ($p.Count -gt 4) { $p[4] } else { "" }

                switch ($status) {
                    "DONE" {
                        $script:procEnded = $true
                        $timer.Stop()
                        Log-UI "HDR" "==========================================="
                        Log-UI "HDR" "  PROCESO COMPLETADO"
                        Log-UI "HDR" "  OK: $($script:nOk)  Alertas: $($script:nAlert)  Errores: $($script:nError)"
                        Log-UI "HDR" "  Log: $($script:LogFile)"
                        Log-UI "HDR" "==========================================="
                        Finalizar-UI
                    }
                    "SAP_ERROR" {
                        Log-UI "ERR" "No se pudo conectar a SAP: $msg"
                        Log-UI "ERR" "SAP debe estar abierto con sesion activa y scripting habilitado."
                    }
                    "OK" {
                        $script:nOk++
                        Log-UI "OK" "[${mat}] Pos $pos -> $comp | $msg"
                        $script:nTotal2 = if ($script:nTotal2) { $script:nTotal2 } else { $script:nTotal }
                        $done = $script:nOk + $script:nAlert + $script:nError
                        Set-Progreso $done $script:nTotal
                        Set-Status "Mat: $mat | Pos: $pos -> $comp"
                        Update-Stats
                    }
                    "ALERT" {
                        $script:nAlert++
                        Log-UI "ALERT" "[${mat}] Pos $pos | $msg"
                        $done = $script:nOk + $script:nAlert + $script:nError
                        Set-Progreso $done $script:nTotal
                        Update-Stats
                    }
                    "ERROR" {
                        $script:nError++
                        Log-UI "ERR" "[${mat}] $msg"
                        $done = $script:nOk + $script:nAlert + $script:nError
                        Set-Progreso $done $script:nTotal
                        Update-Stats
                    }
                    "INFO" {
                        Log-UI "INFO" "[${mat}] $msg"
                    }
                }
            }
        }

        # Si el proceso termino pero no llego DONE (crash), cerrar
        if (-not $script:procEnded -and $proc.HasExited) {
            $timer.Stop()
            if (-not $script:procEnded) {
                Log-UI "ERR" "El worker termino inesperadamente (codigo $($proc.ExitCode))"
                Finalizar-UI
            }
        }

        # Boton Detener
        if ($script:Detener) {
            $timer.Stop()
            try { $proc.Kill() } catch {}
            Log-UI "ALERT" "Proceso detenido por el usuario."
            Finalizar-UI
        }
    })

    $timer.Start()
})

# ── LOGICA SAP ────────────────────────────────────────────────────────────────
function Invoke-CS02Material($sess, $mat, $pos, $comp) {
    $sess.findById("wnd[0]").maximize()
    $sess.findById("wnd[0]/tbar[0]/okcd").text = "/ncs02"
    $sess.findById("wnd[0]").sendVKey(0)
    Start-Sleep -Milliseconds 900

    try { $sess.findById("wnd[0]/usr/ctxtRC29N-MATNR").text = $mat }
    catch {
        Log-UI "ERR" "[${mat}] No se pudo acceder a CS02: $_"
        $script:nError++; Update-Stats; return
    }

    $sess.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = "CO01"
    $sess.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = "1"
    $sess.findById("wnd[0]/usr/txtRC29N-STLAL").text  = ""
    $sess.findById("wnd[0]").sendVKey(0)
    Start-Sleep -Milliseconds 1500

    # Revisar si existe
    $sbar = ""
    try { $sbar = $sess.findById("wnd[0]/sbar").text } catch {}
    if ($sbar -match "no existe|not exist|no hay") {
        Log-UI "ALERT" "[${mat}] LMat no encontrada: $sbar"
        $script:nAlert++; Update-Stats; return
    }

    # Detectar pantalla de alternativas
    $tblAlt = $null; $hayAlt = $false
    try { $tblAlt = $sess.findById("wnd[0]/usr/tblSAPLCSDITCS_ALT_MAT"); $hayAlt = $true } catch {}

    if ($hayAlt) {
        $nAlts = $tblAlt.RowCount
        Log-UI "INFO" "[${mat}] $nAlts alternativa(s)"

        for ($ia = 0; $ia -lt $nAlts; $ia++) {
            if ($script:Detener) { break }
            $altNum = "$($ia+1)"
            try { $altNum = $sess.findById("wnd[0]/usr/tblSAPLCSDITCS_ALT_MAT/txtRC29L-STLAL[0,$ia]").text.Trim() } catch {}
            if ($altNum -eq "") { $altNum = "$($ia+1)" }

            Set-Status "Mat: $mat  Alt: $altNum  buscando Pos $pos"

            $entro = $false
            try { $tblAlt.currentCellRow = $ia; Start-Sleep -Milliseconds 150; $tblAlt.doubleClickCurrentCell(); $entro = $true } catch {}

            if (-not $entro) {
                Log-UI "ALERT" "[${mat}] Alt ${altNum}: no se pudo abrir"
                $script:nAlert++; Update-Stats
            } else {
                Start-Sleep -Milliseconds 1300
                $ok = Invoke-CambiarComponente $sess $mat $pos $comp $altNum
                if ($ok) {
                    $sess.findById("wnd[0]/tbar[0]/btn[11]").press()
                    Start-Sleep -Milliseconds 900
                    Cerrar-Popup $sess
                    $msg = ""; try { $msg = $sess.findById("wnd[0]/sbar").text } catch {}
                    Log-UI "OK" "[${mat}] Alt $altNum  Pos $pos -> ${comp} |  $msg"
                    $script:nOk++; Update-Stats
                }
            }

            $sess.findById("wnd[0]").sendVKey(3)
            Start-Sleep -Milliseconds 900
            try { $tblAlt = $sess.findById("wnd[0]/usr/tblSAPLCSDITCS_ALT_MAT") } catch { break }
        }
    } else {
        Log-UI "INFO" "[${mat}] BOM directo"
        $ok = Invoke-CambiarComponente $sess $mat $pos $comp "1"
        if ($ok) {
            $sess.findById("wnd[0]/tbar[0]/btn[11]").press()
            Start-Sleep -Milliseconds 900
            Cerrar-Popup $sess
            $msg = ""; try { $msg = $sess.findById("wnd[0]/sbar").text } catch {}
            Log-UI "OK" "[${mat}] Pos $pos -> ${comp} |  $msg"
            $script:nOk++; Update-Stats
        }
    }
}

function Invoke-CambiarComponente($sess, $mat, $pos, $comp, $alt) {
    $T = "wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT"
    $tbl = $null
    try { $tbl = $sess.findById($T) }
    catch {
        Log-UI "ALERT" "[${mat}] Alt ${alt}: tabla BOM no encontrada"
        $script:nAlert++; Update-Stats; return $false
    }

    $nVis    = $tbl.RowCount
    $scrMax  = 0
    try { $scrMax = $tbl.VerticalScrollbar.Maximum } catch {}
    if ($nVis -le 0) { $nVis = 10 }

    $iScr = 0
    while ($true) {
        if ($iScr -gt 0) {
            try { $tbl.VerticalScrollbar.Position = $iScr; Start-Sleep -Milliseconds 300 } catch {}
        }
        for ($r = 0; $r -lt $nVis; $r++) {
            $pf = ""
            try { $pf = $sess.findById("$T/txtRC29P-POSNR[0,$r]").text.Trim() } catch {}
            if ($pf -ne "") {
                try { $pf = ([int]$pf).ToString("0000") } catch {}
            }
            if ($pf -eq $pos) {
                try {
                    $sess.findById("$T/ctxtRC29P-IDNRK[2,$r]").text = $comp
                    $sess.findById("$T/ctxtRC29P-IDNRK[2,$r]").setFocus()
                } catch {
                    Log-UI "ALERT" "[${mat}] Alt $alt Pos ${pos}: error al escribir — $_"
                    $script:nAlert++; Update-Stats; return $false
                }
                $sess.findById("wnd[0]").sendVKey(0)
                Start-Sleep -Milliseconds 500
                $sv = ""; try { $sv = $sess.findById("wnd[0]/sbar").text } catch {}
                if ($sv -match "error") {
                    Log-UI "ALERT" "[${mat}] SAP rechazó Pos $pos Alt $alt : $sv"
                    $script:nAlert++; Update-Stats; return $false
                }
                return $true
            }
        }
        if ($scrMax -le 0 -or $iScr -ge $scrMax) { break }
        $iScr += $nVis
    }
    Log-UI "ALERT" "[${mat}] Alt ${alt}: Posición [$pos] NO encontrada en BOM"
    $script:nAlert++; Update-Stats
    return $false
}

function Cerrar-Popup($sess) {
    try { $sess.findById("wnd[1]/usr/btnSPOP-OPTION1").press() } catch {
        try { $sess.findById("wnd[1]").sendVKey(0) } catch {}
    }
}

function Finalizar-UI {
    $btnIniciar.Enabled = $true
    $txtDatos.Enabled   = $true
    $btnDetener.Enabled = $false
    $btnDetener.Visible = $false
    $btnIniciar.Location = New-Object System.Drawing.Point(618, 642)
    $progressBar.Value = 100
    Set-Status "Finalizado  |  OK: $($script:nOk)  Alertas: $($script:nAlert)  Errores: $($script:nError)"
    $lblEstado.Text = "OK: $($script:nOk)   Alertas: $($script:nAlert)   Errores: $($script:nError)"
    [System.Windows.Forms.Application]::DoEvents()
}

# ── ARRANCAR ──────────────────────────────────────────────────────────────────
[System.Windows.Forms.Application]::Run($Form)

