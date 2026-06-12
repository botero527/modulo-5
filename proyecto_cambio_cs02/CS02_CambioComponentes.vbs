' =============================================================
' CS02 - Cambio Masivo de Componentes en Lista de Materiales
' Columnas requeridas en Excel: MAT_CONFIG, POSICION, COMPONENTE
' Planta fija: CO01 | Utilizacion: 1
' =============================================================
Option Explicit

Dim SapGuiAuto, application, connection, session
Dim objExcelApp, objWorkbook, objSheet
Dim sFile, sLog
Dim i, lastRow
Dim colMat, colPos, colComp
Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

' ---- Carpeta de log ----
Dim sLogDir
sLogDir = "C:\Temp"
If Not fso.FolderExists(sLogDir) Then fso.CreateFolder(sLogDir)

Dim sStamp
sStamp = Year(Now) & Right("00" & Month(Now),2) & Right("00" & Day(Now),2) & "_" & _
         Right("00" & Hour(Now),2) & Right("00" & Minute(Now),2) & Right("00" & Second(Now),2)
sLog = sLogDir & "\CS02_Log_" & sStamp & ".txt"

' ---- Seleccionar archivo Excel ----
sFile = InputBox("Ruta completa del archivo Excel:" & vbCrLf & _
                 "(columnas: MAT_CONFIG, POSICION, COMPONENTE)", _
                 "CS02 - Cambio Masivo de Componentes", _
                 "C:\Temp\datos.xlsx")
If sFile = "" Then WScript.Quit
If Not fso.FileExists(sFile) Then
    MsgBox "Archivo no encontrado:" & vbCrLf & sFile, vbCritical, "Error"
    WScript.Quit
End If

' ---- Abrir Excel ----
Set objExcelApp = CreateObject("Excel.Application")
objExcelApp.Visible = False
objExcelApp.DisplayAlerts = False

On Error Resume Next
Set objWorkbook = objExcelApp.Workbooks.Open(sFile)
If Err.Number <> 0 Then
    MsgBox "No se pudo abrir el archivo Excel:" & vbCrLf & Err.Description, vbCritical
    objExcelApp.Quit
    WScript.Quit
End If
On Error GoTo 0

Set objSheet = objWorkbook.Sheets(1)
lastRow = objSheet.UsedRange.Rows.Count

' ---- Detectar columnas por encabezado ----
colMat = -1 : colPos = -1 : colComp = -1
Dim colIdx, hdr
For colIdx = 1 To 30
    hdr = Trim(CStr(objSheet.Cells(1, colIdx).Value))
    Select Case UCase(hdr)
        Case "MAT_CONFIG"  : colMat  = colIdx
        Case "POSICION"    : colPos  = colIdx
        Case "COMPONENTE"  : colComp = colIdx
    End Select
Next

If colMat = -1 Or colPos = -1 Or colComp = -1 Then
    MsgBox "No se encontraron las columnas requeridas en la fila 1:" & vbCrLf & _
           "MAT_CONFIG=" & colMat & "  POSICION=" & colPos & "  COMPONENTE=" & colComp, vbCritical
    objWorkbook.Close False
    objExcelApp.Quit
    WScript.Quit
End If

' ---- Conectar a SAP ----
If Not IsObject(application) Then
    Set SapGuiAuto = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
    Set connection = application.Children(0)
End If
If Not IsObject(session) Then
    Set session = connection.Children(0)
End If
If IsObject(WScript) Then
    WScript.ConnectObject session,     "on"
    WScript.ConnectObject application, "on"
End If

WriteLog sLog, "=== INICIO CS02 | " & Now & " | Archivo: " & sFile & " | Filas datos: " & (lastRow-1) & " ==="

' ---- Procesar cada fila ----
Dim sMaterial, sPosicion, sComponente, skipRow
For i = 2 To lastRow
    skipRow = False

    sMaterial   = Trim(CStr(objSheet.Cells(i, colMat).Value))
    sPosicion   = Trim(CStr(objSheet.Cells(i, colPos).Value))
    sComponente = Trim(CStr(objSheet.Cells(i, colComp).Value))

    If sMaterial = "" Or sPosicion = "" Or sComponente = "" Then
        If sMaterial <> "" Then
            WriteLog sLog, "SKIP | Fila " & i & " | Datos incompletos (Mat=" & sMaterial & " Pos=" & sPosicion & " Comp=" & sComponente & ")"
        End If
        skipRow = True
    End If

    If Not skipRow Then
        ' Normalizar posicion a 4 digitos
        On Error Resume Next
        Dim nPosVal : nPosVal = CLng(sPosicion)
        If Err.Number <> 0 Then
            WriteLog sLog, "ALERTA | Fila " & i & " | Mat " & sMaterial & " | Posicion no numerica: [" & sPosicion & "]"
            Err.Clear
            skipRow = True
        Else
            sPosicion = Right("0000" & CStr(nPosVal), 4)
        End If
        On Error GoTo 0
    End If

    If Not skipRow Then
        On Error Resume Next
        Call ProcesarMaterial(session, sMaterial, sPosicion, sComponente, sLog)
        If Err.Number <> 0 Then
            WriteLog sLog, "ERROR CRITICO | Fila " & i & " | Mat " & sMaterial & " | " & Err.Description
            Err.Clear
            ' Resetear navegacion SAP
            On Error Resume Next
            session.findById("wnd[0]/tbar[0]/okcd").text = "/ncs02"
            session.findById("wnd[0]").sendVKey 0
            SapSleep 1000
            Err.Clear
            On Error GoTo 0
        End If
        On Error GoTo 0
    End If
Next

' ---- Cerrar ----
objWorkbook.Close False
objExcelApp.Quit

WriteLog sLog, "=== FIN | " & Now & " ==="
MsgBox "Proceso finalizado." & vbCrLf & vbCrLf & "Revise el log en:" & vbCrLf & sLog, _
       vbInformation, "CS02 - Completado"

' ==============================================================
'                      SUBRUTINAS Y FUNCIONES
' ==============================================================

Sub ProcesarMaterial(sess, sMat, sPos, sComp, sLogFile)
    ' Ir a CS02
    sess.findById("wnd[0]").maximize
    sess.findById("wnd[0]/tbar[0]/okcd").text = "/ncs02"
    sess.findById("wnd[0]").sendVKey 0
    SapSleep 1000

    ' Verificar que la pantalla CS02 este disponible
    On Error Resume Next
    sess.findById("wnd[0]/usr/ctxtRC29N-MATNR").text = sMat
    If Err.Number <> 0 Then
        WriteLog sLogFile, "ALERTA | Mat " & sMat & " | No se pudo escribir en pantalla CS02: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    sess.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = "CO01"
    sess.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = "1"
    sess.findById("wnd[0]/usr/txtRC29N-STLAL").text = ""   ' Alternativa en blanco
    sess.findById("wnd[0]").sendVKey 0
    SapSleep 1500

    ' Verificar si hay error en barra de estado (material no existe, etc.)
    Dim sSbar
    sSbar = ""
    On Error Resume Next
    sSbar = LCase(sess.findById("wnd[0]/sbar").text)
    On Error GoTo 0
    If InStr(sSbar, "no existe") > 0 Or InStr(sSbar, "not exist") > 0 Or _
       InStr(sSbar, "no hay") > 0 Or InStr(sSbar, "error") > 0 Then
        WriteLog sLogFile, "ALERTA | Mat " & sMat & " | Error al abrir LMat: " & sess.findById("wnd[0]/sbar").text
        Exit Sub
    End If

    ' Detectar pantalla de seleccion de alternativas
    Dim bPantallaAlt : bPantallaAlt = False
    On Error Resume Next
    Dim tblAlt
    Set tblAlt = sess.findById("wnd[0]/usr/tblSAPLCSDITCS_ALT_MAT")
    If Err.Number = 0 And Not IsNull(tblAlt) Then bPantallaAlt = True
    Err.Clear
    On Error GoTo 0

    If bPantallaAlt Then
        ' ---- HAY MULTIPLES ALTERNATIVAS ----
        Dim nAltTotal : nAltTotal = tblAlt.RowCount
        WriteLog sLogFile, "INFO | Mat " & sMat & " | " & nAltTotal & " alternativa(s) encontrada(s)"

        Dim iAlt
        For iAlt = 0 To nAltTotal - 1
            ' Obtener numero de alternativa de la tabla
            Dim sNumAlt
            On Error Resume Next
            sNumAlt = Trim(sess.findById("wnd[0]/usr/tblSAPLCSDITCS_ALT_MAT/txtRC29L-STLAL[0," & iAlt & "]").text)
            If Err.Number <> 0 Then sNumAlt = CStr(iAlt + 1) : Err.Clear
            On Error GoTo 0

            ' Seleccionar fila y hacer doble click para entrar
            On Error Resume Next
            tblAlt.selectedRows = iAlt
            SapSleep 300
            ' Doble click en primera celda para abrir la alternativa
            sess.findById("wnd[0]/usr/tblSAPLCSDITCS_ALT_MAT").currentCellRow = iAlt
            sess.findById("wnd[0]/usr/tblSAPLCSDITCS_ALT_MAT").doubleClickCurrentCell
            If Err.Number <> 0 Then
                WriteLog sLogFile, "ALERTA | Mat " & sMat & " | Alt " & sNumAlt & " | No se pudo abrir alternativa: " & Err.Description
                Err.Clear
                On Error GoTo 0
            Else
                On Error GoTo 0
                SapSleep 1200

                ' Cambiar componente en esta alternativa
                Dim bCambioOk
                bCambioOk = CambiarComponente(sess, sMat, sPos, sComp, sNumAlt, sLogFile)

                If bCambioOk Then
                    ' Guardar (btn[11] = Ctrl+S / Guardar)
                    sess.findById("wnd[0]/tbar[0]/btn[11]").press
                    SapSleep 1000
                    ' Cerrar posibles popups de confirmacion
                    CerrarPopup sess
                    Dim sMsgGuard
                    sMsgGuard = ""
                    On Error Resume Next
                    sMsgGuard = sess.findById("wnd[0]/sbar").text
                    On Error GoTo 0
                    WriteLog sLogFile, "GUARDADO | Mat " & sMat & " | Alt " & sNumAlt & " | Pos " & sPos & " -> " & sComp & " | " & sMsgGuard
                End If

                ' Volver a pantalla de alternativas (F3)
                sess.findById("wnd[0]").sendVKey 3
                SapSleep 1000

                ' Refrescar referencia a tabla de alternativas
                On Error Resume Next
                Set tblAlt = sess.findById("wnd[0]/usr/tblSAPLCSDITCS_ALT_MAT")
                If Err.Number <> 0 Then
                    WriteLog sLogFile, "ALERTA | Mat " & sMat & " | No se pudo volver a lista de alternativas. Abortando restantes."
                    Err.Clear
                    On Error GoTo 0
                    Exit For
                End If
                On Error GoTo 0
            End If
        Next

    Else
        ' ---- UNA SOLA ALTERNATIVA (entrada directa al BOM) ----
        WriteLog sLogFile, "INFO | Mat " & sMat & " | Entrada directa al BOM (1 alternativa)"

        Dim bCambioOk2
        bCambioOk2 = CambiarComponente(sess, sMat, sPos, sComp, "1", sLogFile)

        If bCambioOk2 Then
            sess.findById("wnd[0]/tbar[0]/btn[11]").press
            SapSleep 1000
            CerrarPopup sess
            Dim sMsgGuard2
            sMsgGuard2 = ""
            On Error Resume Next
            sMsgGuard2 = sess.findById("wnd[0]/sbar").text
            On Error GoTo 0
            WriteLog sLogFile, "GUARDADO | Mat " & sMat & " | Pos " & sPos & " -> " & sComp & " | " & sMsgGuard2
        End If
    End If
End Sub

' --------------------------------------------------------------
' Busca la posicion en la tabla BOM y cambia el componente
' Retorna True si tuvo exito, False si no encontro la posicion
' --------------------------------------------------------------
Function CambiarComponente(sess, sMat, sPos, sComp, sAlt, sLogFile)
    CambiarComponente = False
    Dim TABLE_ID
    TABLE_ID = "wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT"

    ' Verificar que existe la tabla
    Dim tblBOM
    On Error Resume Next
    Set tblBOM = sess.findById(TABLE_ID)
    If Err.Number <> 0 Then
        WriteLog sLogFile, "ALERTA | Mat " & sMat & " | Alt " & sAlt & " | Tabla BOM no encontrada: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Dim nVisRows  : nVisRows  = tblBOM.RowCount
    Dim bEncontrado : bEncontrado = False
    Dim scrollActual : scrollActual = 0
    Dim scrollMax    : scrollMax    = 0

    ' Obtener maximo scroll disponible
    On Error Resume Next
    scrollMax = tblBOM.VerticalScrollbar.Maximum
    If Err.Number <> 0 Then scrollMax = 0 : Err.Clear
    On Error GoTo 0

    ' Recorrer filas con scroll si es necesario
    Dim iScroll
    For iScroll = 0 To scrollMax Step nVisRows
        ' Scroll a la posicion actual
        If iScroll > 0 Then
            On Error Resume Next
            tblBOM.VerticalScrollbar.Position = iScroll
            SapSleep 400
            Err.Clear
            On Error GoTo 0
        End If

        ' Revisar filas visibles
        Dim r
        For r = 0 To nVisRows - 1
            Dim sPosEnFila
            sPosEnFila = ""
            On Error Resume Next
            sPosEnFila = Trim(sess.findById(TABLE_ID & "/txtRC29P-POSNR[0," & r & "]").text)
            Err.Clear
            On Error GoTo 0

            ' Normalizar a 4 digitos para comparar
            If sPosEnFila <> "" Then
                On Error Resume Next
                Dim nPF : nPF = CLng(sPosEnFila)
                If Err.Number = 0 Then
                    sPosEnFila = Right("0000" & CStr(nPF), 4)
                End If
                Err.Clear
                On Error GoTo 0
            End If

            If sPosEnFila = sPos Then
                bEncontrado = True
                ' Escribir nuevo componente en columna IDNRK (col index 2)
                On Error Resume Next
                sess.findById(TABLE_ID & "/ctxtRC29P-IDNRK[2," & r & "]").text = sComp
                sess.findById(TABLE_ID & "/ctxtRC29P-IDNRK[2," & r & "]").setFocus
                If Err.Number <> 0 Then
                    WriteLog sLogFile, "ALERTA | Mat " & sMat & " | Alt " & sAlt & " | Pos " & sPos & " | Error al escribir componente: " & Err.Description
                    Err.Clear
                    On Error GoTo 0
                    Exit Function
                End If
                On Error GoTo 0
                ' Enter para validar cambio
                sess.findById("wnd[0]").sendVKey 0
                SapSleep 600
                ' Verificar si SAP mostro algun error de validacion
                Dim sSbarVal
                sSbarVal = ""
                On Error Resume Next
                sSbarVal = LCase(sess.findById("wnd[0]/sbar").text)
                On Error GoTo 0
                If InStr(sSbarVal, "error") > 0 Then
                    WriteLog sLogFile, "ALERTA | Mat " & sMat & " | Alt " & sAlt & " | Pos " & sPos & " | SAP rechazo el componente: " & sess.findById("wnd[0]/sbar").text
                    Exit Function
                End If
                CambiarComponente = True
                Exit For
            End If
        Next

        If bEncontrado Then Exit For
        ' Si ya llegamos al ultimo bloque de filas, salir
        If iScroll + nVisRows > scrollMax And iScroll > 0 Then Exit For
        If scrollMax = 0 Then Exit For
    Next

    If Not bEncontrado Then
        WriteLog sLogFile, "ALERTA | Mat " & sMat & " | Alt " & sAlt & " | Posicion [" & sPos & "] NO encontrada en BOM"
    End If
End Function

' --------------------------------------------------------------
' Cierra popup de confirmacion si aparece (Enter o boton Si)
' --------------------------------------------------------------
Sub CerrarPopup(sess)
    On Error Resume Next
    Dim wnd1
    Set wnd1 = sess.findById("wnd[1]")
    If Err.Number = 0 And Not IsNull(wnd1) Then
        ' Intentar boton "Si" (Yes) o Enter
        sess.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Err.Number <> 0 Then
            Err.Clear
            sess.findById("wnd[1]").sendVKey 0
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' --------------------------------------------------------------
' Escribe una linea al archivo de log
' --------------------------------------------------------------
Sub WriteLog(sFile, sMsg)
    Dim fso2, f
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set f = fso2.OpenTextFile(sFile, 8, True)  ' 8 = ForAppending
    If Err.Number = 0 Then
        f.WriteLine Now & " | " & sMsg
        f.Close
    End If
    Err.Clear
    On Error GoTo 0
    Set f = Nothing
    Set fso2 = Nothing
End Sub

' --------------------------------------------------------------
' Pausa en milisegundos
' --------------------------------------------------------------
Sub SapSleep(ms)
    On Error Resume Next
    If IsObject(WScript) Then WScript.Sleep ms
    On Error GoTo 0
End Sub
