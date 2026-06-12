' TEST rapido - corre con: cscript test_sap.vbs
Option Explicit

WScript.Echo "1. Iniciando test..."

' Test SAP conexion
Dim SapGuiAuto, sapApp, sapConn, sapSess
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: No se encontro SAP GUI - " & Err.Description
    WScript.Quit
End If
WScript.Echo "2. SAP GUI encontrado OK"

Set sapApp = SapGuiAuto.GetScriptingEngine
If Err.Number <> 0 Then
    WScript.Echo "ERROR: No se pudo obtener ScriptingEngine - " & Err.Description
    WScript.Echo "   >> Verifique que el scripting este habilitado en SAP:"
    WScript.Echo "      Menu SAP -> Customizing -> Options -> Accessibility -> Enable Scripting"
    WScript.Quit
End If
WScript.Echo "3. ScriptingEngine OK"

Set sapConn = sapApp.Children(0)
If Err.Number <> 0 Then
    WScript.Echo "ERROR: No hay conexion SAP activa - " & Err.Description
    WScript.Quit
End If
WScript.Echo "4. Conexion SAP OK"

Set sapSess = sapConn.Children(0)
If Err.Number <> 0 Then
    WScript.Echo "ERROR: No hay sesion SAP activa - " & Err.Description
    WScript.Quit
End If
On Error GoTo 0
WScript.Echo "5. Sesion SAP OK"

' Info de la sesion
Dim sInfo
On Error Resume Next
sInfo = sapSess.Info.SystemName & " | Usuario: " & sapSess.Info.User & " | Trans: " & sapSess.Info.Transaction
On Error GoTo 0
WScript.Echo "6. Info sesion: " & sInfo

' Test parseo de texto (simula lo que hace el HTA)
WScript.Echo ""
WScript.Echo "7. Test parseo de datos..."
Dim sTexto, aLineas, i, cols
sTexto = "501055642" & vbTab & "9403" & vbTab & "303047460" & vbCrLf & _
         "501134712" & vbTab & "9403" & vbTab & "303047460"
aLineas = Split(sTexto, vbLf)
For i = 0 To UBound(aLineas)
    Dim sL : sL = Trim(Replace(aLineas(i), vbCr, ""))
    If sL <> "" Then
        cols = Split(sL, vbTab)
        If UBound(cols) >= 2 Then
            WScript.Echo "   Fila " & (i+1) & ": Mat=[" & cols(0) & "] Pos=[" & cols(1) & "] Comp=[" & cols(2) & "]"
        Else
            WScript.Echo "   Fila " & (i+1) & ": MENOS DE 3 COLUMNAS - cols=" & (UBound(cols)+1)
        End If
    End If
Next

WScript.Echo ""
WScript.Echo "=== TODO OK - SAP conectado y parseo funciona ==="
