' ===================================================
' AGREGAR_MACROS.VBS
' Sistema Comparador de Compras Inteligente IA
' Versión: 4.0.0 - Profesional
' ===================================================

Option Explicit

' ===================================================
' CONFIGURACIÓN GLOBAL
' ===================================================

' Variables globales
Dim fso, shell, excelApp, excelWorkbook
Dim projectPath, excelPath, logPath, backupPath
Dim scriptVersion, startTime
Dim errorCount, warningCount, successCount

' ===================================================
' CONSTANTES Y CONFIGURACIONES
' ===================================================
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const TristateTrue = -1

scriptVersion = "4.0.0"
startTime = Now()
errorCount = 0
warningCount = 0
successCount = 0

' ===================================================
' FUNCIONES DE UTILIDAD
' ===================================================

Sub Initialize()
    ' Inicializar objetos principales
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    
    ' Determinar rutas
    projectPath = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName)) & "\Comparador_Compras_IA"
    excelPath = projectPath & "\Comparador_Compras_IA_Completo.xlsm"
    logPath = projectPath & "\Logs\macros_" & FormatDateTime(Now(), 2) & "_" & Replace(FormatDateTime(Now(), 4), ":", "") & ".log"
    backupPath = projectPath & "\Data_Backup\"
    
    ' Crear directorio de logs si no existe
    If Not fso.FolderExists(projectPath & "\Logs") Then
        fso.CreateFolder projectPath & "\Logs"
    End If
    
    ' Crear directorio de backup si no existe
    If Not fso.FolderExists(backupPath) Then
        fso.CreateFolder backupPath
    End If
End Sub

Sub WriteLog(message, logType)
    Dim logFile, logStream, timestamp
    timestamp = FormatDateTime(Now(), 0)
    
    Select Case UCase(logType)
        Case "ERROR"
            logType = "ERROR"
            errorCount = errorCount + 1
        Case "WARNING"
            logType = "WARNING"
            warningCount = warningCount + 1
        Case "SUCCESS"
            logType = "SUCCESS"
            successCount = successCount + 1
        Case Else
            logType = "INFO"
    End Select
    
    ' Crear entrada de log
    Dim logEntry
    logEntry = timestamp & " [" & logType & "] " & message
    
    ' Escribir en archivo de log
    On Error Resume Next
    Set logStream = fso.OpenTextFile(logPath, ForAppending, True)
    logStream.WriteLine logEntry
    logStream.Close
    Set logStream = Nothing
    On Error GoTo 0
    
    ' Mostrar en consola
    If InStr(1, WScript.FullName, "cscript.exe", vbTextCompare) > 0 Then
        WScript.Echo logEntry
    End If
End Sub

Function ExcelInstalled()
    On Error Resume Next
    Dim testExcel
    Set testExcel = CreateObject("Excel.Application")
    If Err.Number = 0 Then
        ExcelInstalled = True
        testExcel.Quit
        Set testExcel = Nothing
    Else
        ExcelInstalled = False
    End If
    On Error GoTo 0
End Function

Function FileExists(filePath)
    On Error Resume Next
    FileExists = fso.FileExists(filePath)
    On Error GoTo 0
End Function

Sub CreateBackup()
    If FileExists(excelPath) Then
        Dim backupName
        backupName = "backup_pre_macros_" & Replace(FormatDateTime(Now(), 2), "/", "") & "_" & Replace(FormatDateTime(Now(), 4), ":", "") & ".xlsm"
        
        On Error Resume Next
        fso.CopyFile excelPath, backupPath & backupName, True
        If Err.Number = 0 Then
            WriteLog "Backup creado: " & backupName, "SUCCESS"
        Else
            WriteLog "No se pudo crear backup: " & Err.Description, "WARNING"
        End If
        On Error GoTo 0
    End If
End Sub

' ===================================================
' FUNCIONES DE CREACIÓN DE CÓDIGO VBA
' ===================================================


Function CreateMainModule()
    Dim moduleCode
    
    moduleCode = "Option Explicit" & vbCrLf & vbCrLf & _
    "' ===================================================" & vbCrLf & _
    "' MÓDULO PRINCIPAL - SISTEMA COMPARADOR DE COMPRAS IA" & vbCrLf & _
    "' Versión: " & scriptVersion & vbCrLf & _
    "' ===================================================" & vbCrLf & vbCrLf & _
    CreateConstantsSection() & vbCrLf & _
    CreateGlobalVariables() & vbCrLf & _
    CreateInitializationFunctions() & vbCrLf & _
    CreateBackupFunctions() & vbCrLf & _
    CreateComparisonFunctions() & vbCrLf & _
    CreateImportExportFunctions() & vbCrLf & _
    CreateUtilityFunctions() & vbCrLf & _
    CreateInterfaceFunctions() & vbCrLf & _
    CreateAutomationFunctions()
    
    CreateMainModule = moduleCode
End Function

Function CreateConstantsSection()
    Dim code
    code = "' CONSTANTES DEL SISTEMA" & vbCrLf & _
    "Public Const SISTEMA_VERSION As String = """ & scriptVersion & """" & vbCrLf & _
    "Public Const SISTEMA_NOMBRE As String = ""Sistema Comparador de Compras Inteligente IA""" & vbCrLf & _
    "Public Const SISTEMA_AUTOR As String = ""Equipo de Desarrollo IA""" & vbCrLf & vbCrLf & _
    "' Códigos de error" & vbCrLf & _
    "Public Const ERR_SUCCESS As Long = 0" & vbCrLf & _
    "Public Const ERR_FILE_NOT_FOUND As Long = 1" & vbCrLf & _
    "Public Const ERR_INVALID_DATA As Long = 2" & vbCrLf & _
    "Public Const ERR_DATABASE As Long = 3" & vbCrLf & _
    "Public Const ERR_CALCULATION As Long = 4" & vbCrLf & vbCrLf & _
    "' Configuración" & vbCrLf & _
    "Public Const MAX_RECORDS As Long = 100000" & vbCrLf & _
    "Public Const MAX_USERS As Long = 1000" & vbCrLf & _
    "Public Const MAX_PRODUCTS As Long = 50000" & vbCrLf & _
    "Public Const MAX_STORES As Long = 10000" & vbCrLf & _
    "Public Const EARTH_RADIUS_KM As Double = 6371.0" & vbCrLf & _
    "Public Const PI As Double = 3.14159265358979"
    
    CreateConstantsSection = code
End Function

Function CreateGlobalVariables()
    Dim code
    code = "' VARIABLES GLOBALES" & vbCrLf & _
    "Public gUsuarioActual As String" & vbCrLf & _
    "Public gRutaProyecto As String" & vbCrLf & _
    "Public gConfigLoaded As Boolean" & vbCrLf & _
    "Public gLastError As String" & vbCrLf & _
    "Public gLastErrorCode As Long" & vbCrLf & _
    "Public gSystemInitialized As Boolean" & vbCrLf & _
    "Public gLogEnabled As Boolean" & vbCrLf & _
    "Public gBackupEnabled As Boolean"
    
    CreateGlobalVariables = code
End Function

Function CreateInitializationFunctions()
    Dim code
    code = "' ===================================================" & vbCrLf & _
    "' FUNCIONES DE INICIALIZACIÓN" & vbCrLf & _
    "' ===================================================" & vbCrLf & vbCrLf & _
    "Public Sub InicializarSistema()" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & vbCrLf & _
    "    ' Configurar variables globales" & vbCrLf & _
    "    gUsuarioActual = """"" & vbCrLf & _
    "    gRutaProyecto = ThisWorkbook.Path" & vbCrLf & _
    "    gConfigLoaded = False" & vbCrLf & _
    "    gLastError = """"" & vbCrLf & _
    "    gLastErrorCode = ERR_SUCCESS" & vbCrLf & _
    "    gSystemInitialized = False" & vbCrLf & _
    "    gLogEnabled = True" & vbCrLf & _
    "    gBackupEnabled = True" & vbCrLf & vbCrLf & _
    "    ' Verificar estructura" & vbCrLf & _
    "    If Not VerificarEstructura() Then" & vbCrLf & _
    "        MsgBox ""Error en la estructura del sistema. Verifique las hojas."", vbCritical" & vbCrLf & _
    "        Exit Sub" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    ' Cargar configuración" & vbCrLf & _
    "    If Not CargarConfiguracion() Then" & vbCrLf & _
    "        MsgBox ""No se pudo cargar la configuración. Se usarán valores por defecto."", vbExclamation" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    ' Crear menú" & vbCrLf & _
    "    CrearMenuPrincipal" & vbCrLf & vbCrLf & _
    "    ' Actualizar estado" & vbCrLf & _
    "    gSystemInitialized = True" & vbCrLf & _
    "    ActualizarBarraEstado ""Sistema inicializado correctamente""" & vbCrLf & vbCrLf & _
    "    ' Mostrar mensaje de bienvenida" & vbCrLf & _
    "    If Application.Version >= 12.0 Then" & vbCrLf & _
    "        MsgBox ""Bienvenido al "" & SISTEMA_NOMBRE & "" v"" & SISTEMA_VERSION & vbCrLf & _" & vbCrLf & _
    "               ""Sistema listo para comparar precios y optimizar compras."", _" & vbCrLf & _
    "               vbInformation, ""Inicialización del Sistema""" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    Exit Sub" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    gLastError = Err.Description" & vbCrLf & _
    "    gLastErrorCode = Err.Number" & vbCrLf & _
    "    MsgBox ""Error durante la inicialización: "" & Err.Description, vbCritical" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Function VerificarEstructura() As Boolean" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & _
    "    Dim hojasRequeridas As Variant" & vbCrLf & _
    "    Dim hoja As Worksheet" & vbCrLf & _
    "    Dim i As Integer" & vbCrLf & _
    "    Dim hojaEncontrada As Boolean" & vbCrLf & vbCrLf & _
    "    hojasRequeridas = Array(""USUARIOS"", ""PRODUCTOS"", ""TIENDAS"", ""PRECIOS"", _" & vbCrLf & _
    "                           ""COMPARATIVA"", ""HISTORIAL_COMPRAS"", ""PREFERENCIAS_IA"")" & vbCrLf & vbCrLf & _
    "    For i = LBound(hojasRequeridas) To UBound(hojasRequeridas)" & vbCrLf & _
    "        hojaEncontrada = False" & vbCrLf & _
    "        For Each hoja In ThisWorkbook.Worksheets" & vbCrLf & _
    "            If hoja.Name = hojasRequeridas(i) Then" & vbCrLf & _
    "                hojaEncontrada = True" & vbCrLf & _
    "                Exit For" & vbCrLf & _
    "            End If" & vbCrLf & _
    "        Next hoja" & vbCrLf & vbCrLf & _
    "        If Not hojaEncontrada Then" & vbCrLf & _
    "            VerificarEstructura = False" & vbCrLf & _
    "            Exit Function" & vbCrLf & _
    "        End If" & vbCrLf & _
    "    Next i" & vbCrLf & vbCrLf & _
    "    VerificarEstructura = True" & vbCrLf & _
    "    Exit Function" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    VerificarEstructura = False" & vbCrLf & _
    "End Function" & vbCrLf & vbCrLf & _
    "Private Function CargarConfiguracion() As Boolean" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & _
    "    Dim configPath As String" & vbCrLf & _
    "    Dim fso As Object" & vbCrLf & _
    "    Dim ts As Object" & vbCrLf & _
    "    Dim configText As String" & vbCrLf & vbCrLf & _
    "    configPath = gRutaProyecto & ""\Configuraciones\config_sistema.json""" & vbCrLf & _
    "    Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & vbCrLf & _
    "    If Not fso.FileExists(configPath) Then" & vbCrLf & _
    "        CargarConfiguracion = False" & vbCrLf & _
    "        Exit Function" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    Set ts = fso.OpenTextFile(configPath, ForReading)" & vbCrLf & _
    "    configText = ts.ReadAll" & vbCrLf & _
    "    ts.Close" & vbCrLf & vbCrLf & _
    "    ' Aquí se procesaría el JSON (simplificado)" & vbCrLf & _
    "    gConfigLoaded = True" & vbCrLf & _
    "    CargarConfiguracion = True" & vbCrLf & vbCrLf & _
    "    Exit Function" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    CargarConfiguracion = False" & vbCrLf & _
    "End Function"
    
    CreateInitializationFunctions = code
End Function

Function CreateImportExportFunctions()
    Dim code, dq
    dq = Chr(34) ' Definimos comilla doble una vez
    
    code = "' ===================================================" & vbCrLf & _
    "' FUNCIONES DE IMPORTACIÓN/EXPORTACIÓN" & vbCrLf & _
    "' ===================================================" & vbCrLf & vbCrLf & _
    "Public Sub ImportarDatosCSV()" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & _
    "    Dim rutaArchivo As String" & vbCrLf & _
    "    Dim hojaDestino As String" & vbCrLf & vbCrLf & _
    "    ' Seleccionar archivo" & vbCrLf & _
    "    With Application.FileDialog(msoFileDialogFilePicker)" & vbCrLf & _
    "        .Title = " & dq & "Seleccionar archivo CSV para importar" & dq & vbCrLf & _
    "        .Filters.Clear" & vbCrLf & _
    "        .Filters.Add " & dq & "Archivos CSV" & dq & ", " & dq & "*.csv" & dq & vbCrLf & _
    "        .AllowMultiSelect = False" & vbCrLf & vbCrLf & _
    "        If .Show = -1 Then" & vbCrLf & _
    "            rutaArchivo = .SelectedItems(1)" & vbCrLf & _
    "        Else" & vbCrLf & _
    "            Exit Sub" & vbCrLf & _
    "        End If" & vbCrLf & _
    "    End With" & vbCrLf & vbCrLf & _
    "    ' Seleccionar hoja destino" & vbCrLf & _
    "    hojaDestino = InputBox(" & dq & "Ingrese el nombre de la hoja destino:" & dq & ", " & dq & "Importar Datos" & dq & ")" & vbCrLf & _
    "    If hojaDestino = " & dq & dq & " Then Exit Sub" & vbCrLf & vbCrLf & _
    "    ' Importar datos" & vbCrLf & _
    "    If ImportarCSV(rutaArchivo, hojaDestino) Then" & vbCrLf & _
    "        MsgBox " & dq & "Datos importados exitosamente." & dq & ", vbInformation" & vbCrLf & _
    "        ActualizarBarraEstado " & dq & "Datos importados desde CSV" & dq & vbCrLf & _
    "    Else" & vbCrLf & _
    "        MsgBox " & dq & "Error al importar datos." & dq & ", vbCritical" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    Exit Sub" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    MsgBox " & dq & "Error durante la importación: " & dq & " & Err.Description, vbCritical" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Function ImportarCSV(rutaArchivo As String, hojaDestino As String) As Boolean" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & _
    "    Dim ws As Worksheet" & vbCrLf & _
    "    Dim fso As Object, ts As Object" & vbCrLf & _
    "    Dim lineas() As String, campos() As String" & vbCrLf & _
    "    Dim i As Long, j As Long" & vbCrLf & _
    "    Dim textoArchivo As String" & vbCrLf & vbCrLf & _
    "    ' Verificar que la hoja existe" & vbCrLf & _
    "    On Error Resume Next" & vbCrLf & _
    "    Set ws = ThisWorkbook.Worksheets(hojaDestino)" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & vbCrLf & _
    "    If ws Is Nothing Then" & vbCrLf & _
    "        MsgBox " & dq & "La hoja especificada no existe." & dq & ", vbExclamation" & vbCrLf & _
    "        ImportarCSV = False" & vbCrLf & _
    "        Exit Function" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    ' Limpiar datos existentes (excepto cabeceras)" & vbCrLf & _
    "    If ws.UsedRange.Rows.Count > 1 Then" & vbCrLf & _
    "        ws.Range(ws.Cells(2, 1), ws.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count)).ClearContents" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    ' Leer archivo CSV" & vbCrLf & _
    "    Set fso = CreateObject(" & dq & "Scripting.FileSystemObject" & dq & ")" & vbCrLf & _
    "    Set ts = fso.OpenTextFile(rutaArchivo, ForReading, TristateTrue)" & vbCrLf & _
    "    textoArchivo = ts.ReadAll" & vbCrLf & _
    "    ts.Close" & vbCrLf & vbCrLf & _
    "    ' Procesar líneas" & vbCrLf & _
    "    lineas = Split(textoArchivo, vbCrLf)" & vbCrLf & _
    "    For i = 1 To UBound(lineas)" & vbCrLf & _
    "        If Trim(lineas(i)) <> " & dq & dq & " Then" & vbCrLf & _
    "            campos = Split(lineas(i), " & dq & "," & dq & ")" & vbCrLf & _
    "            For j = 0 To UBound(campos)" & vbCrLf & _
    "                ws.Cells(i + 1, j + 1).Value = campos(j)" & vbCrLf & _
    "            Next j" & vbCrLf & _
    "        End If" & vbCrLf & _
    "    Next i" & vbCrLf & vbCrLf & _
    "    ImportarCSV = True" & vbCrLf & _
    "    Exit Function" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    ImportarCSV = False" & vbCrLf & _
    "End Function" & vbCrLf & vbCrLf & _
    "Public Sub ExportarDatosCSV()" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & _
    "    Dim hojaOrigen As String" & vbCrLf & _
    "    Dim rutaArchivo As String" & vbCrLf & vbCrLf & _
    "    ' Seleccionar hoja" & vbCrLf & _
    "    hojaOrigen = InputBox(" & dq & "Ingrese el nombre de la hoja a exportar:" & dq & ", " & dq & "Exportar Datos" & dq & ")" & vbCrLf & _
    "    If hojaOrigen = " & dq & dq & " Then Exit Sub" & vbCrLf & vbCrLf & _
    "    ' Seleccionar destino" & vbCrLf & _
    "    With Application.FileDialog(msoFileDialogSaveAs)" & vbCrLf & _
    "        .Title = " & dq & "Guardar archivo CSV" & dq & vbCrLf & _
    "        .InitialFileName = hojaOrigen & " & dq & ".csv" & dq & vbCrLf & _
    "        .Filters.Clear" & vbCrLf & _
    "        .Filters.Add " & dq & "Archivos CSV" & dq & ", " & dq & "*.csv" & dq & vbCrLf & _
    "        .AllowMultiSelect = False" & vbCrLf & vbCrLf & _
    "        If .Show = -1 Then" & vbCrLf & _
    "            rutaArchivo = .SelectedItems(1)" & vbCrLf & _
    "            ' Asegurar extensión .csv" & vbCrLf & _
    "            If LCase(Right(rutaArchivo, 4)) <> " & dq & ".csv" & dq & " Then" & vbCrLf & _
    "                rutaArchivo = rutaArchivo & " & dq & ".csv" & dq & vbCrLf & _
    "            End If" & vbCrLf & _
    "        Else" & vbCrLf & _
    "            Exit Sub" & vbCrLf & _
    "        End If" & vbCrLf & _
    "    End With" & vbCrLf & vbCrLf & _
    "    ' Exportar datos" & vbCrLf & _
    "    If ExportarCSV(hojaOrigen, rutaArchivo) Then" & vbCrLf & _
    "        MsgBox " & dq & "Datos exportados exitosamente a " & dq & " & rutaArchivo, vbInformation" & vbCrLf & _
    "        ActualizarBarraEstado " & dq & "Datos exportados a CSV" & dq & vbCrLf & _
    "    Else" & vbCrLf & _
    "        MsgBox " & dq & "Error al exportar datos." & dq & ", vbCritical" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    Exit Sub" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    MsgBox " & dq & "Error durante la exportación: " & dq & " & Err.Description, vbCritical" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Function ExportarCSV(hojaOrigen As String, rutaArchivo As String) As Boolean" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & _
    "    Dim ws As Worksheet" & vbCrLf & _
    "    Dim fso As Object, ts As Object" & vbCrLf & _
    "    Dim lastRow As Long, lastCol As Long" & vbCrLf & _
    "    Dim i As Long, j As Long" & vbCrLf & _
    "    Dim linea As String" & vbCrLf & _
    "    Dim valorCelda As String" & vbCrLf & vbCrLf & _
    "    ' Verificar que la hoja existe" & vbCrLf & _
    "    On Error Resume Next" & vbCrLf & _
    "    Set ws = ThisWorkbook.Worksheets(hojaOrigen)" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & vbCrLf & _
    "    If ws Is Nothing Then" & vbCrLf & _
    "        ExportarCSV = False" & vbCrLf & _
    "        Exit Function" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    ' Determinar rango de datos" & vbCrLf & _
    "    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf & _
    "    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column" & vbCrLf & vbCrLf & _
    "    ' Crear archivo CSV" & vbCrLf & _
    "    Set fso = CreateObject(" & dq & "Scripting.FileSystemObject" & dq & ")" & vbCrLf & _
    "    Set ts = fso.CreateTextFile(rutaArchivo, True)" & vbCrLf & vbCrLf & _
    "    For i = 1 To lastRow" & vbCrLf & _
    "        linea = " & dq & dq & vbCrLf & _
    "        For j = 1 To lastCol" & vbCrLf & _
    "            If j > 1 Then linea = linea & " & dq & "," & dq & vbCrLf & _
    "            ' Obtener valor de la celda y escapar comillas" & vbCrLf & _
    "            valorCelda = CStr(ws.Cells(i, j).Value)" & vbCrLf & _
    "            valorCelda = Replace(valorCelda, " & dq & dq & dq & dq & ", " & dq & dq & dq & dq & dq & dq & ")" & vbCrLf & _
    "            linea = linea & " & dq & dq & " & valorCelda & " & dq & dq & vbCrLf & _
    "        Next j" & vbCrLf & _
    "        ts.WriteLine linea" & vbCrLf & _
    "    Next i" & vbCrLf & vbCrLf & _
    "    ts.Close" & vbCrLf & vbCrLf & _
    "    ExportarCSV = True" & vbCrLf & _
    "    Exit Function" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    ExportarCSV = False" & vbCrLf & _
    "End Function"
    
    CreateImportExportFunctions = code
End Function

Function CreateBackupFunctions()
    Dim code
    code = "' ===================================================" & vbCrLf & _
    "' FUNCIONES DE BACKUP Y SEGURIDAD" & vbCrLf & _
    "' ===================================================" & vbCrLf & vbCrLf & _
    "Public Sub CrearBackupCompleto()" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & vbCrLf & _
    "    Dim backupPath As String" & vbCrLf & _
    "    Dim backupName As String" & vbCrLf & vbCrLf & _
    "    If Not gBackupEnabled Then" & vbCrLf & _
    "        MsgBox ""El sistema de backup está deshabilitado."", vbExclamation" & vbCrLf & _
    "        Exit Sub" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    ' Generar nombre único" & vbCrLf & _
    "    backupName = ""backup_completo_"" & Format(Now(), ""yyyymmdd_hhnnss"") & "".xlsm""" & vbCrLf & _
    "    backupPath = gRutaProyecto & ""\Data_Backup\"" & backupName" & vbCrLf & vbCrLf & _
    "    ' Crear copia" & vbCrLf & _
    "    ThisWorkbook.SaveCopyAs backupPath" & vbCrLf & vbCrLf & _
    "    MsgBox ""Backup creado exitosamente en:"" & vbCrLf & backupPath, vbInformation" & vbCrLf & _
    "    ActualizarBarraEstado ""Backup completado""" & vbCrLf & vbCrLf & _
    "    Exit Sub" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    MsgBox ""Error al crear backup: "" & Err.Description, vbCritical" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Public Sub RestaurarDesdeBackup()" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & vbCrLf & _
    "    Dim rutaArchivo As String" & vbCrLf & _
    "    Dim respuesta As Integer" & vbCrLf & vbCrLf & _
    "    ' Advertencia" & vbCrLf & _
    "    respuesta = MsgBox(""ADVERTENCIA: Restaurar desde backup sobrescribirá todos los datos actuales."" & vbCrLf & _" & vbCrLf & _
    "               ""¿Está seguro de continuar?"", vbCritical + vbYesNo, ""Confirmar Restauración"")" & vbCrLf & vbCrLf & _
    "    If respuesta <> vbYes Then Exit Sub" & vbCrLf & vbCrLf & _
    "    ' Seleccionar archivo" & vbCrLf & _
    "    With Application.FileDialog(msoFileDialogFilePicker)" & vbCrLf & _
    "        .Title = ""Seleccionar archivo de backup para restaurar""" & vbCrLf & _
    "        .InitialFileName = gRutaProyecto & ""\Data_Backup\""" & vbCrLf & _
    "        .Filters.Clear" & vbCrLf & _
    "        .Filters.Add ""Archivos Excel"", ""*.xls;*.xlsx;*.xlsm""" & vbCrLf & _
    "        .AllowMultiSelect = False" & vbCrLf & vbCrLf & _
    "        If .Show = -1 Then" & vbCrLf & _
    "            rutaArchivo = .SelectedItems(1)" & vbCrLf & _
    "        Else" & vbCrLf & _
    "            Exit Sub" & vbCrLf & _
    "        End If" & vbCrLf & _
    "    End With" & vbCrLf & vbCrLf & _
    "    ' Restaurar datos" & vbCrLf & _
    "    If RestaurarBackup(rutaArchivo) Then" & vbCrLf & _
    "        MsgBox ""Datos restaurados exitosamente desde backup."", vbInformation" & vbCrLf & _
    "        ActualizarBarraEstado ""Sistema restaurado desde backup""" & vbCrLf & _
    "    Else" & vbCrLf & _
    "        MsgBox ""Error al restaurar datos desde backup."", vbCritical" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    Exit Sub" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    MsgBox ""Error durante la restauración: "" & Err.Description, vbCritical" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Function RestaurarBackup(rutaArchivo As String) As Boolean" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & _
    "    Dim backupWorkbook As Workbook" & vbCrLf & _
    "    Dim wsBackup As Worksheet, wsActual As Worksheet" & vbCrLf & _
    "    Dim nombreHoja As String" & vbCrLf & vbCrLf & _
    "    ' Abrir backup" & vbCrLf & _
    "    Set backupWorkbook = Workbooks.Open(rutaArchivo, ReadOnly:=True)" & vbCrLf & vbCrLf & _
    "    ' Copiar datos de cada hoja" & vbCrLf & _
    "    For Each wsBackup In backupWorkbook.Worksheets" & vbCrLf & _
    "        nombreHoja = wsBackup.Name" & vbCrLf & vbCrLf & _
    "        On Error Resume Next" & vbCrLf & _
    "        Set wsActual = ThisWorkbook.Worksheets(nombreHoja)" & vbCrLf & _
    "        On Error GoTo ErrorHandler" & vbCrLf & vbCrLf & _
    "        If Not wsActual Is Nothing Then" & vbCrLf & _
    "            ' Limpiar hoja actual" & vbCrLf & _
    "            wsActual.Cells.ClearContents" & vbCrLf & _
    "            ' Copiar datos" & vbCrLf & _
    "            wsBackup.UsedRange.Copy wsActual.Range(""A1"")" & vbCrLf & _
    "        End If" & vbCrLf & _
    "    Next wsBackup" & vbCrLf & vbCrLf & _
    "    ' Cerrar backup" & vbCrLf & _
    "    backupWorkbook.Close SaveChanges:=False" & vbCrLf & vbCrLf & _
    "    RestaurarBackup = True" & vbCrLf & _
    "    Exit Function" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    RestaurarBackup = False" & vbCrLf & _
    "End Function"
    
    CreateBackupFunctions = code
End Function

Function CreateComparisonFunctions()
    Dim code
    code = "' ===================================================" & vbCrLf & _
    "' FUNCIONES DE COMPARACIГғвҖңN" & vbCrLf & _
    "' ===================================================" & vbCrLf & vbCrLf & _
    "Public Sub CompararProducto()" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & vbCrLf & _
    "    Dim nombreProducto As String, categoria As String" & vbCrLf & _
    "    Dim wsProductos As Worksheet, wsPrecios As Worksheet, wsTiendas As Worksheet" & vbCrLf & _
    "    Dim rngProductos As Range, celda As Range" & vbCrLf & _
    "    Dim idProducto As String, precioMasBarato As Double, tiendaMasBarata As String" & vbCrLf & _
    "    Dim resultado As String, contador As Integer" & vbCrLf & _
    "    Dim arrResultados() As Variant, i As Integer" & vbCrLf & _
    "    Dim distancia As Double, latUsuario As Double, lonUsuario As Double" & vbCrLf & vbCrLf & _
    "    ' Solicitar datos del producto" & vbCrLf & _
    "    nombreProducto = InputBox(""Ingrese el nombre del producto a comparar:"", ""Comparar Producto"")" & vbCrLf & _
    "    If nombreProducto = """" Then Exit Sub" & vbCrLf & _
    "    categoria = InputBox(""Ingrese la categorГғВӯa del producto (opcional):"", ""Comparar Producto"")" & vbCrLf & vbCrLf & _
    "    ' Obtener coordenadas del usuario (opcional)" & vbCrLf & _
    "    On Error Resume Next" & vbCrLf & _
    "    latUsuario = CDbl(InputBox(""Ingrese su latitud (opcional, para calcular distancias):"", ""UbicaciГғВіn"", ""0""))" & vbCrLf & _
    "    lonUsuario = CDbl(InputBox(""Ingrese su longitud (opcional):"", ""UbicaciГғВіn"", ""0""))" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & vbCrLf & _
    "    ' Configurar hojas" & vbCrLf & _
    "    Set wsProductos = ThisWorkbook.Sheets(""PRODUCTOS"")" & vbCrLf & _
    "    Set wsPrecios = ThisWorkbook.Sheets(""PRECIOS"")" & vbCrLf & _
    "    Set wsTiendas = ThisWorkbook.Sheets(""TIENDAS"")" & vbCrLf & vbCrLf & _
    "    ' Buscar producto en la base de datos" & vbCrLf & _
    "    Set rngProductos = wsProductos.Range(""A2:A"" & wsProductos.Cells(wsProductos.Rows.Count, 1).End(xlUp).Row)" & vbCrLf & _
    "    idProducto = """"" & vbCrLf & _
    "    For Each celda In rngProductos" & vbCrLf & _
    "        If InStr(1, wsProductos.Cells(celda.Row, 2).Value, nombreProducto, vbTextCompare) > 0 Then" & vbCrLf & _
    "            If categoria = """" Or LCase(wsProductos.Cells(celda.Row, 3).Value) = LCase(categoria) Then" & vbCrLf & _
    "                idProducto = celda.Value" & vbCrLf & _
    "                Exit For" & vbCrLf & _
    "            End If" & vbCrLf & _
    "        End If" & vbCrLf & _
    "    Next celda" & vbCrLf & vbCrLf & _
    "    If idProducto = """" Then" & vbCrLf & _
    "        MsgBox ""Producto no encontrado en la base de datos."", vbExclamation" & vbCrLf & _
    "        Exit Sub" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    ' Buscar precios para el producto" & vbCrLf & _
    "    ReDim arrResultados(1 To 100, 1 To 6) ' ID, Tienda, Precio, Descuento, Distancia, Precio/Unidad" & vbCrLf & _
    "    contador = 0" & vbCrLf & _
    "    precioMasBarato = 9999999" & vbCrLf & _
    "    tiendaMasBarata = """"" & vbCrLf & vbCrLf & _
    "    Dim rngPrecios As Range, celdaPrecio As Range" & vbCrLf & _
    "    Set rngPrecios = wsPrecios.Range(""A2:A"" & wsPrecios.Cells(wsPrecios.Rows.Count, 1).End(xlUp).Row)" & vbCrLf & vbCrLf & _
    "    For Each celdaPrecio In rngPrecios" & vbCrLf & _
    "        If celdaPrecio.Value = idProducto Then" & vbCrLf & _
    "            Dim idTienda As String, precio As Double, descuento As Double" & vbCrLf & _
    "            Dim nombreTienda As String, latTienda As Double, lonTienda As Double" & vbCrLf & _
    "            Dim precioPorUnidad As Double, unidad As String" & vbCrLf & vbCrLf & _
    "            idTienda = wsPrecios.Cells(celdaPrecio.Row, 2).Value" & vbCrLf & _
    "            precio = wsPrecios.Cells(celdaPrecio.Row, 3).Value" & vbCrLf & _
    "            descuento = wsPrecios.Cells(celdaPrecio.Row, 4).Value" & vbCrLf & _
    "            unidad = wsPrecios.Cells(celdaPrecio.Row, 5).Value" & vbCrLf & vbCrLf & _
    "            ' Buscar informaciГғВіn de la tienda" & vbCrLf & _
    "            Dim rngTiendas As Range, celdaTienda As Range" & vbCrLf & _
    "            Set rngTiendas = wsTiendas.Range(""A2:A"" & wsTiendas.Cells(wsTiendas.Rows.Count, 1).End(xlUp).Row)" & vbCrLf & vbCrLf & _
    "            For Each celdaTienda In rngTiendas" & vbCrLf & _
    "                If celdaTienda.Value = idTienda Then" & vbCrLf & _
    "                    nombreTienda = wsTiendas.Cells(celdaTienda.Row, 2).Value" & vbCrLf & _
    "                    latTienda = wsTiendas.Cells(celdaTienda.Row, 6).Value" & vbCrLf & _
    "                    lonTienda = wsTiendas.Cells(celdaTienda.Row, 7).Value" & vbCrLf & _
    "                    Exit For" & vbCrLf & _
    "                End If" & vbCrLf & _
    "            Next celdaTienda" & vbCrLf & vbCrLf & _
    "            ' Calcular distancia si se proporcionaron coordenadas" & vbCrLf & _
    "            If latUsuario <> 0 And lonUsuario <> 0 Then" & vbCrLf & _
    "                distancia = CalcularDistanciaHaversine(latUsuario, lonUsuario, latTienda, lonTienda)" & vbCrLf & _
    "            Else" & vbCrLf & _
    "                distancia = 0" & vbCrLf & _
    "            End If" & vbCrLf & vbCrLf & _
    "            ' Calcular precio por unidad" & vbCrLf & _
    "            If unidad <> """" Then" & vbCrLf & _
    "                precioPorUnidad = CalcularPrecioPorUnidad(precio, 1, unidad)" & vbCrLf & _
    "            Else" & vbCrLf & _
    "                precioPorUnidad = precio" & vbCrLf & _
    "            End If" & vbCrLf & vbCrLf & _
    "            ' Almacenar resultado" & vbCrLf & _
    "            contador = contador + 1" & vbCrLf & _
    "            arrResultados(contador, 1) = idTienda" & vbCrLf & _
    "            arrResultados(contador, 2) = nombreTienda" & vbCrLf & _
    "            arrResultados(contador, 3) = precio" & vbCrLf & _
    "            arrResultados(contador, 4) = descuento" & vbCrLf & _
    "            arrResultados(contador, 5) = distancia" & vbCrLf & _
    "            arrResultados(contador, 6) = precioPorUnidad" & vbCrLf & vbCrLf & _
    "            ' Verificar si es el precio mГғВЎs barato" & vbCrLf & _
    "            If precio < precioMasBarato Then" & vbCrLf & _
    "                precioMasBarato = precio" & vbCrLf & _
    "                tiendaMasBarata = nombreTienda" & vbCrLf & _
    "            End If" & vbCrLf & _
    "        End If" & vbCrLf & _
    "    Next celdaPrecio" & vbCrLf & vbCrLf & _
    "    ' Mostrar resultados" & vbCrLf & _
    "    If contador > 0 Then" & vbCrLf & _
    "        resultado = ""COMPARATIVA DE PRECIOS"" & vbCrLf & vbCrLf" & vbCrLf & _
    "        resultado = resultado & ""Producto: "" & nombreProducto & vbCrLf" & vbCrLf & _
    "        resultado = resultado & ""ID Producto: "" & idProducto & vbCrLf & vbCrLf" & vbCrLf & _
    "        resultado = resultado & ""Se encontraron "" & contador & "" precios diferentes:"" & vbCrLf & vbCrLf" & vbCrLf & _
    "        resultado = resultado & String(50, ""-"") & vbCrLf" & vbCrLf & _
    "        For i = 1 To contador" & vbCrLf & _
    "            resultado = resultado & i & "". "" & arrResultados(i, 2) & vbCrLf" & vbCrLf & _
    "            resultado = resultado & ""   Precio: "" & FormatearMoneda(arrResultados(i, 3)) & vbCrLf" & vbCrLf & _
    "            If arrResultados(i, 4) > 0 Then" & vbCrLf & _
    "                resultado = resultado & ""   Descuento: "" & arrResultados(i, 4) & ""%"" & vbCrLf" & vbCrLf & _
    "            End If" & vbCrLf & _
    "            If arrResultados(i, 5) > 0 Then" & vbCrLf & _
    "                resultado = resultado & ""   Distancia: "" & Format(arrResultados(i, 5), ""0.0"") & "" km"" & vbCrLf" & vbCrLf & _
    "            End If" & vbCrLf & _
    "            resultado = resultado & ""   Precio por unidad: "" & FormatearMoneda(arrResultados(i, 6)) & vbCrLf & vbCrLf" & vbCrLf & _
    "            resultado = resultado & String(50, ""-"") & vbCrLf" & vbCrLf & _
    "        Next i" & vbCrLf & vbCrLf & _
    "        resultado = resultado & vbCrLf & ""MEJOR PRECIO: "" & FormatearMoneda(precioMasBarato) & "" en "" & tiendaMasBarata & vbCrLf" & vbCrLf & _
    "        If contador > 1 Then" & vbCrLf & _
    "            Dim ahorroPromedio As Double" & vbCrLf & _
    "            Dim precioMasCaro As Double" & vbCrLf & _
    "            precioMasCaro = 0" & vbCrLf & _
    "            For i = 1 To contador" & vbCrLf & _
    "                If arrResultados(i, 3) > precioMasCaro Then precioMasCaro = arrResultados(i, 3)" & vbCrLf & _
    "            Next i" & vbCrLf & _
    "            ahorroPromedio = CalcularAhorroPorcentual(precioMasCaro, precioMasBarato)" & vbCrLf & _
    "            resultado = resultado & ""Ahorro mГғВЎximo posible: "" & Format(ahorroPromedio, ""0.0"") & ""%"" & vbCrLf" & vbCrLf & _
    "        End If" & vbCrLf & vbCrLf & _
    "        ' Mostrar en cuadro de mensaje" & vbCrLf & _
    "        MsgBox resultado, vbInformation, ""Comparativa de Precios""" & vbCrLf & vbCrLf & _
    "        ' Registrar en hoja COMPARATIVA" & vbCrLf & _
    "        Dim wsComparativa As Worksheet" & vbCrLf & _
    "        Set wsComparativa = ThisWorkbook.Sheets(""COMPARATIVA"")" & vbCrLf & _
    "        Dim ultimaFila As Long" & vbCrLf & _
    "        ultimaFila = wsComparativa.Cells(wsComparativa.Rows.Count, 1).End(xlUp).Row + 1" & vbCrLf & _
    "        wsComparativa.Cells(ultimaFila, 1).Value = Now()" & vbCrLf & _
    "        wsComparativa.Cells(ultimaFila, 2).Value = idProducto" & vbCrLf & _
    "        wsComparativa.Cells(ultimaFila, 3).Value = nombreProducto" & vbCrLf & _
    "        wsComparativa.Cells(ultimaFila, 4).Value = contador" & vbCrLf & _
    "        wsComparativa.Cells(ultimaFila, 5).Value = precioMasBarato" & vbCrLf & _
    "        wsComparativa.Cells(ultimaFila, 6).Value = tiendaMasBarata" & vbCrLf & _
    "        wsComparativa.Cells(ultimaFila, 7).Value = gUsuarioActual" & vbCrLf & vbCrLf & _
    "        ActualizarBarraEstado ""Comparativa completada para "" & nombreProducto" & vbCrLf & _
    "    Else" & vbCrLf & _
    "        MsgBox ""No se encontraron precios para el producto especificado."", vbExclamation" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    Exit Sub" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    MsgBox ""Error en la comparaciГғВіn: "" & Err.Description, vbCritical" & vbCrLf & _
    "End Sub"
    
    CreateComparisonFunctions = code
End Function

Function CreateUtilityFunctions()
    Dim code
    code = "' ===================================================" & vbCrLf & _
    "' FUNCIONES DE UTILIDAD" & vbCrLf & _
    "' ===================================================" & vbCrLf & vbCrLf & _
    "Public Sub ActualizarBarraEstado(mensaje As String)" & vbCrLf & _
    "    Application.StatusBar = SISTEMA_NOMBRE & "" - "" & mensaje & "" - "" & Format(Now(), ""dd/mm/yyyy HH:mm:ss"")" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Public Sub LimpiarBarraEstado()" & vbCrLf & _
    "    Application.StatusBar = False" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Public Sub LimpiarDatos()" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & _
    "    Dim respuesta As Integer" & vbCrLf & _
    "    Dim ws As Worksheet" & vbCrLf & vbCrLf & _
    "    ' Confirmación" & vbCrLf & _
    "    respuesta = MsgBox(""ADVERTENCIA: Esta acción eliminará todos los datos (excepto cabeceras) de todas las hojas."" & vbCrLf & _" & vbCrLf & _
    "               ""¿Está seguro de continuar?"", vbCritical + vbYesNo, ""Confirmar Limpieza"")" & vbCrLf & vbCrLf & _
    "    If respuesta <> vbYes Then Exit Sub" & vbCrLf & vbCrLf & _
    "    ' Limpiar cada hoja" & vbCrLf & _
    "    For Each ws In ThisWorkbook.Worksheets" & vbCrLf & _
    "        If ws.UsedRange.Rows.Count > 1 Then" & vbCrLf & _
    "            ws.Range(ws.Cells(2, 1), ws.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count)).ClearContents" & vbCrLf & _
    "        End If" & vbCrLf & _
    "    Next ws" & vbCrLf & vbCrLf & _
    "    MsgBox ""Todos los datos han sido limpiados."", vbInformation" & vbCrLf & _
    "    ActualizarBarraEstado ""Datos limpiados""" & vbCrLf & vbCrLf & _
    "    Exit Sub" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    MsgBox ""Error durante la limpieza: "" & Err.Description, vbCritical" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Public Sub GenerarReporteSimple()" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & _
    "    Dim wsUsuarios As Worksheet, wsCompras As Worksheet" & vbCrLf & _
    "    Dim totalUsuarios As Long, totalCompras As Long" & vbCrLf & _
    "    Dim gastoTotal As Double, gastoPromedio As Double" & vbCrLf & _
    "    Dim mensaje As String" & vbCrLf & vbCrLf & _
    "    Set wsUsuarios = ThisWorkbook.Sheets(""USUARIOS"")" & vbCrLf & _
    "    Set wsCompras = ThisWorkbook.Sheets(""HISTORIAL_COMPRAS"")" & vbCrLf & vbCrLf & _
    "    ' Calcular estadísticas" & vbCrLf & _
    "    totalUsuarios = Application.WorksheetFunction.CountA(wsUsuarios.Range(""A:A"")) - 1" & vbCrLf & _
    "    totalCompras = Application.WorksheetFunction.CountA(wsCompras.Range(""A:A"")) - 1" & vbCrLf & _
    "    gastoTotal = Application.WorksheetFunction.Sum(wsCompras.Range(""E:E""))" & vbCrLf & vbCrLf & _
    "    If totalCompras > 0 Then" & vbCrLf & _
    "        gastoPromedio = gastoTotal / totalCompras" & vbCrLf & _
    "    Else" & vbCrLf & _
    "        gastoPromedio = 0" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    ' Construir reporte" & vbCrLf & _
    "    mensaje = ""REPORTE DEL SISTEMA"" & vbCrLf & vbCrLf" & vbCrLf & _
    "    mensaje = mensaje & ""Usuarios registrados: "" & totalUsuarios & vbCrLf" & vbCrLf & _
    "    mensaje = mensaje & ""Compras registradas: "" & totalCompras & vbCrLf" & vbCrLf & _
    "    mensaje = mensaje & ""Gasto total: "" & Format(gastoTotal, ""0.00€"") & vbCrLf" & vbCrLf & _
    "    mensaje = mensaje & ""Gasto promedio por compra: "" & Format(gastoPromedio, ""0.00€"") & vbCrLf & vbCrLf" & vbCrLf & _
    "    mensaje = mensaje & ""Fecha del reporte: "" & Format(Now(), ""dd/mm/yyyy HH:mm:ss"")" & vbCrLf & vbCrLf & _
    "    ' Mostrar reporte" & vbCrLf & _
    "    MsgBox mensaje, vbInformation, ""Reporte del Sistema""" & vbCrLf & vbCrLf & _
    "    Exit Sub" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    MsgBox ""Error al generar reporte: "" & Err.Description, vbCritical" & vbCrLf & _
    "End Sub"
    
    CreateUtilityFunctions = code
End Function

Function CreateInterfaceFunctions()
    Dim code
    code = "' ===================================================" & vbCrLf & _
    "' FUNCIONES DE INTERFAZ" & vbCrLf & _
    "' ===================================================" & vbCrLf & vbCrLf & _
    "Public Sub CrearMenuPrincipal()" & vbCrLf & _
    "    On Error Resume Next" & vbCrLf & _
    "    Dim menuBar As CommandBar" & vbCrLf & _
    "    Dim menuItem As CommandBarControl" & vbCrLf & vbCrLf & _
    "    ' Eliminar menú anterior si existe" & vbCrLf & _
    "    Application.CommandBars(""Comparador IA"").Delete" & vbCrLf & vbCrLf & _
    "    ' Crear nueva barra de herramientas" & vbCrLf & _
    "    Set menuBar = Application.CommandBars.Add(Name:=""Comparador IA"", Position:=msoBarFloating, Temporary:=True)" & vbCrLf & vbCrLf & _
    "    ' Agregar botones" & vbCrLf & _
    "    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)" & vbCrLf & _
    "    With menuItem" & vbCrLf & _
    "        .Caption = ""&Comparar Producto""" & vbCrLf & _
    "        .TooltipText = ""Comparar precios de un producto""" & vbCrLf & _
    "        .OnAction = ""CompararProducto""" & vbCrLf & _
    "        .FaceId = 172" & vbCrLf & _
    "    End With" & vbCrLf & vbCrLf & _
    "    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)" & vbCrLf & _
    "    With menuItem" & vbCrLf & _
    "        .Caption = ""&Importar CSV""" & vbCrLf & _
    "        .TooltipText = ""Importar datos desde archivo CSV""" & vbCrLf & _
    "        .OnAction = ""ImportarDatosCSV""" & vbCrLf & _
    "        .FaceId = 23" & vbCrLf & _
    "    End With" & vbCrLf & vbCrLf & _
    "    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)" & vbCrLf & _
    "    With menuItem" & vbCrLf & _
    "        .Caption = ""&Exportar CSV""" & vbCrLf & _
    "        .TooltipText = ""Exportar datos a archivo CSV""" & vbCrLf & _
    "        .OnAction = ""ExportarDatosCSV""" & vbCrLf & _
    "        .FaceId = 308" & vbCrLf & _
    "    End With" & vbCrLf & vbCrLf & _
    "    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)" & vbCrLf & _
    "    With menuItem" & vbCrLf & _
    "        .Caption = ""&Backup""" & vbCrLf & _
    "        .TooltipText = ""Crear copia de seguridad del sistema""" & vbCrLf & _
    "        .OnAction = ""CrearBackupCompleto""" & vbCrLf & _
    "        .FaceId = 204" & vbCrLf & _
    "    End With" & vbCrLf & vbCrLf & _
    "    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)" & vbCrLf & _
    "    With menuItem" & vbCrLf & _
    "        .Caption = ""&Restaurar""" & vbCrLf & _
    "        .TooltipText = ""Restaurar datos desde backup""" & vbCrLf & _
    "        .OnAction = ""RestaurarDesdeBackup""" & vbCrLf & _
    "        .FaceId = 252" & vbCrLf & _
    "    End With" & vbCrLf & vbCrLf & _
    "    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)" & vbCrLf & _
    "    With menuItem" & vbCrLf & _
    "        .Caption = ""&Reporte""" & vbCrLf & _
    "        .TooltipText = ""Generar reporte del sistema""" & vbCrLf & _
    "        .OnAction = ""GenerarReporteSimple""" & vbCrLf & _
    "        .FaceId = 487" & vbCrLf & _
    "    End With" & vbCrLf & vbCrLf & _
    "    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)" & vbCrLf & _
    "    With menuItem" & vbCrLf & _
    "        .Caption = ""&Limpiar""" & vbCrLf & _
    "        .TooltipText = ""Limpiar todos los datos""" & vbCrLf & _
    "        .OnAction = ""LimpiarDatos""" & vbCrLf & _
    "        .FaceId = 196" & vbCrLf & _
    "    End With" & vbCrLf & vbCrLf & _
    "    ' Mostrar barra de herramientas" & vbCrLf & _
    "    menuBar.Visible = True" & vbCrLf & vbCrLf & _
    "    ' Actualizar estado" & vbCrLf & _
    "    ActualizarBarraEstado ""Menú creado - Sistema listo""" & vbCrLf & vbCrLf & _
    "    Set menuItem = Nothing" & vbCrLf & _
    "    Set menuBar = Nothing" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Public Sub MostrarMenuPrincipal()" & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & vbCrLf & _
    "    Dim respuesta As Variant" & vbCrLf & _
    "    Dim opcionNum As Integer" & vbCrLf & _
    "    Dim salir As Boolean" & vbCrLf & vbCrLf & _
    "    salir = False" & vbCrLf & vbCrLf & _
    "    Do While Not salir" & vbCrLf & _
    "        ' Crear el mensaje del menГғВә" & vbCrLf & _
    "        Dim menuTexto As String" & vbCrLf & _
    "        menuTexto = ""MENU PRINCIPAL - SISTEMA COMPARADOR DE COMPRAS IA"" & vbCrLf & vbCrLf & _" & vbCrLf & _
    "                   ""1. Comparar precios de producto"" & vbCrLf & _" & vbCrLf & _
    "                   ""2. Importar datos desde CSV"" & vbCrLf & _" & vbCrLf & _
    "                   ""3. Exportar datos a CSV"" & vbCrLf & _" & vbCrLf & _
    "                   ""4. Crear backup completo"" & vbCrLf & _" & vbCrLf & _
    "                   ""5. Restaurar desde backup"" & vbCrLf & _" & vbCrLf & _
    "                   ""6. Generar reporte simple"" & vbCrLf & _" & vbCrLf & _
    "                   ""7. Limpiar todos los datos"" & vbCrLf & _" & vbCrLf & _
    "                   ""8. Crear menu de herramientas"" & vbCrLf & _" & vbCrLf & _
    "                   ""9. Salir"" & vbCrLf & vbCrLf & _" & vbCrLf & _
    "                   ""Seleccione una opcion (1-9):""" & vbCrLf & vbCrLf & _
    "        ' Usar InputBox estГғВЎndar" & vbCrLf & _
    "        respuesta = InputBox(menuTexto, ""Menu del Sistema"", """")" & vbCrLf & vbCrLf & _
    "        ' Si el usuario cancela o presiona Cancel" & vbCrLf & _
    "        If respuesta = """" Then" & vbCrLf & _
    "            Exit Sub" & vbCrLf & _
    "        End If" & vbCrLf & vbCrLf & _
    "        ' Verificar que sea un nГғВәmero" & vbCrLf & _
    "        If Not IsNumeric(respuesta) Then" & vbCrLf & _
    "            MsgBox ""Por favor ingrese un nГғВәmero del 1 al 9."", vbExclamation" & vbCrLf & _
    "            GoTo ContinueLoop" & vbCrLf & _
    "        End If" & vbCrLf & vbCrLf & _
    "        ' Convertir a entero" & vbCrLf & _
    "        opcionNum = CInt(respuesta)" & vbCrLf & vbCrLf & _
    "        ' Validar rango" & vbCrLf & _
    "        If opcionNum < 1 Or opcionNum > 9 Then" & vbCrLf & _
    "            MsgBox ""Opcion no valida. Por favor seleccione 1-9."", vbExclamation" & vbCrLf & _
    "            GoTo ContinueLoop" & vbCrLf & _
    "        End If" & vbCrLf & vbCrLf & _
    "        ' Ejecutar opciГғВіn seleccionada" & vbCrLf & _
    "        Select Case opcionNum" & vbCrLf & _
    "            Case 1" & vbCrLf & _
    "                CompararProducto" & vbCrLf & _
    "            Case 2" & vbCrLf & _
    "                ImportarDatosCSV" & vbCrLf & _
    "            Case 3" & vbCrLf & _
    "                ExportarDatosCSV" & vbCrLf & _
    "            Case 4" & vbCrLf & _
    "                CrearBackupCompleto" & vbCrLf & _
    "            Case 5" & vbCrLf & _
    "                RestaurarDesdeBackup" & vbCrLf & _
    "            Case 6" & vbCrLf & _
    "                GenerarReporteSimple" & vbCrLf & _
    "            Case 7" & vbCrLf & _
    "                LimpiarDatos" & vbCrLf & _
    "            Case 8" & vbCrLf & _
    "                CrearMenuPrincipal" & vbCrLf & _
    "            Case 9" & vbCrLf & _
    "                salir = True" & vbCrLf & _
    "                MsgBox ""Saliendo del menГғВә principal."", vbInformation" & vbCrLf & _
    "        End Select" & vbCrLf & vbCrLf & _
    "ContinueLoop:" & vbCrLf & _
    "    Loop" & vbCrLf & vbCrLf & _
    "    Exit Sub" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    MsgBox ""Error en el menu principal: "" & Err.Description, vbCritical" & vbCrLf & _
    "End Sub"
    
    CreateInterfaceFunctions = code
End Function

Function CreateAutomationFunctions()
    Dim code
    code = "' ===================================================" & vbCrLf & _
    "' FUNCIONES DE AUTOMATIZACIÓN" & vbCrLf & _
    "' ===================================================" & vbCrLf & vbCrLf & _
    "Private Sub Workbook_Open()" & vbCrLf & _
    "    ' Ejecutar al abrir el libro" & vbCrLf & _
    "    Call InicializarSistema" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbCrLf & _
    "    ' Limpiar antes de cerrar" & vbCrLf & _
    "    Call LimpiarBarraEstado" & vbCrLf & _
    "End Sub"
    
    CreateAutomationFunctions = code
End Function

' ===================================================
' FUNCIÓN PRINCIPAL DEL SCRIPT
' ===================================================

Sub Main()
    Dim moduleCode, mathModuleCode
    Dim modulePath, mathModulePath
    Dim moduleStream, mathModuleStream
    
    ' Inicializar
    Call Initialize
    WriteLog "Iniciando agregado de macros al sistema", "INFO"
    WriteLog "Versión del script: " & scriptVersion, "INFO"
    WriteLog "Ruta del proyecto: " & projectPath, "INFO"
    
    ' Verificar que Excel esté instalado
    If Not ExcelInstalled() Then
        WriteLog "Excel no está disponible en el sistema", "ERROR"
        MsgBox "Excel no está disponible. No se pueden agregar macros.", vbCritical
        Exit Sub
    End If
    
    ' Verificar que el archivo Excel existe
    If Not FileExists(excelPath) Then
        WriteLog "Archivo Excel no encontrado en: " & excelPath, "ERROR"
        MsgBox "Archivo Excel no encontrado. Ejecute primero crear_sistema.bat", vbCritical
        Exit Sub
    End If
    
    ' Crear backup del archivo Excel
    WriteLog "Creando backup del archivo Excel actual", "INFO"
    Call CreateBackup
    
    ' Crear códigos de módulos
    WriteLog "Generando código VBA para módulos", "INFO"
    moduleCode = CreateMainModule()
    mathModuleCode = CreateMathModule()
    
    ' Guardar módulos en archivos .bas
    modulePath = projectPath & "\ModPrincipal.bas"
    mathModulePath = projectPath & "\ModMatematicas.bas"
    
    On Error Resume Next
    
    ' Guardar módulo principal
    Set moduleStream = fso.CreateTextFile(modulePath, True)
    moduleStream.Write moduleCode
    moduleStream.Close
    Set moduleStream = Nothing
    
    If Err.Number = 0 Then
        WriteLog "Módulo principal creado: " & modulePath, "SUCCESS"
    Else
        WriteLog "Error al crear módulo principal: " & Err.Description, "ERROR"
        Exit Sub
    End If
    
    ' Guardar módulo matemático
    Set mathModuleStream = fso.CreateTextFile(mathModulePath, True)
    mathModuleStream.Write mathModuleCode
    mathModuleStream.Close
    Set mathModuleStream = Nothing
    
    If Err.Number = 0 Then
        WriteLog "Módulo matemático creado: " & mathModulePath, "SUCCESS"
    Else
        WriteLog "Error al crear módulo matemático: " & Err.Description, "ERROR"
    End If
    
    On Error GoTo 0
    
    'Intentar importar módulos automáticamente
    If ImportarModulosExcel() Then
       WriteLog "Módulos importados exitosamente en Excel", "SUCCESS"
    Else
       WriteLog "No se pudieron importar los módulos automáticamente", "WARNING"
       WriteLog "Por favor, importe manualmente los archivos .bas", "INFO"
    End If
    
    ShowSummary
End Sub

Function CreateMathModule()
    Dim code
    code = "Option Explicit" & vbCrLf & vbCrLf & _
    "' ===================================================" & vbCrLf & _
    "' MÓDULO MATEMÁTICO - FUNCIONES DE CÁLCULO" & vbCrLf & _
    "' ===================================================" & vbCrLf & vbCrLf & _
    "Public Function CalcularDistanciaHaversine(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double" & vbCrLf & _
    "    Dim dLat As Double, dLon As Double" & vbCrLf & _
    "    Dim a As Double, c As Double" & vbCrLf & vbCrLf & _
    "    ' Convertir grados a radianes" & vbCrLf & _
    "    dLat = GradosARadianes(lat2 - lat1)" & vbCrLf & _
    "    dLon = GradosARadianes(lon2 - lon1)" & vbCrLf & vbCrLf & _
    "    ' Fórmula de Haversine" & vbCrLf & _
    "    a = Sin(dLat / 2) * Sin(dLat / 2) + _" & vbCrLf & _
    "        Cos(GradosARadianes(lat1)) * Cos(GradosARadianes(lat2)) * _" & vbCrLf & _
    "        Sin(dLon / 2) * Sin(dLon / 2)" & vbCrLf & vbCrLf & _
    "    c = 2 * WorksheetFunction.Atan2(Sqr(a), Sqr(1 - a))" & vbCrLf & vbCrLf & _
    "    ' Distancia en kilómetros" & vbCrLf & _
    "    CalcularDistanciaHaversine = EARTH_RADIUS_KM * c" & vbCrLf & _
    "End Function" & vbCrLf & vbCrLf & _
    "Public Function GradosARadianes(grados As Double) As Double" & vbCrLf & _
    "    GradosARadianes = grados * PI / 180" & vbCrLf & _
    "End Function" & vbCrLf & vbCrLf & _
    "Public Function CalcularPrecioPorUnidad(precioTotal As Double, cantidad As Double, unidadOrigen As String) As Double" & vbCrLf & _
    "    Dim factorConversion As Double" & vbCrLf & vbCrLf & _
    "    Select Case LCase(unidadOrigen)" & vbCrLf & _
    "        Case ""kg"", ""litro"", ""unidad""" & vbCrLf & _
    "            factorConversion = 1" & vbCrLf & _
    "        Case ""g""" & vbCrLf & _
    "            factorConversion = 1000" & vbCrLf & _
    "        Case ""ml""" & vbCrLf & _
    "            factorConversion = 1000" & vbCrLf & _
    "        Case ""mg""" & vbCrLf & _
    "            factorConversion = 1000000" & vbCrLf & _
    "        Case Else" & vbCrLf & _
    "            factorConversion = 1" & vbCrLf & _
    "    End Select" & vbCrLf & vbCrLf & _
    "    If cantidad > 0 Then" & vbCrLf & _
    "        CalcularPrecioPorUnidad = (precioTotal / cantidad) * factorConversion" & vbCrLf & _
    "    Else" & vbCrLf & _
    "        CalcularPrecioPorUnidad = 0" & vbCrLf & _
    "    End If" & vbCrLf & _
    "End Function" & vbCrLf & vbCrLf & _
    "Public Function CalcularAhorroPorcentual(precioOriginal As Double, precioOferta As Double) As Double" & vbCrLf & _
    "    If precioOriginal > 0 Then" & vbCrLf & _
    "        CalcularAhorroPorcentual = ((precioOriginal - precioOferta) / precioOriginal) * 100" & vbCrLf & _
    "    Else" & vbCrLf & _
    "        CalcularAhorroPorcentual = 0" & vbCrLf & _
    "    End If" & vbCrLf & _
    "End Function" & vbCrLf & vbCrLf & _
    "Public Function ValidarEmail(email As String) As Boolean" & vbCrLf & _
    "    Dim regex As Object" & vbCrLf & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & vbCrLf & _
    "    Set regex = CreateObject(""VBScript.RegExp"")" & vbCrLf & _
    "    With regex" & vbCrLf & _
    "        .Pattern = ""^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$""" & vbCrLf & _
    "        .IgnoreCase = True" & vbCrLf & _
    "        .Global = False" & vbCrLf & _
    "    End With" & vbCrLf & vbCrLf & _
    "    ValidarEmail = regex.Test(email)" & vbCrLf & vbCrLf & _
    "    Exit Function" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    ValidarEmail = False" & vbCrLf & _
    "End Function" & vbCrLf & vbCrLf & _
    "Public Function FormatearMoneda(valor As Double, Optional moneda As String = ""EUR"") As String" & vbCrLf & _
    "    Select Case UCase(moneda)" & vbCrLf & _
    "        Case ""EUR"", ""€""" & vbCrLf & _
    "            FormatearMoneda = Format(valor, ""0.00€"")" & vbCrLf & _
    "        Case ""USD"", ""$""" & vbCrLf & _
    "            FormatearMoneda = Format(valor, ""$0.00"")" & vbCrLf & _
    "        Case ""GBP"", ""£""" & vbCrLf & _
    "            FormatearMoneda = Format(valor, ""£0.00"")" & vbCrLf & _
    "        Case Else" & vbCrLf & _
    "            FormatearMoneda = Format(valor, ""0.00"")" & vbCrLf & _
    "    End Select" & vbCrLf & _
    "End Function"
    
    CreateMathModule = code
End Function

Function ImportarModulosExcel()
    
    Dim excel, workbook, vbComponent
    Dim modulePath, mathModulePath
    
    modulePath = projectPath & "\ModPrincipal.bas"
    mathModulePath = projectPath & "\ModMatematicas.bas"
    
    If Not (FileExists(modulePath) And FileExists(mathModulePath)) Then
        ImportarModulosExcel = False
        Exit Function
    End If
    
    ' Abrir Excel
    Set excel = CreateObject("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    ' Abrir workbook
    Set workbook = excel.Workbooks.Open(excelPath)
    
    ' Eliminar módulos existentes si los hay
    For Each vbComponent In workbook.VBProject.VBComponents
        If vbComponent.Type = 1 Then ' vbext_ct_StdModule
            If vbComponent.Name = "ModPrincipal" Or vbComponent.Name = "ModMatematicas" Then
                workbook.VBProject.VBComponents.Remove vbComponent
            End If
        End If
    Next
    
    ' Importar módulo principal
    workbook.VBProject.VBComponents.Import modulePath
    
    ' Importar módulo matemático
    workbook.VBProject.VBComponents.Import mathModulePath
    
    ' Guardar y cerrar
    workbook.Save
    workbook.Close True
    excel.Quit
    
    ' Limpiar objetos
    Set vbComponent = Nothing
    Set workbook = Nothing
    Set excel = Nothing
    
    ImportarModulosExcel = True
    Exit Function
    
ErrorHandler:
    ImportarModulosExcel = False
    On Error Resume Next
    If Not excel Is Nothing Then
        workbook.Close False
        excel.Quit
    End If
    Set vbComponent = Nothing
    Set workbook = Nothing
    Set excel = Nothing
End Function

Sub ShowSummary()
    Dim summary, duration
    duration = DateDiff("s", startTime, Now())
    
    summary = "===================================================" & vbCrLf & _
              " RESUMEN DE AGREGADO DE MACROS" & vbCrLf & _
              "===================================================" & vbCrLf & _
              "Versión: " & scriptVersion & vbCrLf & _
              "Duración: " & duration & " segundos" & vbCrLf & _
              "Éxitos: " & successCount & vbCrLf & _
              "Advertencias: " & warningCount & vbCrLf & _
              "Errores: " & errorCount & vbCrLf & _
              "===================================================" & vbCrLf & _
              "ARCHIVOS CREADOS:" & vbCrLf & _
              "• " & projectPath & "\ModPrincipal.bas" & vbCrLf & _
              "• " & projectPath & "\ModMatematicas.bas" & vbCrLf & _
              "===================================================" & vbCrLf & _
              "INSTRUCCIONES:" & vbCrLf & _
              "1. Los módulos VBA han sido creados en la carpeta del proyecto" & vbCrLf & _
              "2. Si no se importaron automáticamente, importe manualmente:" & vbCrLf & _
              "   a) Abra " & excelPath & vbCrLf & _
              "   b) Presione ALT+F11 (Editor VBA)" & vbCrLf & _
              "   c) Archivo → Importar archivo" & vbCrLf & _
              "   d) Seleccione los archivos .bas" & vbCrLf & _
              "3. Guarde el archivo Excel" & vbCrLf & _
              "==================================================="
    
    WriteLog summary, "INFO"
    
    If InStr(1, WScript.FullName, "cscript.exe", vbTextCompare) > 0 Then
        WScript.Echo summary
    Else
        MsgBox "Proceso completado." & vbCrLf & vbCrLf & _
               "Éxitos: " & successCount & vbCrLf & _
               "Advertencias: " & warningCount & vbCrLf & _
               "Errores: " & errorCount & vbCrLf & vbCrLf & _
               "Revise el archivo de log para más detalles:" & vbCrLf & _
               logPath, vbInformation, "Agregado de Macros Completado"
    End If
End Sub

' ===================================================
' EJECUCIÓN PRINCIPAL
' ===================================================
Call Main

' Limpiar
Set fso = Nothing
Set shell = Nothing

WScript.Quit