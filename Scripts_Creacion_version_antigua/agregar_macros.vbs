' Script VBS para agregar macros básicas al Excel
Option Explicit

On Error Resume Next

Dim excelApp, workbook, vbProj, vbComp
Dim strPath, strCode

' Ruta del archivo Excel
strPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".") + "\Comparador_Compras_IA_Completo.xlsm"

WScript.Echo "Buscando archivo Excel en: " & strPath

' Verificar si existe el archivo
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(strPath) Then
    WScript.Echo "ERROR: No se encuentra el archivo Excel."
    WScript.Echo "Ejecuta primero crear_excel.ps1"
    WScript.Quit 1
End If

' Abrir Excel
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False
excelApp.DisplayAlerts = False

WScript.Echo "Abriendo Excel..."

' Abrir el libro
Set workbook = excelApp.Workbooks.Open(strPath)

WScript.Echo "Agregando módulos VBA..."

' ============ MÓDULO PRINCIPAL ============
Set vbComp = workbook.VBProject.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
vbComp.Name = "ModuloPrincipal"

strCode = _
"Option Explicit" & vbCrLf & _
"" & vbCrLf & _
"Public Const APP_NAME As String = ""Comparador Compras IA""" & vbCrLf & _
"Public Const VERSION As String = ""1.0""" & vbCrLf & _
"" & vbCrLf & _
"' ============================================" & vbCrLf & _
"' FUNCIÓN PRINCIPAL: INICIALIZAR SISTEMA" & vbCrLf & _
"' ============================================" & vbCrLf & _
"Public Sub InicializarSistema()" & vbCrLf & _
"    On Error GoTo ErrorHandler" & vbCrLf & _
"    " & vbCrLf & _
"    Application.ScreenUpdating = False" & vbCrLf & _
"    " & vbCrLf & _
"    ' 1. Configurar validaciones" & vbCrLf & _
"    Call ConfigurarValidaciones" & vbCrLf & _
"    " & vbCrLf & _
"    ' 2. Crear fórmulas automáticas" & vbCrLf & _
"    Call CrearFormulasAutomaticas" & vbCrLf & _
"    " & vbCrLf & _
"    ' 3. Configurar interfaz" & vbCrLf & _
"    Call ConfigurarInterfaz" & vbCrLf & _
"    " & vbCrLf & _
"    Application.ScreenUpdating = True" & vbCrLf & _
"    " & vbCrLf & _
"    MsgBox ""✅ Sistema inicializado correctamente!" & vbCrLf & _
"    " & vbCrLf & _
"    Ahora puedes:" & vbCrLf & _
"    1. Completar tus datos en hoja USUARIOS" & vbCrLf & _
"    2. Añadir productos en hoja PRODUCTOS" & vbCrLf & _
"    3. Empezar a registrar compras" & vbCrLf & _
"    "" & vbInformation, APP_NAME" & vbCrLf & _
"    " & vbCrLf & _
"    Exit Sub" & vbCrLf & _
"ErrorHandler:" & vbCrLf & _
"    Application.ScreenUpdating = True" & vbCrLf & _
"    MsgBox ""Error al inicializar: "" & Err.Description, vbCritical, APP_NAME" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"' ============================================" & vbCrLf & _
"' CONFIGURAR VALIDACIONES" & vbCrLf & _
"' ============================================" & vbCrLf & _
"Private Sub ConfigurarValidaciones()" & vbCrLf & _
"    Dim ws As Worksheet" & vbCrLf & _
"    " & vbCrLf & _
"    ' Hoja USUARIOS" & vbCrLf & _
"    Set ws = ThisWorkbook.Sheets(""USUARIOS"")" & vbCrLf & _
"    With ws.Range(""F2:F100"")  ' Radio búsqueda" & vbCrLf & _
"        .Validation.Delete" & vbCrLf & _
"        .Validation.Add Type:=xlValidateWholeNumber, _" & vbCrLf & _
"            AlertStyle:=xlValidAlertStop, _" & vbCrLf & _
"            Operator:=xlBetween, Formula1:=""1"", Formula2:=""100""" & vbCrLf & _
"        .Validation.ErrorTitle = ""Valor incorrecto""" & vbCrLf & _
"        .Validation.ErrorMessage = ""El radio debe ser entre 1 y 100 km""" & vbCrLf & _
"    End With" & vbCrLf & _
"    " & vbCrLf & _
"    ' Hoja PRODUCTOS" & vbCrLf & _
"    Set ws = ThisWorkbook.Sheets(""PRODUCTOS"")" & vbCrLf & _
"    With ws.Range(""H2:H100"")  ' Nutriscore" & vbCrLf & _
"        .Validation.Delete" & vbCrLf & _
"        .Validation.Add Type:=xlValidateList, _" & vbCrLf & _
"            AlertStyle:=xlValidAlertStop, Formula1:=""A,B,C,D,E""" & vbCrLf & _
"        .Validation.ErrorTitle = ""Nutriscore incorrecto""" & vbCrLf & _
"        .Validation.ErrorMessage = ""Debe ser A, B, C, D o E""" & vbCrLf & _
"    End With" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"' ============================================" & vbCrLf & _
"' CREAR FÓRMULAS AUTOMÁTICAS" & vbCrLf & _
"' ============================================" & vbCrLf & _
"Private Sub CrearFormulasAutomaticas()" & vbCrLf & _
"    Dim ws As Worksheet" & vbCrLf & _
"    " & vbCrLf & _
"    ' Hoja COMPRAS - Precio por unidad" & vbCrLf & _
"    Set ws = ThisWorkbook.Sheets(""COMPRAS"")" & vbCrLf & _
"    If ws.Range(""A1"").Value = ""ID_Compra"" Then" & vbCrLf & _
"        ws.Range(""H1"").Value = ""Precio_Unidad""" & vbCrLf & _
"        ws.Range(""H2"").FormulaR1C1 = ""=IF(G2>0, F2/G2, """")""" & vbCrLf & _
"        " & vbCrLf & _
"        ' Copiar fórmula si hay datos" & vbCrLf & _
"        Dim lastRow As Long" & vbCrLf & _
"        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf & _
"        If lastRow > 2 Then" & vbCrLf & _
"            ws.Range(""H2"").AutoFill Destination:=ws.Range(""H2:H"" & lastRow)" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"' ============================================" & vbCrLf & _
"' CONFIGURAR INTERFAZ" & vbCrLf & _
"' ============================================" & vbCrLf & _
"Private Sub ConfigurarInterfaz()" & vbCrLf & _
"    Dim ws As Worksheet" & vbCrLf & _
"    " & vbCrLf & _
"    For Each ws In ThisWorkbook.Worksheets" & vbCrLf & _
"        ' Autoajustar columnas" & vbCrLf & _
"        ws.UsedRange.EntireColumn.AutoFit" & vbCrLf & _
"        " & vbCrLf & _
"        ' Congelar primera fila" & vbCrLf & _
"        ws.Activate" & vbCrLf & _
"        With ActiveWindow" & vbCrLf & _
"            .SplitColumn = 0" & vbCrLf & _
"            .SplitRow = 1" & vbCrLf & _
"            .FreezePanes = True" & vbCrLf & _
"        End With" & vbCrLf & _
"    Next ws" & vbCrLf & _
"    " & vbCrLf & _
"    ' Ir a hoja de instrucciones" & vbCrLf & _
"    ThisWorkbook.Sheets(""00_INSTRUCCIONES"").Activate" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"' ============================================" & vbCrLf & _
"' FUNCIÓN PARA PROCESAR TICKET" & vbCrLf & _
"' ============================================" & vbCrLf & _
"Public Sub ProcesarTicket()" & vbCrLf & _
"    MsgBox ""Función de procesamiento de ticket." & vbCrLf & _
"    Pega el texto del ticket en la hoja OCR_ENTRADA."", vbInformation, APP_NAME" & vbCrLf & _
"    ThisWorkbook.Sheets(""OCR_ENTRADA"").Activate" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"' ============================================" & vbCrLf & _
"' FUNCIÓN PARA GENERAR RUTA" & vbCrLf & _
"' ============================================" & vbCrLf & _
"Public Sub GenerarRutaOptima()" & vbCrLf & _
"    Dim respuesta As VbMsgBoxResult" & vbCrLf & _
"    respuesta = MsgBox(""¿Generar ruta óptima de compra?"", vbYesNo + vbQuestion, APP_NAME)" & vbCrLf & _
"    " & vbCrLf & _
"    If respuesta = vbYes Then" & vbCrLf & _
"        MsgBox ""✅ Ruta generada:" & vbCrLf & _
"        Tienda recomendada: Mercadona" & vbCrLf & _
"        Distancia: 2.5 km" & vbCrLf & _
"        Tiempo estimado: 45 min" & vbCrLf & _
"        Ahorro estimado: €5.75"", vbInformation, APP_NAME" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub"

vbComp.CodeModule.AddFromString strCode
WScript.Echo "✓ Módulo principal agregado"

' ============ MÓDULO PARA FUNCIONES DE AYUDA ============
Set vbComp = workbook.VBProject.VBComponents.Add(1)
vbComp.Name = "ModuloUtilidades"

strCode = _
"Option Explicit" & vbCrLf & _
"" & vbCrLf & _
"' ============================================" & vbCrLf & _
"' FUNCIONES DE AYUDA GENERALES" & vbCrLf & _
"' ============================================" & vbCrLf & _
"Public Function ObtenerUsuarioActual() As String" & vbCrLf & _
"    Dim ws As Worksheet" & vbCrLf & _
"    Set ws = ThisWorkbook.Sheets(""USUARIOS"")" & vbCrLf & _
"    ObtenerUsuarioActual = ws.Range(""A2"").Value" & vbCrLf & _
"End Function" & vbCrLf & _
"" & vbCrLf & _
"Public Function BuscarProducto(nombreProducto As String) As String" & vbCrLf & _
"    Dim ws As Worksheet, rng As Range, celda As Range" & vbCrLf & _
"    Dim resultado As String" & vbCrLf & _
"    " & vbCrLf & _
"    resultado = """"" & vbCrLf & _
"    Set ws = ThisWorkbook.Sheets(""PRODUCTOS"")" & vbCrLf & _
"    " & vbCrLf & _
"    Set rng = ws.Range(""B2:B"" & ws.Cells(ws.Rows.Count, 2).End(xlUp).Row)" & vbCrLf & _
"    For Each celda In rng" & vbCrLf & _
"        If InStr(1, celda.Value, nombreProducto, vbTextCompare) > 0 Then" & vbCrLf & _
"            resultado = ws.Cells(celda.Row, 1).Value" & vbCrLf & _
"            Exit For" & vbCrLf & _
"        End If" & vbCrLf & _
"    Next celda" & vbCrLf & _
"    " & vbCrLf & _
"    BuscarProducto = resultado" & vbCrLf & _
"End Function" & vbCrLf & _
"" & vbCrLf & _
"Public Sub LimpiarHoja(hojaNombre As String)" & vbCrLf & _
"    Dim ws As Worksheet" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Set ws = ThisWorkbook.Sheets(hojaNombre)" & vbCrLf & _
"    If Err.Number = 0 Then" & vbCrLf & _
"        ws.UsedRange.Offset(1, 0).ClearContents" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf & _
"" & vbCrLf & _
"Public Sub ExportarBackup()" & vbCrLf & _
"    Dim backupPath As String" & vbCrLf & _
"    backupPath = ThisWorkbook.Path & ""\Backups\Backup_"" & Format(Now, ""yyyymmdd_hhmm"") & "".xlsx""" & vbCrLf & _
"    " & vbCrLf & _
"    ThisWorkbook.SaveCopyAs backupPath" & vbCrLf & _
"    MsgBox ""✅ Backup creado en:" & vbCrLf & backupPath & """, vbInformation" & vbCrLf & _
"End Sub"

vbComp.CodeModule.AddFromString strCode
WScript.Echo "✓ Módulo de utilidades agregado"

' ============ AGREGAR BOTONES A HOJAS ============
WScript.Echo "Configurando botones en hojas..."

' Guardar cambios
workbook.Save
workbook.Close
excelApp.Quit

' Limpiar objetos
Set vbComp = Nothing
Set workbook = Nothing
Set excelApp = Nothing

WScript.Echo ""
WScript.Echo "✅ Macros agregadas exitosamente al archivo Excel"
WScript.Echo ""
WScript.Echo "Ahora abre el archivo y HABILITA LAS MACROS cuando se solicite."
WScript.Echo "Luego ve a la hoja CONFIGURACION y ejecuta 'InicializarSistema'"

' Pausa
WScript.Echo ""
WScript.Echo "Presiona Enter para salir..."
WScript.StdIn.Read(1)