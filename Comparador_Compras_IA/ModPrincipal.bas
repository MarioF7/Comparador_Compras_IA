Option Explicit

' ===================================================
' MÓDULO PRINCIPAL - SISTEMA COMPARADOR DE COMPRAS IA
' Versión: 4.0.0
' ===================================================

' CONSTANTES DEL SISTEMA
Public Const SISTEMA_VERSION As String = "4.0.0"
Public Const SISTEMA_NOMBRE As String = "Sistema Comparador de Compras Inteligente IA"
Public Const SISTEMA_AUTOR As String = "Equipo de Desarrollo IA"

' Códigos de error
Public Const ERR_SUCCESS As Long = 0
Public Const ERR_FILE_NOT_FOUND As Long = 1
Public Const ERR_INVALID_DATA As Long = 2
Public Const ERR_DATABASE As Long = 3
Public Const ERR_CALCULATION As Long = 4

' Configuración
Public Const MAX_RECORDS As Long = 100000
Public Const MAX_USERS As Long = 1000
Public Const MAX_PRODUCTS As Long = 50000
Public Const MAX_STORES As Long = 10000
Public Const EARTH_RADIUS_KM As Double = 6371.0
Public Const PI As Double = 3.14159265358979
' VARIABLES GLOBALES
Public gUsuarioActual As String
Public gRutaProyecto As String
Public gConfigLoaded As Boolean
Public gLastError As String
Public gLastErrorCode As Long
Public gSystemInitialized As Boolean
Public gLogEnabled As Boolean
Public gBackupEnabled As Boolean
' ===================================================
' FUNCIONES DE INICIALIZACIÓN
' ===================================================

Public Sub InicializarSistema()
    On Error GoTo ErrorHandler

    ' Configurar variables globales
    gUsuarioActual = ""
    gRutaProyecto = ThisWorkbook.Path
    gConfigLoaded = False
    gLastError = ""
    gLastErrorCode = ERR_SUCCESS
    gSystemInitialized = False
    gLogEnabled = True
    gBackupEnabled = True

    ' Verificar estructura
    If Not VerificarEstructura() Then
        MsgBox "Error en la estructura del sistema. Verifique las hojas.", vbCritical
        Exit Sub
    End If

    ' Cargar configuración
    If Not CargarConfiguracion() Then
        MsgBox "No se pudo cargar la configuración. Se usarán valores por defecto.", vbExclamation
    End If

    ' Crear menú
    CrearMenuPrincipal

    ' Actualizar estado
    gSystemInitialized = True
    ActualizarBarraEstado "Sistema inicializado correctamente"

    ' Mostrar mensaje de bienvenida
    If Application.Version >= 12.0 Then
        MsgBox "Bienvenido al " & SISTEMA_NOMBRE & " v" & SISTEMA_VERSION & vbCrLf & _
               "Sistema listo para comparar precios y optimizar compras.", _
               vbInformation, "Inicialización del Sistema"
    End If

    Exit Sub

ErrorHandler:
    gLastError = Err.Description
    gLastErrorCode = Err.Number
    MsgBox "Error durante la inicialización: " & Err.Description, vbCritical
End Sub

Private Function VerificarEstructura() As Boolean
    On Error GoTo ErrorHandler
    Dim hojasRequeridas As Variant
    Dim hoja As Worksheet
    Dim i As Integer
    Dim hojaEncontrada As Boolean

    hojasRequeridas = Array("USUARIOS", "PRODUCTOS", "TIENDAS", "PRECIOS", _
                           "COMPARATIVA", "HISTORIAL_COMPRAS", "PREFERENCIAS_IA")

    For i = LBound(hojasRequeridas) To UBound(hojasRequeridas)
        hojaEncontrada = False
        For Each hoja In ThisWorkbook.Worksheets
            If hoja.Name = hojasRequeridas(i) Then
                hojaEncontrada = True
                Exit For
            End If
        Next hoja

        If Not hojaEncontrada Then
            VerificarEstructura = False
            Exit Function
        End If
    Next i

    VerificarEstructura = True
    Exit Function

ErrorHandler:
    VerificarEstructura = False
End Function

Private Function CargarConfiguracion() As Boolean
    On Error GoTo ErrorHandler
    Dim configPath As String
    Dim fso As Object
    Dim ts As Object
    Dim configText As String

    configPath = gRutaProyecto & "\Configuraciones\config_sistema.json"
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(configPath) Then
        CargarConfiguracion = False
        Exit Function
    End If

    Set ts = fso.OpenTextFile(configPath, ForReading)
    configText = ts.ReadAll
    ts.Close

    ' Aquí se procesaría el JSON (simplificado)
    gConfigLoaded = True
    CargarConfiguracion = True

    Exit Function

ErrorHandler:
    CargarConfiguracion = False
End Function
' ===================================================
' FUNCIONES DE BACKUP Y SEGURIDAD
' ===================================================

Public Sub CrearBackupCompleto()
    On Error GoTo ErrorHandler

    Dim backupPath As String
    Dim backupName As String

    If Not gBackupEnabled Then
        MsgBox "El sistema de backup está deshabilitado.", vbExclamation
        Exit Sub
    End If

    ' Generar nombre único
    backupName = "backup_completo_" & Format(Now(), "yyyymmdd_hhnnss") & ".xlsm"
    backupPath = gRutaProyecto & "\Data_Backup\" & backupName

    ' Crear copia
    ThisWorkbook.SaveCopyAs backupPath

    MsgBox "Backup creado exitosamente en:" & vbCrLf & backupPath, vbInformation
    ActualizarBarraEstado "Backup completado"

    Exit Sub

ErrorHandler:
    MsgBox "Error al crear backup: " & Err.Description, vbCritical
End Sub

Public Sub RestaurarDesdeBackup()
    On Error GoTo ErrorHandler

    Dim rutaArchivo As String
    Dim respuesta As Integer

    ' Advertencia
    respuesta = MsgBox("ADVERTENCIA: Restaurar desde backup sobrescribirá todos los datos actuales." & vbCrLf & _
               "¿Está seguro de continuar?", vbCritical + vbYesNo, "Confirmar Restauración")

    If respuesta <> vbYes Then Exit Sub

    ' Seleccionar archivo
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Seleccionar archivo de backup para restaurar"
        .InitialFileName = gRutaProyecto & "\Data_Backup\"
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xls;*.xlsx;*.xlsm"
        .AllowMultiSelect = False

        If .Show = -1 Then
            rutaArchivo = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With

    ' Restaurar datos
    If RestaurarBackup(rutaArchivo) Then
        MsgBox "Datos restaurados exitosamente desde backup.", vbInformation
        ActualizarBarraEstado "Sistema restaurado desde backup"
    Else
        MsgBox "Error al restaurar datos desde backup.", vbCritical
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error durante la restauración: " & Err.Description, vbCritical
End Sub

Private Function RestaurarBackup(rutaArchivo As String) As Boolean
    On Error GoTo ErrorHandler
    Dim backupWorkbook As Workbook
    Dim wsBackup As Worksheet, wsActual As Worksheet
    Dim nombreHoja As String

    ' Abrir backup
    Set backupWorkbook = Workbooks.Open(rutaArchivo, ReadOnly:=True)

    ' Copiar datos de cada hoja
    For Each wsBackup In backupWorkbook.Worksheets
        nombreHoja = wsBackup.Name

        On Error Resume Next
        Set wsActual = ThisWorkbook.Worksheets(nombreHoja)
        On Error GoTo ErrorHandler

        If Not wsActual Is Nothing Then
            ' Limpiar hoja actual
            wsActual.Cells.ClearContents
            ' Copiar datos
            wsBackup.UsedRange.Copy wsActual.Range("A1")
        End If
    Next wsBackup

    ' Cerrar backup
    backupWorkbook.Close SaveChanges:=False

    RestaurarBackup = True
    Exit Function

ErrorHandler:
    RestaurarBackup = False
End Function
' ===================================================
' FUNCIONES DE COMPARACIГғвҖңN
' ===================================================

Public Sub CompararProducto()
    On Error GoTo ErrorHandler

    Dim nombreProducto As String, categoria As String
    Dim wsProductos As Worksheet, wsPrecios As Worksheet, wsTiendas As Worksheet
    Dim rngProductos As Range, celda As Range
    Dim idProducto As String, precioMasBarato As Double, tiendaMasBarata As String
    Dim resultado As String, contador As Integer
    Dim arrResultados() As Variant, i As Integer
    Dim distancia As Double, latUsuario As Double, lonUsuario As Double

    ' Solicitar datos del producto
    nombreProducto = InputBox("Ingrese el nombre del producto a comparar:", "Comparar Producto")
    If nombreProducto = "" Then Exit Sub
    categoria = InputBox("Ingrese la categorГғВӯa del producto (opcional):", "Comparar Producto")

    ' Obtener coordenadas del usuario (opcional)
    On Error Resume Next
    latUsuario = CDbl(InputBox("Ingrese su latitud (opcional, para calcular distancias):", "UbicaciГғВіn", "0"))
    lonUsuario = CDbl(InputBox("Ingrese su longitud (opcional):", "UbicaciГғВіn", "0"))
    On Error GoTo ErrorHandler

    ' Configurar hojas
    Set wsProductos = ThisWorkbook.Sheets("PRODUCTOS")
    Set wsPrecios = ThisWorkbook.Sheets("PRECIOS")
    Set wsTiendas = ThisWorkbook.Sheets("TIENDAS")

    ' Buscar producto en la base de datos
    Set rngProductos = wsProductos.Range("A2:A" & wsProductos.Cells(wsProductos.Rows.Count, 1).End(xlUp).Row)
    idProducto = ""
    For Each celda In rngProductos
        If InStr(1, wsProductos.Cells(celda.Row, 2).Value, nombreProducto, vbTextCompare) > 0 Then
            If categoria = "" Or LCase(wsProductos.Cells(celda.Row, 3).Value) = LCase(categoria) Then
                idProducto = celda.Value
                Exit For
            End If
        End If
    Next celda

    If idProducto = "" Then
        MsgBox "Producto no encontrado en la base de datos.", vbExclamation
        Exit Sub
    End If

    ' Buscar precios para el producto
    ReDim arrResultados(1 To 100, 1 To 6) ' ID, Tienda, Precio, Descuento, Distancia, Precio/Unidad
    contador = 0
    precioMasBarato = 9999999
    tiendaMasBarata = ""

    Dim rngPrecios As Range, celdaPrecio As Range
    Set rngPrecios = wsPrecios.Range("A2:A" & wsPrecios.Cells(wsPrecios.Rows.Count, 1).End(xlUp).Row)

    For Each celdaPrecio In rngPrecios
        If celdaPrecio.Value = idProducto Then
            Dim idTienda As String, precio As Double, descuento As Double
            Dim nombreTienda As String, latTienda As Double, lonTienda As Double
            Dim precioPorUnidad As Double, unidad As String

            idTienda = wsPrecios.Cells(celdaPrecio.Row, 2).Value
            precio = wsPrecios.Cells(celdaPrecio.Row, 3).Value
            descuento = wsPrecios.Cells(celdaPrecio.Row, 4).Value
            unidad = wsPrecios.Cells(celdaPrecio.Row, 5).Value

            ' Buscar informaciГғВіn de la tienda
            Dim rngTiendas As Range, celdaTienda As Range
            Set rngTiendas = wsTiendas.Range("A2:A" & wsTiendas.Cells(wsTiendas.Rows.Count, 1).End(xlUp).Row)

            For Each celdaTienda In rngTiendas
                If celdaTienda.Value = idTienda Then
                    nombreTienda = wsTiendas.Cells(celdaTienda.Row, 2).Value
                    latTienda = wsTiendas.Cells(celdaTienda.Row, 6).Value
                    lonTienda = wsTiendas.Cells(celdaTienda.Row, 7).Value
                    Exit For
                End If
            Next celdaTienda

            ' Calcular distancia si se proporcionaron coordenadas
            If latUsuario <> 0 And lonUsuario <> 0 Then
                distancia = CalcularDistanciaHaversine(latUsuario, lonUsuario, latTienda, lonTienda)
            Else
                distancia = 0
            End If

            ' Calcular precio por unidad
            If unidad <> "" Then
                precioPorUnidad = CalcularPrecioPorUnidad(precio, 1, unidad)
            Else
                precioPorUnidad = precio
            End If

            ' Almacenar resultado
            contador = contador + 1
            arrResultados(contador, 1) = idTienda
            arrResultados(contador, 2) = nombreTienda
            arrResultados(contador, 3) = precio
            arrResultados(contador, 4) = descuento
            arrResultados(contador, 5) = distancia
            arrResultados(contador, 6) = precioPorUnidad

            ' Verificar si es el precio mГғВЎs barato
            If precio < precioMasBarato Then
                precioMasBarato = precio
                tiendaMasBarata = nombreTienda
            End If
        End If
    Next celdaPrecio

    ' Mostrar resultados
    If contador > 0 Then
        resultado = "COMPARATIVA DE PRECIOS" & vbCrLf & vbCrLf
        resultado = resultado & "Producto: " & nombreProducto & vbCrLf
        resultado = resultado & "ID Producto: " & idProducto & vbCrLf & vbCrLf
        resultado = resultado & "Se encontraron " & contador & " precios diferentes:" & vbCrLf & vbCrLf
        resultado = resultado & String(50, "-") & vbCrLf
        For i = 1 To contador
            resultado = resultado & i & ". " & arrResultados(i, 2) & vbCrLf
            resultado = resultado & "   Precio: " & FormatearMoneda(arrResultados(i, 3)) & vbCrLf
            If arrResultados(i, 4) > 0 Then
                resultado = resultado & "   Descuento: " & arrResultados(i, 4) & "%" & vbCrLf
            End If
            If arrResultados(i, 5) > 0 Then
                resultado = resultado & "   Distancia: " & Format(arrResultados(i, 5), "0.0") & " km" & vbCrLf
            End If
            resultado = resultado & "   Precio por unidad: " & FormatearMoneda(arrResultados(i, 6)) & vbCrLf & vbCrLf
            resultado = resultado & String(50, "-") & vbCrLf
        Next i

        resultado = resultado & vbCrLf & "MEJOR PRECIO: " & FormatearMoneda(precioMasBarato) & " en " & tiendaMasBarata & vbCrLf
        If contador > 1 Then
            Dim ahorroPromedio As Double
            Dim precioMasCaro As Double
            precioMasCaro = 0
            For i = 1 To contador
                If arrResultados(i, 3) > precioMasCaro Then precioMasCaro = arrResultados(i, 3)
            Next i
            ahorroPromedio = CalcularAhorroPorcentual(precioMasCaro, precioMasBarato)
            resultado = resultado & "Ahorro mГғВЎximo posible: " & Format(ahorroPromedio, "0.0") & "%" & vbCrLf
        End If

        ' Mostrar en cuadro de mensaje
        MsgBox resultado, vbInformation, "Comparativa de Precios"

        ' Registrar en hoja COMPARATIVA
        Dim wsComparativa As Worksheet
        Set wsComparativa = ThisWorkbook.Sheets("COMPARATIVA")
        Dim ultimaFila As Long
        ultimaFila = wsComparativa.Cells(wsComparativa.Rows.Count, 1).End(xlUp).Row + 1
        wsComparativa.Cells(ultimaFila, 1).Value = Now()
        wsComparativa.Cells(ultimaFila, 2).Value = idProducto
        wsComparativa.Cells(ultimaFila, 3).Value = nombreProducto
        wsComparativa.Cells(ultimaFila, 4).Value = contador
        wsComparativa.Cells(ultimaFila, 5).Value = precioMasBarato
        wsComparativa.Cells(ultimaFila, 6).Value = tiendaMasBarata
        wsComparativa.Cells(ultimaFila, 7).Value = gUsuarioActual

        ActualizarBarraEstado "Comparativa completada para " & nombreProducto
    Else
        MsgBox "No se encontraron precios para el producto especificado.", vbExclamation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error en la comparaciГғВіn: " & Err.Description, vbCritical
End Sub
' ===================================================
' FUNCIONES DE IMPORTACIÓN/EXPORTACIÓN
' ===================================================

Public Sub ImportarDatosCSV()
    On Error GoTo ErrorHandler
    Dim rutaArchivo As String
    Dim hojaDestino As String

    ' Seleccionar archivo
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Seleccionar archivo CSV para importar"
        .Filters.Clear
        .Filters.Add "Archivos CSV", "*.csv"
        .AllowMultiSelect = False

        If .Show = -1 Then
            rutaArchivo = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With

    ' Seleccionar hoja destino
    hojaDestino = InputBox("Ingrese el nombre de la hoja destino:", "Importar Datos")
    If hojaDestino = "" Then Exit Sub

    ' Importar datos
    If ImportarCSV(rutaArchivo, hojaDestino) Then
        MsgBox "Datos importados exitosamente.", vbInformation
        ActualizarBarraEstado "Datos importados desde CSV"
    Else
        MsgBox "Error al importar datos.", vbCritical
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error durante la importación: " & Err.Description, vbCritical
End Sub

Private Function ImportarCSV(rutaArchivo As String, hojaDestino As String) As Boolean
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim fso As Object, ts As Object
    Dim lineas() As String, campos() As String
    Dim i As Long, j As Long
    Dim textoArchivo As String

    ' Verificar que la hoja existe
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(hojaDestino)
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        MsgBox "La hoja especificada no existe.", vbExclamation
        ImportarCSV = False
        Exit Function
    End If

    ' Limpiar datos existentes (excepto cabeceras)
    If ws.UsedRange.Rows.Count > 1 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count)).ClearContents
    End If

    ' Leer archivo CSV
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(rutaArchivo, ForReading, TristateTrue)
    textoArchivo = ts.ReadAll
    ts.Close

    ' Procesar líneas
    lineas = Split(textoArchivo, vbCrLf)
    For i = 1 To UBound(lineas)
        If Trim(lineas(i)) <> "" Then
            campos = Split(lineas(i), ",")
            For j = 0 To UBound(campos)
                ws.Cells(i + 1, j + 1).Value = campos(j)
            Next j
        End If
    Next i

    ImportarCSV = True
    Exit Function

ErrorHandler:
    ImportarCSV = False
End Function

Public Sub ExportarDatosCSV()
    On Error GoTo ErrorHandler
    Dim hojaOrigen As String
    Dim rutaArchivo As String

    ' Seleccionar hoja
    hojaOrigen = InputBox("Ingrese el nombre de la hoja a exportar:", "Exportar Datos")
    If hojaOrigen = "" Then Exit Sub

    ' Seleccionar destino
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "Guardar archivo CSV"
        .InitialFileName = hojaOrigen & ".csv"
        .Filters.Clear
        .Filters.Add "Archivos CSV", "*.csv"
        .AllowMultiSelect = False

        If .Show = -1 Then
            rutaArchivo = .SelectedItems(1)
            ' Asegurar extensión .csv
            If LCase(Right(rutaArchivo, 4)) <> ".csv" Then
                rutaArchivo = rutaArchivo & ".csv"
            End If
        Else
            Exit Sub
        End If
    End With

    ' Exportar datos
    If ExportarCSV(hojaOrigen, rutaArchivo) Then
        MsgBox "Datos exportados exitosamente a " & rutaArchivo, vbInformation
        ActualizarBarraEstado "Datos exportados a CSV"
    Else
        MsgBox "Error al exportar datos.", vbCritical
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error durante la exportación: " & Err.Description, vbCritical
End Sub

Private Function ExportarCSV(hojaOrigen As String, rutaArchivo As String) As Boolean
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim fso As Object, ts As Object
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim linea As String
    Dim valorCelda As String

    ' Verificar que la hoja existe
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(hojaOrigen)
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        ExportarCSV = False
        Exit Function
    End If

    ' Determinar rango de datos
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Crear archivo CSV
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(rutaArchivo, True)

    For i = 1 To lastRow
        linea = ""
        For j = 1 To lastCol
            If j > 1 Then linea = linea & ","
            ' Obtener valor de la celda y escapar comillas
            valorCelda = CStr(ws.Cells(i, j).Value)
            valorCelda = Replace(valorCelda, """", """""")
            linea = linea & "" & valorCelda & ""
        Next j
        ts.WriteLine linea
    Next i

    ts.Close

    ExportarCSV = True
    Exit Function

ErrorHandler:
    ExportarCSV = False
End Function
' ===================================================
' FUNCIONES DE UTILIDAD
' ===================================================

Public Sub ActualizarBarraEstado(mensaje As String)
    Application.StatusBar = SISTEMA_NOMBRE & " - " & mensaje & " - " & Format(Now(), "dd/mm/yyyy HH:mm:ss")
End Sub

Public Sub LimpiarBarraEstado()
    Application.StatusBar = False
End Sub

Public Sub LimpiarDatos()
    On Error GoTo ErrorHandler
    Dim respuesta As Integer
    Dim ws As Worksheet

    ' Confirmación
    respuesta = MsgBox("ADVERTENCIA: Esta acción eliminará todos los datos (excepto cabeceras) de todas las hojas." & vbCrLf & _
               "¿Está seguro de continuar?", vbCritical + vbYesNo, "Confirmar Limpieza")

    If respuesta <> vbYes Then Exit Sub

    ' Limpiar cada hoja
    For Each ws In ThisWorkbook.Worksheets
        If ws.UsedRange.Rows.Count > 1 Then
            ws.Range(ws.Cells(2, 1), ws.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count)).ClearContents
        End If
    Next ws

    MsgBox "Todos los datos han sido limpiados.", vbInformation
    ActualizarBarraEstado "Datos limpiados"

    Exit Sub

ErrorHandler:
    MsgBox "Error durante la limpieza: " & Err.Description, vbCritical
End Sub

Public Sub GenerarReporteSimple()
    On Error GoTo ErrorHandler
    Dim wsUsuarios As Worksheet, wsCompras As Worksheet
    Dim totalUsuarios As Long, totalCompras As Long
    Dim gastoTotal As Double, gastoPromedio As Double
    Dim mensaje As String

    Set wsUsuarios = ThisWorkbook.Sheets("USUARIOS")
    Set wsCompras = ThisWorkbook.Sheets("HISTORIAL_COMPRAS")

    ' Calcular estadísticas
    totalUsuarios = Application.WorksheetFunction.CountA(wsUsuarios.Range("A:A")) - 1
    totalCompras = Application.WorksheetFunction.CountA(wsCompras.Range("A:A")) - 1
    gastoTotal = Application.WorksheetFunction.Sum(wsCompras.Range("E:E"))

    If totalCompras > 0 Then
        gastoPromedio = gastoTotal / totalCompras
    Else
        gastoPromedio = 0
    End If

    ' Construir reporte
    mensaje = "REPORTE DEL SISTEMA" & vbCrLf & vbCrLf
    mensaje = mensaje & "Usuarios registrados: " & totalUsuarios & vbCrLf
    mensaje = mensaje & "Compras registradas: " & totalCompras & vbCrLf
    mensaje = mensaje & "Gasto total: " & Format(gastoTotal, "0.00€") & vbCrLf
    mensaje = mensaje & "Gasto promedio por compra: " & Format(gastoPromedio, "0.00€") & vbCrLf & vbCrLf
    mensaje = mensaje & "Fecha del reporte: " & Format(Now(), "dd/mm/yyyy HH:mm:ss")

    ' Mostrar reporte
    MsgBox mensaje, vbInformation, "Reporte del Sistema"

    Exit Sub

ErrorHandler:
    MsgBox "Error al generar reporte: " & Err.Description, vbCritical
End Sub
' ===================================================
' FUNCIONES DE INTERFAZ
' ===================================================

Public Sub CrearMenuPrincipal()
    On Error Resume Next
    Dim menuBar As CommandBar
    Dim menuItem As CommandBarControl

    ' Eliminar menú anterior si existe
    Application.CommandBars("Comparador IA").Delete

    ' Crear nueva barra de herramientas
    Set menuBar = Application.CommandBars.Add(Name:="Comparador IA", Position:=msoBarFloating, Temporary:=True)

    ' Agregar botones
    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)
    With menuItem
        .Caption = "&Comparar Producto"
        .TooltipText = "Comparar precios de un producto"
        .OnAction = "CompararProducto"
        .FaceId = 172
    End With

    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)
    With menuItem
        .Caption = "&Importar CSV"
        .TooltipText = "Importar datos desde archivo CSV"
        .OnAction = "ImportarDatosCSV"
        .FaceId = 23
    End With

    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)
    With menuItem
        .Caption = "&Exportar CSV"
        .TooltipText = "Exportar datos a archivo CSV"
        .OnAction = "ExportarDatosCSV"
        .FaceId = 308
    End With

    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)
    With menuItem
        .Caption = "&Backup"
        .TooltipText = "Crear copia de seguridad del sistema"
        .OnAction = "CrearBackupCompleto"
        .FaceId = 204
    End With

    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)
    With menuItem
        .Caption = "&Restaurar"
        .TooltipText = "Restaurar datos desde backup"
        .OnAction = "RestaurarDesdeBackup"
        .FaceId = 252
    End With

    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)
    With menuItem
        .Caption = "&Reporte"
        .TooltipText = "Generar reporte del sistema"
        .OnAction = "GenerarReporteSimple"
        .FaceId = 487
    End With

    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)
    With menuItem
        .Caption = "&Limpiar"
        .TooltipText = "Limpiar todos los datos"
        .OnAction = "LimpiarDatos"
        .FaceId = 196
    End With

    ' Mostrar barra de herramientas
    menuBar.Visible = True

    ' Actualizar estado
    ActualizarBarraEstado "Menú creado - Sistema listo"

    Set menuItem = Nothing
    Set menuBar = Nothing
End Sub

Public Sub MostrarMenuPrincipal()
    On Error GoTo ErrorHandler

    Dim respuesta As Variant
    Dim opcionNum As Integer
    Dim salir As Boolean

    salir = False

    Do While Not salir
        ' Crear el mensaje del menГғВә
        Dim menuTexto As String
        menuTexto = "MENU PRINCIPAL - SISTEMA COMPARADOR DE COMPRAS IA" & vbCrLf & vbCrLf & _
                   "1. Comparar precios de producto" & vbCrLf & _
                   "2. Importar datos desde CSV" & vbCrLf & _
                   "3. Exportar datos a CSV" & vbCrLf & _
                   "4. Crear backup completo" & vbCrLf & _
                   "5. Restaurar desde backup" & vbCrLf & _
                   "6. Generar reporte simple" & vbCrLf & _
                   "7. Limpiar todos los datos" & vbCrLf & _
                   "8. Crear menu de herramientas" & vbCrLf & _
                   "9. Salir" & vbCrLf & vbCrLf & _
                   "Seleccione una opcion (1-9):"

        ' Usar InputBox estГғВЎndar
        respuesta = InputBox(menuTexto, "Menu del Sistema", "")

        ' Si el usuario cancela o presiona Cancel
        If respuesta = "" Then
            Exit Sub
        End If

        ' Verificar que sea un nГғВәmero
        If Not IsNumeric(respuesta) Then
            MsgBox "Por favor ingrese un nГғВәmero del 1 al 9.", vbExclamation
            GoTo ContinueLoop
        End If

        ' Convertir a entero
        opcionNum = CInt(respuesta)

        ' Validar rango
        If opcionNum < 1 Or opcionNum > 9 Then
            MsgBox "Opcion no valida. Por favor seleccione 1-9.", vbExclamation
            GoTo ContinueLoop
        End If

        ' Ejecutar opciГғВіn seleccionada
        Select Case opcionNum
            Case 1
                CompararProducto
            Case 2
                ImportarDatosCSV
            Case 3
                ExportarDatosCSV
            Case 4
                CrearBackupCompleto
            Case 5
                RestaurarDesdeBackup
            Case 6
                GenerarReporteSimple
            Case 7
                LimpiarDatos
            Case 8
                CrearMenuPrincipal
            Case 9
                salir = True
                MsgBox "Saliendo del menГғВә principal.", vbInformation
        End Select

ContinueLoop:
    Loop

    Exit Sub

ErrorHandler:
    MsgBox "Error en el menu principal: " & Err.Description, vbCritical
End Sub
' ===================================================
' FUNCIONES DE AUTOMATIZACIÓN
' ===================================================

Private Sub Workbook_Open()
    ' Ejecutar al abrir el libro
    Call InicializarSistema
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Limpiar antes de cerrar
    Call LimpiarBarraEstado
End Sub