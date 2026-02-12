Option Explicit

' =============================================
' CONFIGURACIÓN GLOBAL DEL SISTEMA
' =============================================

Public Const APP_VERSION As String = "3.5.0"
Public Const APP_NAME As String = "Comparador de Compras IA"

' Nombres de hojas
Public Const SHEET_USUARIOS As String = "USUARIOS"
Public Const SHEET_PRODUCTOS As String = "PRODUCTOS"
Public Const SHEET_TIENDAS As String = "TIENDAS"
Public Const SHEET_PRECIOS As String = "PRECIOS"
Public Const SHEET_COMPARATIVA As String = "COMPARATIVA"
Public Const SHEET_HISTORIAL As String = "HISTORIAL_COMPRAS"
Public Const SHEET_PREFERENCIAS As String = "PREFERENCIAS_IA"
Public Const SHEET_DASHBOARD As String = "DASHBOARD"

' Pesos por defecto para el algoritmo (0-1)
Public Const DEFAULT_WEIGHT_PRICE As Double = 0.5
Public Const DEFAULT_WEIGHT_DISTANCE As Double = 0.3
Public Const DEFAULT_WEIGHT_RATING As Double = 0.2

' Colores personalizados (RGB)
Public Const COLOR_SUCCESS As Long = 5287936   ' Verde oscuro
Public Const COLOR_WARNING As Long = 65535      ' Amarillo
Public Const COLOR_ERROR As Long = 255          ' Rojo
Public Const COLOR_HIGHLIGHT As Long = 15773696 ' Azul claro

' Tipos de transporte
Public Enum TransportType
    Coche = 1
    TransportePublico = 2
    Andando = 3
    Bicicleta = 4
End Enum

' Prefijos para IDs automáticos
Public Const ID_PREFIX_PRODUCTO As String = "PROD"
Public Const ID_PREFIX_TIENDA As String = "TND"
Public Const ID_PREFIX_PRECIO As String = "PRC"
Public Const ID_PREFIX_COMPARATIVA As String = "CMP"

' Rutas relativas del proyecto
Public g_strRutaProyecto As String
Public g_strRutaBackup As String
Public g_strRutaReportes As String

' =============================================
' INICIALIZACIÓN AL ABRIR EL LIBRO
' =============================================
Sub InicializarSistema()
    g_strRutaProyecto = ThisWorkbook.Path
    g_strRutaBackup = g_strRutaProyecto & "\Data_Backup\Automatico\"
    g_strRutaReportes = g_strRutaProyecto & "\Reportes\"
    
    Call CrearMenuPersonalizado
    Call VerificarHojas
End Sub

Private Sub CrearMenuPersonalizado()
    On Error Resume Next
    ' Eliminar menú anterior si existe
    Application.CommandBars("Comparador IA").Delete
    On Error GoTo 0
    
    ' Crear nuevo menú
    Dim menuBar As CommandBar
    Set menuBar = Application.CommandBars.Add(Name:="Comparador IA", _
                                              Position:=msoBarTop, _
                                              Temporary:=True)
    menuBar.Visible = True
    
    ' Botón: Alta de Producto
    With menuBar.Controls.Add(Type:=msoControlButton)
        .Caption = "Alta Producto"
        .FaceId = 160
        .OnAction = "AbrirAltaProducto"
        .TooltipText = "Añadir nuevo producto"
    End With
    
    ' Botón: Alta Tienda
    With menuBar.Controls.Add(Type:=msoControlButton)
        .Caption = "Alta Tienda"
        .FaceId = 161
        .OnAction = "AbrirAltaTienda"
        .TooltipText = "Añadir nueva tienda"
    End With
    
    ' Botón: Alta Precio
    With menuBar.Controls.Add(Type:=msoControlButton)
        .Caption = "Alta Precio"
        .FaceId = 162
        .OnAction = "AbrirAltaPrecio"
        .TooltipText = "Registrar precio de producto en tienda"
    End With
    
    ' 4. Botón: Comparar
    With menuBar.Controls.Add(Type:=msoControlButton)
        .Caption = "Comparar Precios"
        .FaceId = 163
        .OnAction = "AbrirComparar"
        .TooltipText = "Comparar productos entre tiendas"
        .BeginGroup = True
    End With
    
    ' 5. Botón: Dashboard
    With menuBar.Controls.Add(Type:=msoControlButton)
        .Caption = "Dashboard"
        .FaceId = 164
        .OnAction = "MostrarDashboard"
        .TooltipText = "Ver panel de control"
    End With
End Sub

Private Sub VerificarHojas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_DASHBOARD Then Exit Sub
    Next ws
    ' Si no existe, crear hoja DASHBOARD
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = SHEET_DASHBOARD
    Call InicializarDashboard
End Sub

Sub InicializarDashboard()
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets(SHEET_DASHBOARD)
    On Error GoTo 0
    
    If wsDash Is Nothing Then Exit Sub
    
    ' Limpiar hoja
    wsDash.Cells.Clear
    
    ' Título
    With wsDash.Range("A1")
        .Value = "PANEL DE CONTROL - COMPARADOR DE COMPRAS IA"
        .Font.Size = 16
        .Font.Bold = True
    End With
    
    ' Indicadores
    wsDash.Range("A3").Value = "Resumen General"
    wsDash.Range("A3").Font.Bold = True
    
    ' --- Fórmulas usando Rangos (Hoja!Columna) ---
    ' Usamos COUNTA y restamos 1 para no contar el encabezado
    
    wsDash.Range("A4").Value = "Total Productos:"
    wsDash.Range("B4").Formula = "=COUNTA('" & SHEET_PRODUCTOS & "'!A:A)-1"
    
    wsDash.Range("A5").Value = "Total Tiendas:"
    wsDash.Range("B5").Formula = "=COUNTA('" & SHEET_TIENDAS & "'!A:A)-1"
    
    wsDash.Range("A6").Value = "Total Precios Registrados:"
    wsDash.Range("B6").Formula = "=COUNTA('" & SHEET_PRECIOS & "'!A:A)-1"
    
    ' Formato y ajuste
    wsDash.Columns("A:B").AutoFit
    wsDash.Range("B4:B6").HorizontalAlignment = -4152 ' xlRight
End Sub

Sub MostrarDashboard()
    ThisWorkbook.Worksheets(SHEET_DASHBOARD).Activate
End Sub
Sub AbrirAltaProducto(): frmAltaProducto.Show: End Sub
Sub AbrirAltaTienda():   frmAltaTienda.Show:   End Sub
Sub AbrirAltaPrecio():   frmAltaPrecio.Show:   End Sub
Sub AbrirComparar():     frmComparar.Show:     End Sub