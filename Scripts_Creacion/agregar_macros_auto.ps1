# agregar_macros_auto.ps1
# Script PowerShell para importar macros automáticamente

param(
    [string]$ProjectPath = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

function Write-Log {
    param(
        [string]$Message,
        [string]$Type = "INFO"
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    switch ($Type) {
        "SUCCESS" { Write-Host "[$timestamp] [OK] $Message" -ForegroundColor Green }
        "ERROR" { Write-Host "[$timestamp] [!] $Message" -ForegroundColor Red }
        "WARNING" { Write-Host "[$timestamp] [*] $Message" -ForegroundColor Yellow }
        default { Write-Host "[$timestamp] [*] $Message" -ForegroundColor Cyan }
    }
}

try {
    Write-Log "Iniciando importacion de macros..." -Type "INFO"
    
    # Definir rutas
    $excelPath = Join-Path $ProjectPath "..\Comparador_Compras_IA\Comparador_Compras_IA_Completo.xlsm"
    $modulePath = Join-Path $ProjectPath "modulo_macros.bas"
    
    # Verificar que el archivo Excel existe
    if (-not (Test-Path $excelPath)) {
        Write-Log "No se encuentra el archivo Excel: $excelPath" -Type "ERROR"
        Write-Log "Creando archivo Excel primero..." -Type "WARNING"
        
        # Ejecutar crear_excel.ps1 primero
        $createExcelScript = Join-Path $ProjectPath "crear_excel.ps1"
        if (Test-Path $createExcelScript) {
            & powershell -ExecutionPolicy Bypass -File $createExcelScript
        } else {
            throw "No se encuentra crear_excel.ps1"
        }
    }
    
    # Verificar que el módulo existe
    if (-not (Test-Path $modulePath)) {
        Write-Log "Creando modulo de macros..." -Type "INFO"
        
        # Crear el código del módulo
        $moduleCode = @"
Attribute VB_Name = "ModuloMacros"
Option Explicit

' MACROS PRINCIPALES DEL SISTEMA COMPARADOR DE COMPRAS IA

Sub CargarDatosEjemplo()
    ' Macro para cargar datos de ejemplo en las hojas
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Application.ScreenUpdating = False
    
    ' Cargar datos en hoja USUARIOS
    Set ws = ThisWorkbook.Sheets("USUARIOS")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow = 1 Then ' Solo encabezados
        ' Usuario 1
        ws.Cells(2, 1).Value = "USR001"
        ws.Cells(2, 2).Value = "Juan Perez"
        ws.Cells(2, 3).Value = "juan.perez@email.com"
        ws.Cells(2, 4).Value = "600123456"
        ws.Cells(2, 5).Value = "Calle Mayor 123"
        ws.Cells(2, 6).Value = "Madrid"
        ws.Cells(2, 7).Value = "28001"
        ws.Cells(2, 8).Value = 40.4168
        ws.Cells(2, 9).Value = -3.7038
        ws.Cells(2, 10).Value = 10
        ws.Cells(2, 11).Value = "Coche"
        ws.Cells(2, 12).Value = "Hacendado,Dia"
        ws.Cells(2, 13).Value = "Lacteos,Panaderia"
        ws.Cells(2, 14).Value = "Lactosa"
        ws.Cells(2, 15).Value = "[]"
        ws.Cells(2, 16).Value = Date
        
        ' Usuario 2
        ws.Cells(3, 1).Value = "USR002"
        ws.Cells(3, 2).Value = "Maria Garcia"
        ws.Cells(3, 3).Value = "maria.garcia@email.com"
        ws.Cells(3, 4).Value = "600654321"
        ws.Cells(3, 5).Value = "Avenida Diagonal 456"
        ws.Cells(3, 6).Value = "Barcelona"
        ws.Cells(3, 7).Value = "08001"
        ws.Cells(3, 8).Value = 41.3851
        ws.Cells(3, 9).Value = 2.1734
        ws.Cells(3, 10).Value = 5
        ws.Cells(3, 11).Value = "Andando"
        ws.Cells(3, 12).Value = "Mercadona,Carrefour"
        ws.Cells(3, 13).Value = "Frutas,Verduras"
        ws.Cells(3, 14).Value = "Gluten"
        ws.Cells(3, 15).Value = "[]"
        ws.Cells(3, 16).Value = Date
    End If
    
    ' Cargar datos en hoja PRODUCTOS
    Set ws = ThisWorkbook.Sheets("PRODUCTOS")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow = 1 Then
        ' Producto 1
        ws.Cells(2, 1).Value = "PROD001"
        ws.Cells(2, 2).Value = "Leche Entera UHT"
        ws.Cells(2, 3).Value = "Lactis liquidum"
        ws.Cells(2, 4).Value = "Alimentacion"
        ws.Cells(2, 5).Value = "Lacteos"
        ws.Cells(2, 6).Value = "Hacendado"
        ws.Cells(2, 7).Value = "Leche entera UHT 1L"
        ws.Cells(2, 8).Value = "Litro"
        ws.Cells(2, 9).Value = 1
        ws.Cells(2, 10).Value = "Calorias: 620kcal/1000ml"
        ws.Cells(2, 11).Value = "Lactosa"
        ws.Cells(2, 12).Value = "Leche entera"
        ws.Cells(2, 13).Value = "http://example.com/info_leche"
        ws.Cells(2, 14).Value = "http://example.com/nutri_leche"
        ws.Cells(2, 15).Value = ""
        ws.Cells(2, 16).Value = ""
        ws.Cells(2, 17).Value = "lacteo,fresco"
        ws.Cells(2, 18).Value = Date
        
        ' Producto 2
        ws.Cells(3, 1).Value = "PROD002"
        ws.Cells(3, 2).Value = "Arroz Bomba"
        ws.Cells(3, 3).Value = "Oryza sativa"
        ws.Cells(3, 4).Value = "Alimentacion"
        ws.Cells(3, 5).Value = "Arroces"
        ws.Cells(3, 6).Value = "Dia"
        ws.Cells(3, 7).Value = "Arroz bomba extra 1kg"
        ws.Cells(3, 8).Value = "Kilogramo"
        ws.Cells(3, 9).Value = 1
        ws.Cells(3, 10).Value = "Calorias: 350kcal/100g"
        ws.Cells(3, 11).Value = ""
        ws.Cells(3, 12).Value = "Arroz"
        ws.Cells(3, 13).Value = "http://example.com/info_arroz"
        ws.Cells(3, 14).Value = "http://example.com/nutri_arroz"
        ws.Cells(3, 15).Value = ""
        ws.Cells(3, 16).Value = ""
        ws.Cells(3, 17).Value = "arroz,basico"
        ws.Cells(3, 18).Value = Date
    End If
    
    ' Cargar datos en hoja TIENDAS
    Set ws = ThisWorkbook.Sheets("TIENDAS")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow = 1 Then
        ' Tienda 1
        ws.Cells(2, 1).Value = "STORE001"
        ws.Cells(2, 2).Value = "Mercadona Centro"
        ws.Cells(2, 3).Value = "Mercadona"
        ws.Cells(2, 4).Value = "Calle Gran Via 123"
        ws.Cells(2, 5).Value = "Madrid"
        ws.Cells(2, 6).Value = "28013"
        ws.Cells(2, 7).Value = "Madrid"
        ws.Cells(2, 8).Value = "Espana"
        ws.Cells(2, 9).Value = 40.4192
        ws.Cells(2, 10).Value = -3.7055
        ws.Cells(2, 11).Value = "09:00-21:00"
        ws.Cells(2, 12).Value = "Supermercado"
        ws.Cells(2, 13).Value = "Si"
        ws.Cells(2, 14).Value = 4.2
        ws.Cells(2, 15).Value = 150
        ws.Cells(2, 16).Value = "2020-01-15"
        
        ' Tienda 2
        ws.Cells(3, 1).Value = "STORE002"
        ws.Cells(3, 2).Value = "Carrefour Express"
        ws.Cells(3, 3).Value = "Carrefour"
        ws.Cells(3, 4).Value = "Paseo de Gracia 456"
        ws.Cells(3, 5).Value = "Barcelona"
        ws.Cells(3, 6).Value = "08007"
        ws.Cells(3, 7).Value = "Barcelona"
        ws.Cells(3, 8).Value = "Espana"
        ws.Cells(3, 9).Value = 41.3917
        ws.Cells(3, 10).Value = 2.1649
        ws.Cells(3, 11).Value = "08:00-22:00"
        ws.Cells(3, 12).Value = "Supermercado"
        ws.Cells(3, 13).Value = "Si"
        ws.Cells(3, 14).Value = 4.0
        ws.Cells(3, 15).Value = 89
        ws.Cells(3, 16).Value = "2019-05-20"
    End If
    
    Application.ScreenUpdating = True
    MsgBox "Datos de ejemplo cargados exitosamente.", vbInformation, "Carga completada"
End Sub

Sub LimpiarDatos()
    ' Macro para limpiar los datos de ejemplo
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "CONFIG" Then
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            If lastRow > 1 Then
                ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, ws.Columns.Count)).ClearContents
            End If
        End If
    Next ws
    
    Application.ScreenUpdating = True
    MsgBox "Datos limpiados exitosamente.", vbInformation, "Limpieza completada"
End Sub

Sub CrearMenu()
    ' Macro para crear un menu de acceso rapido
    Dim cb As CommandBar
    Dim ctrl As CommandBarControl
    
    ' Eliminar menu si ya existe
    On Error Resume Next
    Application.CommandBars("Comparador IA").Delete
    On Error GoTo 0
    
    ' Crear nuevo menu
    Set cb = Application.CommandBars.Add(Name:="Comparador IA", _
                                          Position:=msoBarTop, _
                                          MenuBar:=False, _
                                          Temporary:=True)
    
    ' Agregar botones al menu
    Set ctrl = cb.Controls.Add(Type:=msoControlButton)
    With ctrl
        .Caption = "Cargar Datos Ejemplo"
        .OnAction = "CargarDatosEjemplo"
        .FaceId = 217
    End With
    
    Set ctrl = cb.Controls.Add(Type:=msoControlButton)
    With ctrl
        .Caption = "Limpiar Datos"
        .OnAction = "LimpiarDatos"
        .FaceId = 96
    End With
    
    cb.Visible = True
    MsgBox "Menu de Comparador IA creado exitosamente.", vbInformation, "Menu creado"
End Sub

Sub Auto_Open()
    ' Macro que se ejecuta automaticamente al abrir el archivo
    Call CrearMenu
    MsgBox "Bienvenido al Sistema Comparador de Compras IA" & vbNewLine & _
           "Use el menu 'Comparador IA' para cargar datos de ejemplo.", _
           vbInformation, "Bienvenido"
End Sub
"@
        
        # Guardar el módulo como archivo .bas (ANSI)
        $moduleCode | Out-File -FilePath $modulePath -Encoding ASCII
        Write-Log "Modulo de macros creado: $modulePath" -Type "SUCCESS"
    }
    
    Write-Log "Para importar las macros, por favor:" -Type "INFO"
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "PASOS PARA IMPORTAR MACROS MANUALMENTE:" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "1. Abra el archivo Excel: $excelPath" -ForegroundColor White
    Write-Host "2. Presione ALT + F11 para abrir el Editor de VBA" -ForegroundColor White
    Write-Host "3. Vaya a Archivo -> Importar archivo..." -ForegroundColor White
    Write-Host "4. Seleccione: $modulePath" -ForegroundColor White
    Write-Host "5. Guarde el archivo Excel (CTRL+S)" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    
    # Crear archivo de instrucciones
    $instrucciones = @"
INSTRUCCIONES PARA IMPORTAR MACROS

1. Abra el archivo Excel: $excelPath

2. Acceda al Editor de VBA:
   - Presione ALT + F11
   O
   - Vaya a la pestaña 'Desarrollador' -> 'Visual Basic'

3. Importe el modulo:
   - En el Editor VBA, vaya a 'Archivo' -> 'Importar archivo...'
   - Busque y seleccione: $modulePath
   - Haga clic en 'Abrir'

4. Verifique la importacion:
   - En el panel izquierdo (Explorador de proyectos)
   - Debe aparecer un modulo llamado 'ModuloMacros'

5. Guarde el archivo:
   - Presione CTRL + S
   - Cierre el Editor VBA (ALT + Q)

6. Para usar las macros:
   - Deberia aparecer una nueva barra de menu llamada 'Comparador IA'
   - O puede ejecutarlas desde: Desarrollador -> Macros

MACROS DISPONIBLES:
- CargarDatosEjemplo: Carga datos de ejemplo
- LimpiarDatos: Elimina todos los datos
- CrearMenu: Crea la barra de menu
- Auto_Open: Se ejecuta automaticamente al abrir

NOTA: Si no ve la pestaña 'Desarrollador':
1. Haga clic derecho en la cinta de opciones
2. Seleccione 'Personalizar la cinta de opciones'
3. Marque la casilla 'Desarrollador'
4. Haga clic en 'Aceptar'
"@
    
    $instruccionesPath = Join-Path $ProjectPath "INSTRUCCIONES_MACROS.txt"
    $instrucciones | Out-File -FilePath $instruccionesPath -Encoding UTF8 -Force
    
    Write-Log "Instrucciones detalladas guardadas en: $instruccionesPath" -Type "SUCCESS"
    
} catch {
    Write-Log "Error: $($_.Exception.Message)" -Type "ERROR"
}

Write-Log "Proceso completado. Siga las instrucciones para importar las macros." -Type "SUCCESS"