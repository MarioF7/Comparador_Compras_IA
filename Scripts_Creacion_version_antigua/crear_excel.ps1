# PowerShell Script para crear Excel del Comparador de Compras IA
# Versión: 1.0 - Sistema Completo

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  CREANDO SISTEMA DE COMPARACION DE COMPRAS IA" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Función para crear el Excel
function Crear-ExcelCompleto {
    param([string]$RutaSalida)
    
    Write-Host "Iniciando creación del Excel..." -ForegroundColor Yellow
    
    try {
        # Cargar ensamblado de Excel
        Add-Type -AssemblyName Microsoft.Office.Interop.Excel
        
        # Crear aplicación Excel
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        
        Write-Host "✓ Aplicación Excel creada" -ForegroundColor Green
        
        # Crear nuevo libro
        $workbook = $excel.Workbooks.Add()
        Write-Host "✓ Libro de trabajo creado" -ForegroundColor Green
        
        # Lista de todas las hojas del sistema
        $hojas = @(
            "00_INSTRUCCIONES",
            "USUARIOS", 
            "CATEGORIAS",
            "UNIDADES_MEDIDA",
            "PRODUCTOS",
            "TIENDAS",
            "PRECIOS_TIENDA",
            "COMPRAS",
            "INTELIGENCIA_USUARIO",
            "LISTAS_COMPRA",
            "ELEMENTOS_LISTA",
            "RUTAS_OPTIMAS",
            "ALERTAS_PRECIO",
            "DASHBOARD",
            "OCR_ENTRADA",
            "CONFIGURACION",
            "LOG_ACTIVIDAD",
            "REPORTES",
            "BACKUP",
            "API_INTEGRACION"
        )
        
        # Renombrar primera hoja
        $workbook.Worksheets.Item(1).Name = $hojas[0]
        Write-Host "✓ Hoja 00_INSTRUCCIONES creada" -ForegroundColor Green
        
        # Crear las demás hojas
        for ($i = 1; $i -lt $hojas.Count; $i++) {
            $newSheet = $workbook.Worksheets.Add()
            $newSheet.Name = $hojas[$i]
            Write-Host "  - Hoja $($hojas[$i]) creada" -ForegroundColor Gray
        }
        
        # ============ CONFIGURAR HOJA DE INSTRUCCIONES ============
        $instSheet = $workbook.Worksheets("00_INSTRUCCIONES")
        
        # Título principal
        $instSheet.Cells(1, 1) = "🚀 SISTEMA DE COMPARACIÓN DE COMPRAS INTELIGENTE"
        $instSheet.Cells(1, 1).Font.Size = 16
        $instSheet.Cells(1, 1).Font.Bold = $true
        $instSheet.Cells(1, 1).Font.Color = RGB(0, 91, 187)
        
        # Configuración inicial
        $instSheet.Cells(3, 1) = "📋 CONFIGURACIÓN INICIAL (Haz esto PRIMERO):"
        $instSheet.Cells(3, 1).Font.Bold = $true
        $instSheet.Cells(3, 1).Font.Size = 12
        
        $instrucciones = @(
            "1️⃣ Ir a hoja USUARIOS y completar TUS DATOS (fila 2)",
            "2️⃣ Ir a hoja TIENDAS y añadir tus tiendas frecuentes",
            "3️⃣ Ir a hoja PRODUCTOS y añadir 10 productos que compres",
            "4️⃣ Ir a hoja CONFIGURACION y pulsar 'INICIALIZAR SISTEMA'",
            "5️⃣ Empezar a registrar compras en hoja COMPRAS"
        )
        
        for ($i = 0; $i -lt $instrucciones.Count; $i++) {
            $instSheet.Cells(4 + $i, 1) = $instrucciones[$i]
            $instSheet.Cells(4 + $i, 1).Font.Size = 11
        }
        
        # Botones principales
        $instSheet.Cells(10, 1) = "⚙️ BOTONES PRINCIPALES:"
        $instSheet.Cells(10, 1).Font.Bold = $true
        
        $botones = @(
            "🛒 Procesar Ticket OCR → Hoja OCR_ENTRADA",
            "🗺️ Generar Ruta Óptima → Hoja LISTAS_COMPRA", 
            "📊 Ver Análisis → Hoja DASHBOARD",
            "🤖 Recomendaciones IA → Hoja INTELIGENCIA_USUARIO",
            "💰 Actualizar Precios → Hoja CONFIGURACION"
        )
        
        for ($i = 0; $i -lt $botones.Count; $i++) {
            $instSheet.Cells(11 + $i, 1) = $botones[$i]
        }
        
        # Notas importantes
        $instSheet.Cells(17, 1) = "⚠️ IMPORTANTE:"
        $instSheet.Cells(17, 1).Font.Bold = $true
        $instSheet.Cells(17, 1).Font.Color = RGB(255, 0, 0)
        
        $notas = @(
            "• HABILITAR MACROS cuando Excel lo solicite",
            "• Guardar como .xlsm (Excel con macros)",
            "• Hacer copias de seguridad semanales",
            "• Los datos se guardan en este archivo"
        )
        
        for ($i = 0; $i -lt $notas.Count; $i++) {
            $instSheet.Cells(18 + $i, 1) = $notas[$i]
        }
        
        # ============ CONFIGURAR HOJA USUARIOS ============
        $userSheet = $workbook.Worksheets("USUARIOS")
        $encabezadosUsuarios = @(
            "ID_Usuario", "Nombre", "Apodo", "Direccion", "Coordenadas",
            "Radio_Busqueda_km", "Prioridad_Compra", "Restricciones",
            "Presupuesto_Mensual_€", "Preferencia_Horario", "Frecuencia_Compra_Sem",
            "Vehiculo_Propio", "Tarjetas_Fidelizacion", "Email", "Telefono",
            "Fecha_Registro", "Ultimo_Acceso", "Activo", "Configuracion_IA", "Notas"
        )
        
        for ($i = 0; $i -lt $encabezadosUsuarios.Count; $i++) {
            $userSheet.Cells(1, $i + 1) = $encabezadosUsuarios[$i]
            $userSheet.Cells(1, $i + 1).Font.Bold = $true
            $userSheet.Cells(1, $i + 1).Interior.Color = RGB(219, 229, 241)
        }
        
        # Datos de ejemplo del usuario principal
        $userSheet.Cells(2, 1) = "U001"
        $userSheet.Cells(2, 2) = "TU NOMBRE AQUÍ"
        $userSheet.Cells(2, 3) = "Yo"
        $userSheet.Cells(2, 4) = "Tu dirección"
        $userSheet.Cells(2, 5) = "40.4168,-3.7038"  # Madrid por defecto
        $userSheet.Cells(2, 6) = 10
        $userSheet.Cells(2, 7) = "Precio-Calidad-Distancia"
        $userSheet.Cells(2, 8) = "Ninguna"
        $userSheet.Cells(2, 9) = 600
        $userSheet.Cells(2, 10) = "Tarde"
        $userSheet.Cells(2, 11) = 3
        $userSheet.Cells(2, 12) = "Sí"
        $userSheet.Cells(2, 13) = "Mercadona,Carrefour"
        $userSheet.Cells(2, 14) = "tucorreo@email.com"
        $userSheet.Cells(2, 15) = "600123456"
        $userSheet.Cells(2, 16) = (Get-Date -Format "dd/MM/yyyy")
        $userSheet.Cells(2, 17) = (Get-Date -Format "dd/MM/yyyy")
        $userSheet.Cells(2, 18) = "Sí"
        $userSheet.Cells(2, 19) = "Avanzado"
        $userSheet.Cells(2, 20) = "Usuario principal - completar datos"
        
        # ============ CONFIGURAR HOJA PRODUCTOS ============
        $prodSheet = $workbook.Worksheets("PRODUCTOS")
        $encabezadosProductos = @(
            "ID_Producto", "Nombre", "Categoria", "Marca", "Peso_Volumen",
            "Unidad", "Precio_Medio_€", "Nutriscore", "Ecologico",
            "Vida_Util_Dias", "Stock_Actual", "Stock_Minimo", "Stock_Maximo",
            "Prioridad", "Ultima_Compra", "Frecuencia_Compra_Dias",
            "Tienda_Preferida", "Notas"
        )
        
        for ($i = 0; $i -lt $encabezadosProductos.Count; $i++) {
            $prodSheet.Cells(1, $i + 1) = $encabezadosProductos[$i]
            $prodSheet.Cells(1, $i + 1).Font.Bold = $true
            $prodSheet.Cells(1, $i + 1).Interior.Color = RGB(234, 241, 221)
        }
        
        # 10 productos de ejemplo
        $productosEjemplo = @(
            @("P001", "Leche Entera", "Lácteos", "Pascual", "1", "L", "0.95", "B", "No", "7", "2", "1", "6", "Alta", "", "3", "Mercadona", "Leche UHT"),
            @("P002", "Huevos M", "Lácteos", "Camperos", "12", "ud", "2.50", "A", "Sí", "28", "6", "6", "24", "Alta", "", "14", "Mercadona", "Huevos gallinas camperas"),
            @("P003", "Pan Integral", "Panadería", "Bimbo", "400", "g", "1.20", "B", "No", "5", "1", "1", "3", "Alta", "", "3", "Carrefour", "Pan de molde"),
            @("P004", "Plátanos", "Frutas", "", "1", "kg", "1.80", "A", "Sí", "7", "2", "2", "5", "Alta", "", "5", "Mercadona", "Plátanos de Canarias"),
            @("P005", "Tomates", "Verduras", "", "1", "kg", "1.50", "A", "Sí", "7", "3", "2", "5", "Alta", "", "7", "Mercadona", "Tomates pera"),
            @("P006", "Pollo", "Carnes", "Campofrío", "1", "kg", "6.50", "B", "No", "3", "1", "0.5", "2", "Media", "", "10", "Carrefour", "Pollo fresco"),
            @("P007", "Arroz", "Legumbres", "Brillante", "1", "kg", "1.10", "A", "No", "365", "2", "1", "4", "Media", "", "60", "DIA", "Arroz largo"),
            @("P008", "Aceite Oliva", "Aceites", "Carbonell", "1", "L", "6.50", "C", "Sí", "365", "1", "1", "3", "Baja", "", "90", "Carrefour", "Aceite virgen extra"),
            @("P009", "Café", "Bebidas", "Marcilla", "250", "g", "4.50", "C", "No", "180", "1", "0.5", "2", "Media", "", "30", "Carrefour", "Café molido"),
            @("P010", "Yogur", "Lácteos", "Danone", "125", "ml", "0.35", "A", "Sí", "30", "12", "8", "24", "Alta", "", "3", "Mercadona", "Yogur natural")
        )
        
        for ($i = 0; $i -lt $productosEjemplo.Count; $i++) {
            for ($j = 0; $j -lt $productosEjemplo[$i].Count; $j++) {
                $prodSheet.Cells($i + 2, $j + 1) = $productosEjemplo[$i][$j]
            }
        }
        
        # ============ CONFIGURAR HOJA TIENDAS ============
        $tiendasSheet = $workbook.Worksheets("TIENDAS")
        $encabezadosTiendas = @(
            "ID_Tienda", "Nombre", "Cadena", "Direccion", "Distancia_km",
            "Tiempo_min", "Valoracion", "Horario", "Aparcamiento",
            "Servicio_Domicilio", "Coste_Envio_€", "Pedido_Minimo_€",
            "Ofertas_Semanales", "Notas"
        )
        
        for ($i = 0; $i -lt $encabezadosTiendas.Count; $i++) {
            $tiendasSheet.Cells(1, $i + 1) = $encabezadosTiendas[$i]
            $tiendasSheet.Cells(1, $i + 1).Font.Bold = $true
            $tiendasSheet.Cells(1, $i + 1).Interior.Color = RGB(242, 222, 222)
        }
        
        # Tiendas de ejemplo
        $tiendasEjemplo = @(
            @("T001", "Mercadona", "Mercadona", "Calle Principal 123", "2.5", "15", "4.2", "9:00-21:30", "Sí", "Sí", "2.95", "30", "15", "La más cercana"),
            @("T002", "Carrefour", "Carrefour", "Avenida Central 456", "3.8", "22", "4.0", "9:00-22:00", "Sí", "Sí", "3.95", "50", "20", "Gran variedad"),
            @("T003", "DIA", "DIA", "Plaza Pequeña 789", "1.2", "8", "3.8", "9:00-21:00", "No", "No", "", "", "10", "Precios bajos"),
            @("T004", "Lidl", "Lidl", "Calle Secundaria 101", "4.5", "25", "4.1", "8:30-21:30", "Sí", "Sí", "2.49", "25", "12", "Calidad-precio"),
            @("T005", "Aldi", "Aldi", "Calle Nueva 202", "5.2", "28", "4.3", "9:00-21:00", "Sí", "Sí", "2.99", "30", "8", "Productos alemanes")
        )
        
        for ($i = 0; $i -lt $tiendasEjemplo.Count; $i++) {
            for ($j = 0; $j -lt $tiendasEjemplo[$i].Count; $j++) {
                $tiendasSheet.Cells($i + 2, $j + 1) = $tiendasEjemplo[$i][$j]
            }
        }
        
        # ============ CONFIGURAR HOJA COMPRAS ============
        $comprasSheet = $workbook.Worksheets("COMPRAS")
        $encabezadosCompras = @(
            "ID_Compra", "Fecha", "ID_Tienda", "ID_Producto", "Producto",
            "Cantidad", "Precio_Total_€", "Precio_Unidad_€", "Precio_KgL_€",
            "Descuento_%", "Promocion", "Metodo_Pago", "Ticket_Numero",
            "Satisfaccion", "Notas"
        )
        
        for ($i = 0; $i -lt $encabezadosCompras.Count; $i++) {
            $comprasSheet.Cells(1, $i + 1) = $encabezadosCompras[$i]
            $comprasSheet.Cells(1, $i + 1).Font.Bold = $true
            $comprasSheet.Cells(1, $i + 1).Interior.Color = RGB(252, 242, 204)
        }
        
        # ============ CONFIGURAR HOJA DASHBOARD ============
        $dashSheet = $workbook.Worksheets("DASHBOARD")
        $dashSheet.Cells(1, 1) = "📊 DASHBOARD - RESUMEN DEL SISTEMA"
        $dashSheet.Cells(1, 1).Font.Size = 14
        $dashSheet.Cells(1, 1).Font.Bold = $true
        
        $metricas = @(
            "Total Usuarios: =COUNTA(USUARIOS!A:A)-1",
            "Total Productos: =COUNTA(PRODUCTOS!A:A)-1",
            "Total Tiendas: =COUNTA(TIENDAS!A:A)-1",
            "Total Compras: =COUNTA(COMPRAS!A:A)-1",
            "Gasto Total: =SUM(COMPRAS!G:G)",
            "Ahorro Estimado: =SUM(COMPRAS!J:J)",
            "Producto Más Comprado: =INDEX(PRODUCTOS!B:B,MATCH(MAX(FREQUENCY(COMPRAS!D:D,COMPRAS!D:D)),FREQUENCY(COMPRAS!D:D,COMPRAS!D:D),0))",
            "Tienda Más Visitada: =INDEX(TIENDAS!B:B,MATCH(MODE(COMPRAS!C:C),TIENDAS!A:A,0))"
        )
        
        $titulosMetricas = @(
            "👥 Total Usuarios",
            "🏷️ Total Productos", 
            "🏪 Total Tiendas",
            "🛒 Total Compras",
            "💰 Gasto Total",
            "💸 Ahorro Estimado",
            "🏆 Producto Más Comprado",
            "📍 Tienda Más Visitada"
        )
        
        for ($i = 0; $i -lt $metricas.Count; $i++) {
            $dashSheet.Cells(3 + $i, 1) = $titulosMetricas[$i]
            $dashSheet.Cells(3 + $i, 1).Font.Bold = $true
            $dashSheet.Cells(3 + $i, 2) = $metricas[$i]
            $dashSheet.Cells(3 + $i, 2).NumberFormat = "0.00"
        }
        
        # ============ APLICAR FORMATO GENERAL ============
        Write-Host "Aplicando formatos..." -ForegroundColor Yellow
        
        $workbook.Worksheets | ForEach-Object {
            $worksheet = $_
            
            # Autoajustar columnas
            $worksheet.UsedRange.EntireColumn.AutoFit()
            
            # Centrar encabezados
            $usedRange = $worksheet.UsedRange
            $headerRow = $usedRange.Rows.Item(1)
            $headerRow.HorizontalAlignment = -4108  # xlCenter
            $headerRow.VerticalAlignment = -4108    # xlCenter
            
            # Aplicar bordes
            $usedRange.Borders.LineStyle = 1        # xlContinuous
            $usedRange.Borders.Weight = 2           # xlThin
            
            # Congelar paneles (primera fila)
            $worksheet.Activate()
            $worksheet.Application.ActiveWindow.SplitRow = 1
            $worksheet.Application.ActiveWindow.FreezePanes = $true
        }
        
        # ============ GUARDAR ARCHIVO ============
        Write-Host "Guardando archivo..." -ForegroundColor Yellow
        
        $filePath = Join-Path (Get-Location) "Comparador_Compras_IA_Completo.xlsm"
        $workbook.SaveAs($filePath, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled
        
        Write-Host "✅ Archivo Excel creado exitosamente: $filePath" -ForegroundColor Green
        
        # Cerrar todo
        $workbook.Close($true)
        $excel.Quit()
        
        # Liberar recursos COM
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        return $true
        
    } catch {
        Write-Host "❌ ERROR al crear Excel: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Detalles: $($_.ScriptStackTrace)" -ForegroundColor DarkGray
        return $false
    }
}

# Función RGB para colores Excel
function RGB {
    param([int]$R, [int]$G, [int]$B)
    return $R + ($G * 256) + ($B * 65536)
}

# ============ EJECUCIÓN PRINCIPAL ============
Write-Host "Iniciando creación del sistema..." -ForegroundColor Yellow
Write-Host ""

# Verificar si Excel está instalado
try {
    $excelCheck = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excelCheck.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelCheck) | Out-Null
    Write-Host "✓ Microsoft Excel encontrado" -ForegroundColor Green
} catch {
    Write-Host "❌ ERROR: Microsoft Excel no está instalado o no está disponible" -ForegroundColor Red
    Write-Host "Por favor, instala Microsoft Excel y vuelve a ejecutar el script." -ForegroundColor Yellow
    pause
    exit 1
}

# Crear el archivo Excel
$resultado = Crear-ExcelCompleto -RutaSalida (Get-Location)

if ($resultado) {
    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Green
    Write-Host "✅ SISTEMA CREADO EXITOSAMENTE" -ForegroundColor Green
    Write-Host "==========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "ARCHIVO PRINCIPAL: Comparador_Compras_IA_Completo.xlsm" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "PASOS SIGUIENTES:" -ForegroundColor Yellow
    Write-Host "1. Abrir el archivo Excel creado" -ForegroundColor White
    Write-Host "2. HABILITAR MACROS cuando se solicite" -ForegroundColor White
    Write-Host "3. Ir a hoja '00_INSTRUCCIONES' y seguir los pasos" -ForegroundColor White
    Write-Host ""
    
    # Preguntar si abrir el archivo
    $abrir = Read-Host "¿Deseas abrir el archivo ahora? (S/N)"
    if ($abrir -eq "S" -or $abrir -eq "s") {
        Start-Process "Comparador_Compras_IA_Completo.xlsm"
    }
} else {
    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Red
    Write-Host "❌ ERROR AL CREAR EL SISTEMA" -ForegroundColor Red
    Write-Host "==========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Por favor, intenta:" -ForegroundColor Yellow
    Write-Host "1. Cerrar todos los archivos de Excel abiertos" -ForegroundColor White
    Write-Host "2. Ejecutar como administrador" -ForegroundColor White
    Write-Host "3. Verificar permisos de escritura" -ForegroundColor White
}

Write-Host ""
Write-Host "Presiona cualquier tecla para salir..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")