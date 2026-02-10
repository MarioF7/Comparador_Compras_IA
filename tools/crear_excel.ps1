param(
    [Parameter(Mandatory=$false)]
    [string]$ProjectPath,
    
    [Parameter(Mandatory=$false)]
    [switch]$Force,  # Valor por defecto: $false (si no se usa)
    
    [Parameter(Mandatory=$false)]
    [switch]$Silent = $true   # Valor por defecto: $false (si no se usa)
)

# ===================================================
# CREAR_EXCEL.PS1 - Sistema Comparador de Compras IA
# Versión: 4.0.0 - Profesional
# Autor: Sistema IA
# ===================================================

# Configuración de codificación UTF-8 con BOM
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Si ProjectPath está vacío, calculamos la ruta por defecto aquí abajo
if ([string]::IsNullOrWhiteSpace($ProjectPath)) {
    $ProjectPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}

# ===================================================
# CONFIGURACIÓN GLOBAL
# ===================================================
$VERSION = "4.0.0"
$GLOBAL_ERRORS = 0
$EXCEL_AVAILABLE = $false
$START_TIME = Get-Date

# Rutas
$PROJECT_ROOT = Join-Path (Split-Path $ProjectPath -Parent) "Comparador_Compras_IA"
$EXCEL_FILE = Join-Path $PROJECT_ROOT "Comparador_Compras_IA_Completo.xlsm"
$LOG_DIR = Join-Path $PROJECT_ROOT "Logs"
$LOG_FILE = Join-Path $LOG_DIR "crear_excel_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$BACKUP_DIR = Join-Path $PROJECT_ROOT "Data_Backup"

Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host "  INICIANDO CREACION DE EXCEL" -ForegroundColor Cyan
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "Directorio del proyecto: $PROJECT_ROOT" -ForegroundColor Yellow
Write-Host "Archivo Excel a crear: $EXCEL_FILE" -ForegroundColor Yellow

if ((-not $Silent) -or $ForcePause) {
    Write-Host "`nPresiona una tecla para comenzar..." -ForegroundColor Gray
	[Console]::ReadKey($true) | Out-Null
}

# ===================================================
# FUNCIONES DE UTILIDAD
# ===================================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO",
        [bool]$ConsoleOutput = $true
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $logEntry = "$timestamp [$Level] $Message"
    
    # Guardar en archivo de log
    try {
        Add-Content -Path $LOG_FILE -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {
        # Si falla el log, continuar
    }
    
    # Mostrar en consola si no es modo silencioso
    if ($ConsoleOutput -and (-not $Silent)) {
        switch ($Level) {
            "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
            "ERROR"   { Write-Host $logEntry -ForegroundColor Red }
            "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
            "DEBUG"   { Write-Host $logEntry -ForegroundColor Gray }
            default   { Write-Host $logEntry -ForegroundColor Cyan }
        }
    }
}

function Pause-Script {
    param(
        [string]$Message = "Presiona una tecla para continuar...",
        [bool]$ForcePause = $false
    )
    
    if ((-not $Silent) -or $ForcePause) {
        Write-Host "`n$Message" -ForegroundColor Magenta
        [Console]::ReadKey($true) | Out-Null
    }
}

function Test-ExcelInstalled {
    Write-Host "`n[PASO 1/7] Verificando si Excel está instalado..." -ForegroundColor Cyan
    Pause-Script -Message "Verificando Excel. Presiona una tecla..."
    
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        $version = $excel.Version
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
        Write-Log "Excel $version detectado correctamente" -Level "SUCCESS"
        Write-Host "✓ Excel $version detectado" -ForegroundColor Green
        return $true
    } catch {
        Write-Log "Excel no está instalado o no es accesible: $($_.Exception.Message)" -Level "WARNING"
        Write-Host "✗ Excel no está instalado o no es accesible" -ForegroundColor Red
        Write-Host "  Se crearán archivos CSV como alternativa" -ForegroundColor Yellow
        return $false
    }
}

# NUEVA FUNCIÓN: Desbloquear archivo Excel
function Unlock-ExcelFile {
    param([string]$FilePath)
    
    Write-Host "`nDesbloqueando archivo Excel..." -ForegroundColor Cyan
    
    try {
        # 1. Quitar atributo de solo lectura
        if (Test-Path $FilePath) {
            $file = Get-Item -Path $FilePath
            if ($file.IsReadOnly) {
                $file.IsReadOnly = $false
                Write-Host "✓ Atributo de solo lectura removido" -ForegroundColor Green
            }
        }
        
        # 2. Eliminar Zone.Identifier (bloqueo de seguridad)
        $zoneIdentifier = "$($FilePath):Zone.Identifier"
        if (Test-Path -LiteralPath $zoneIdentifier) {
            Remove-Item -LiteralPath $zoneIdentifier -Force
            Write-Host "✓ Bloqueo de seguridad (Zone.Identifier) removido" -ForegroundColor Green
        }
        
        # 3. Usar Unblock-File si está disponible (PowerShell 3.0+)
        if (Get-Command Unblock-File -ErrorAction SilentlyContinue) {
            Unblock-File -Path $FilePath -ErrorAction SilentlyContinue
            Write-Host "✓ Archivo desbloqueado con Unblock-File" -ForegroundColor Green
        }
        
        # 4. Verificar permisos
        $acl = Get-Acl -Path $FilePath
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $currentUser,
            "FullControl",
            "Allow"
        )
        $acl.SetAccessRule($accessRule)
        Set-Acl -Path $FilePath -AclObject $acl
        Write-Host "✓ Permisos establecidos para el usuario actual" -ForegroundColor Green
        
        return $true
    } catch {
        Write-Host "✗ Error al desbloquear archivo: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Log "Error al desbloquear archivo: $($_.Exception.Message)" -Level "WARNING"
        return $false
    }
}

function Create-ExcelStructure {
    param(
        [object]$Excel,
        [object]$Workbook
    )
    
    Write-Host "`n[PASO 3/7] Creando estructura completa de hojas..." -ForegroundColor Cyan
    Pause-Script -Message "Creando estructura de hojas. Presiona una tecla..."
    
    # Definición completa de hojas según documentación
    $sheetsConfig = @(
        @{
            Name = "USUARIOS"
            Headers = @(
                "UserID", "Nombre", "Email", "Teléfono", "Dirección", "Ciudad", "CP",
                "Coord_Lat", "Coord_Lon", "Radio_Búsqueda_KM", "Pref_Transporte",
                "Pref_Marcas", "Pref_Categorías", "Restricciones", "Presupuesto_Mensual",
                "Historial_Búsqueda", "Fecha_Registro", "Último_Acceso", "Activo", "Nivel_Usuario"
            )
            ColumnWidths = @(12, 25, 25, 15, 35, 15, 10, 12, 12, 8, 15, 20, 20, 25, 12, 30, 15, 15, 8, 12)
        },
        @{
            Name = "PRODUCTOS"
            Headers = @(
                "ProductID", "Nombre", "Nombre_Científico", "Categoría", "Subcategoría", "Marca", "Descripción",
                "Características", "Unidad_Medida", "Tamaño_Paquete", "Unidades_Paquete", "Peso_Bruto", "Peso_Neto",
                "Dimensiones", "UPC/EAN", "Código_Interno", "URL_Imagen", "URL_Info", "URL_Nutricional",
                "Alérgenos", "Caducidad_Mínima", "Refrigerado", "Congelado", "Orgánico", "Comercio_Justo",
                "Fecha_Alta", "Activo"
            )
            ColumnWidths = @(12, 35, 20, 15, 15, 15, 40, 25, 15, 12, 12, 12, 12, 20, 15, 20, 30, 30, 30, 20, 10, 10, 10, 10, 10, 15, 8)
        },
        @{
            Name = "TIENDAS"
            Headers = @(
                "StoreID", "Nombre_Tienda", "Cadena", "Dirección", "Ciudad", "CP", "Provincia", "País",
                "Coord_Lat", "Coord_Lon", "Horario", "Teléfono", "Email", "Web", "Tipo_Tienda", "Tamaño_Tienda",
                "Servicios", "Parking", "Acceso_Discapacitados", "Wifi_Gratis", "Cajeros_Automáticos", "Farmacia",
                "Valoración_Media", "N_Opiniones", "Fecha_Valoración", "Distancia_Usuario", "Tiempo_Desplazamiento",
                "Coste_Desplazamiento", "Activo"
            )
            ColumnWidths = @(12, 30, 15, 35, 15, 10, 15, 10, 12, 12, 20, 15, 25, 30, 15, 15, 25, 8, 8, 8, 8, 8, 8, 10, 15, 12, 15, 12, 8)
        },
        @{
            Name = "PRECIOS"
            Headers = @(
                "PriceID", "ProductID", "StoreID", "Precio_Unitario", "Precio_Paquete", "Unidad_Medida",
                "Precio_x_KG", "Precio_x_Litro", "Precio_x_Unidad", "Oferta", "Descuento_%", "Precio_Original",
                "Tipo_Oferta", "Fecha_Inicio_Oferta", "Fecha_Fin_Oferta", "Stock", "Cantidad_Stock",
                "Unidades_Mínimas", "Unidades_Máximas", "Fecha_Actualización", "Fuente_Datos", "URL_Oferta",
                "Confianza_Datos", "Historial_Precios"
            )
            ColumnWidths = @(20, 12, 12, 12, 12, 15, 12, 12, 12, 8, 10, 12, 15, 15, 15, 10, 12, 12, 12, 15, 15, 30, 10, 30)
        },
        @{
            Name = "COMPARATIVA"
            Headers = @(
                "ComparativaID", "UserID", "ProductID", "Lista_Productos", "Fecha_Comparación", "Mejor_Precio",
                "Tienda_Mejor_Precio", "Precio_Medio", "Precio_Máximo", "Precio_Mínimo", "Desviación_Estándar",
                "Distancia_Mejor", "Tiempo_Mejor", "Coste_Desplazamiento", "Ahorro_Estimado", "Ahorro_Porcentual",
                "N_Tiendas_Comparadas", "Ruta_Recomendada", "Tiendas_Ruta", "Distancia_Total_Ruta", "Tiempo_Total_Ruta",
                "Coste_Total_Ruta", "Puntuación_Global", "Puntuación_Precio", "Puntuación_Distancia", "Puntuación_Calidad",
                "Recomendación", "Notas"
            )
            ColumnWidths = @(20, 12, 12, 30, 15, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 10, 10, 30, 25, 12, 15, 12, 10, 10, 10, 10, 15, 30)
        },
        @{
            Name = "HISTORIAL_COMPRAS"
            Headers = @(
                "CompraID", "UserID", "StoreID", "Fecha_Compra", "Total_Compra", "Total_Descuentos",
                "Total_Sin_Descuentos", "N_Productos", "N_Items", "Lista_Productos", "Método_Pago", "Tipo_Compra",
                "Ticket_Image", "Ticket_PDF", "Valoración_Compra", "Valoración_Productos", "Valoración_Atención",
                "Valoración_Tienda", "Comentarios", "Problemas", "Sugerencias", "Fecha_Registro"
            )
            ColumnWidths = @(20, 12, 12, 15, 12, 12, 12, 10, 10, 30, 15, 15, 30, 30, 10, 10, 10, 10, 40, 30, 30, 15)
        },
        @{
            Name = "PREFERENCIAS_IA"
            Headers = @(
                "PrefID", "UserID", "Categoría_Favorita", "Subcategoría_Favorita", "Marca_Favorita", "Tienda_Favorita",
                "Gasto_Promedio_Mes", "Frecuencia_Compra", "Día_Preferido_Compra", "Hora_Preferida", "Sensibilidad_Precio",
                "Sensibilidad_Calidad", "Sensibilidad_Distancia", "Sensibilidad_Tiempo", "Sensibilidad_Marca",
                "Tolerancia_Desplazamiento", "Presupuesto_Máx_Producto", "Preferencia_Ofertas", "Preferencia_Ecológico",
                "Preferencia_Local", "Historial_Recomendaciones", "Acierto_Recomendaciones", "Última_Actualización",
                "Modelo_IA", "Versión_Modelo"
            )
            ColumnWidths = @(20, 12, 20, 20, 15, 15, 12, 12, 15, 12, 10, 10, 10, 10, 10, 12, 12, 8, 8, 8, 30, 10, 15, 20, 15)
        }
    )
    
    Write-Host "Creando las siguientes hojas:" -ForegroundColor Yellow
    foreach ($config in $sheetsConfig) {
        Write-Host "  • $($config.Name)" -ForegroundColor White
    }
    
    Pause-Script -Message "Lista de hojas a crear. Presiona una tecla para proceder..."
    
    # Crear cada hoja
    foreach ($config in $sheetsConfig) {
        try {
            Write-Host "Creando hoja: $($config.Name)..." -ForegroundColor Gray
            
            # Crear hoja
            $worksheet = $Workbook.Worksheets.Add()
            $worksheet.Name = $config.Name
            
            # Agregar encabezados
            for ($i = 0; $i -lt $config.Headers.Count; $i++) {
                $cell = $worksheet.Cells.Item(1, $i + 1)
                $cell.Value = $config.Headers[$i]
                
                # Formato de encabezado
                $cell.Font.Bold = $true
                $cell.Interior.Color = 0xCCE5FF  # Azul claro
                $cell.HorizontalAlignment = -4108  # Centrado
                $cell.VerticalAlignment = -4108
                $cell.Borders.LineStyle = 1
                $cell.Borders.Weight = 2
                
                # Ajustar ancho de columna
                if ($config.ColumnWidths[$i]) {
                    $worksheet.Columns($i + 1).ColumnWidth = $config.ColumnWidths[$i]
                }
            }
            
            # Congelar paneles
            $worksheet.Activate()
            $worksheet.Application.ActiveWindow.SplitRow = 1
            $worksheet.Application.ActiveWindow.FreezePanes = $true
            
            Write-Host "  ✓ Hoja '$($config.Name)' creada" -ForegroundColor Green
            
        } catch {
            Write-Host "  ✗ Error al crear hoja $($config.Name): $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "Error al crear hoja $($config.Name): $($_.Exception.Message)" -Level "ERROR"
            $script:GLOBAL_ERRORS++
            Pause-Script -Message "Error detectado. Presiona una tecla para continuar..." -ForcePause $true
        }
    }
    
    # Eliminar hojas por defecto
    Write-Host "`nEliminando hojas por defecto de Excel..." -ForegroundColor Gray
    while ($Workbook.Worksheets.Count -gt $sheetsConfig.Count) {
        try {
            $Workbook.Worksheets.Item(1).Delete()
        } catch {
            break
        }
    }
    
    Write-Host "✓ Estructura de hojas completada" -ForegroundColor Green
}

function Add-FormulasAndValidations {
    param(
        [object]$Workbook
    )
    
    Write-Host "`n[PASO 4/7] Agregando fórmulas y validaciones..." -ForegroundColor Cyan
    Pause-Script -Message "Agregando fórmulas. Presiona una tecla..."
    
    try {
        # Hoja PRECIOS - Fórmulas de cálculo
        $pricesSheet = $Workbook.Worksheets("PRECIOS")
        
        # Fórmula para precio por kg
        $pricesSheet.Range("G2:G1000").Formula = "=IFERROR(IF(F2=""kg"",D2,IF(F2=""g"",D2/1000,"""")),"""")"
        
        # Fórmula para precio por litro
        $pricesSheet.Range("H2:H1000").Formula = "=IFERROR(IF(F2=""litro"",D2,IF(F2=""ml"",D2/1000,"""")),"""")"
        
        # Fórmula para precio por unidad
        $pricesSheet.Range("I2:I1000").Formula = "=IFERROR(IF(F2=""unidad"",D2,""""),"""")"
        
        # Hoja COMPARATIVA - Fórmulas de puntuación
        $compSheet = $Workbook.Worksheets("COMPARATIVA")
        $compSheet.Range("W2:W1000").Formula = "=IFERROR((U2*0.4)+(V2*0.3)+(T2*0.2)+(S2*0.1),0)"
        
        Write-Host "✓ Fórmulas agregadas" -ForegroundColor Green
        
    } catch {
        Write-Host "✗ Error al agregar fórmulas: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error al agregar fórmulas: $($_.Exception.Message)" -Level "ERROR"
        Pause-Script -Message "Error en fórmulas. Presiona una tecla para continuar..." -ForcePause $true
    }
}

function Create-PivotTables {
    param(
        [object]$Workbook
    )
    
    Write-Host "`n[PASO 5/7] Creando tablas dinámicas de análisis..." -ForegroundColor Cyan
    Pause-Script -Message "Creando tablas dinámicas. Presiona una tecla..."
    
    try {
        # Verificar que la hoja PRECIOS existe
        if ($Workbook.Worksheets.Count -eq 0 -or !($Workbook.Worksheets("PRECIOS"))) {
            Write-Host "✗ Hoja PRECIOS no encontrada, omitiendo tablas dinámicas" -ForegroundColor Yellow
            Write-Log "Hoja PRECIOS no encontrada para crear tablas dinámicas" -Level "WARNING"
            return
        }
        
        $pricesSheet = $Workbook.Worksheets("PRECIOS")
        
        # Verificar que hay datos (más de 1 fila, incluyendo encabezados)
        if ($pricesSheet.UsedRange.Rows.Count -le 1) {
            Write-Host "✗ No hay datos en la hoja PRECIOS, omitiendo tablas dinámicas" -ForegroundColor Yellow
            Write-Log "No hay datos en PRECIOS para crear tablas dinámicas" -Level "WARNING"
            return
        }
        
        # Intentar crear caché de tabla dinámica
        $pivotCache = $null
        try {
            $pivotCache = $Workbook.PivotCaches().Create(1, $pricesSheet.UsedRange, 7)
        } catch {
            Write-Host "✗ No se pudo crear caché de tabla dinámica: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Log "Error creando caché de tabla dinámica: $($_.Exception.Message)" -Level "WARNING"
            return
        }
        
        # Crear hoja para análisis
        $pivotSheet = $Workbook.Worksheets.Add()
        $pivotSheet.Name = "ANALISIS_PRECIOS"
        
        # Crear tabla dinámica básica (sin campos)
        $pivotTable = $pivotCache.CreatePivotTable($pivotSheet.Range("A3"), "PivotAnalisisBásico")
        
        # Solo agregar campos si existen
        try {
            # Verificar si el campo "Precio_Unitario" existe
            $priceField = $null
            foreach ($field in $pivotTable.PivotFields()) {
                if ($field.Name -like "*Precio*") {
                    $priceField = $field
                    break
                }
            }
            
            if ($priceField) {
                $priceField.Orientation = 4  # xlDataField
                $priceField.Function = -4136  # xlAverage
            }
        } catch {
            # Si no se pueden agregar campos, continuar con tabla vacía
            Write-Host "  Nota: Tabla dinámica creada sin campos específicos" -ForegroundColor Gray
        }
        
        # Formato básico
        try {
            $pivotTable.TableStyle2 = "PivotStyleLight1"
        } catch {
            # Continuar si falla el formato
        }
        
        Write-Host "✓ Tablas dinámicas básicas creadas" -ForegroundColor Green
        
    } catch {
        Write-Host "✗ Error al crear tablas dinámicas: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Log "Error al crear tablas dinámicas: $($_.Exception.Message)" -Level "WARNING"
        
        # NO pausar aquí - dejar continuar
        Write-Host "  Continuando sin tablas dinámicas..." -ForegroundColor Gray
    }
}

function Create-BackupFile {
    param(
        [string]$SourceFile
    )
    
    $backupFile = Join-Path $BACKUP_DIR "excel_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsm"
    
    try {
        Copy-Item -Path $SourceFile -Destination $backupFile -Force
        # Desbloquear también el backup
        Unlock-ExcelFile -FilePath $backupFile
        Write-Host "✓ Copia de seguridad creada: $backupFile" -ForegroundColor Green
        return $backupFile
    } catch {
        Write-Host "✗ Error al crear backup: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Log "Error al crear backup: $($_.Exception.Message)" -Level "WARNING"
        return $null
    }
}

# ===================================================
# FUNCIÓN PRINCIPAL
# ===================================================

function Main {
    # Encabezado
    if (-not $Silent) {
        Write-Host "`n"
        Write-Host "===================================================" -ForegroundColor Cyan
        Write-Host "  CREANDO EXCEL - SISTEMA COMPARADOR DE COMPRAS IA" -ForegroundColor Cyan
        Write-Host "  Versión: $VERSION" -ForegroundColor Cyan
        Write-Host "===================================================" -ForegroundColor Cyan
        Write-Host "`n"
    }
    
    Write-Log "Iniciando creación de archivo Excel..." -Level "INFO"
    Write-Log "Directorio del proyecto: $PROJECT_ROOT" -Level "INFO"
    
    # Verificar directorios
    Write-Host "`n[PASO 0/7] Preparando directorios..." -ForegroundColor Cyan
    if (-not (Test-Path $LOG_DIR)) {
        New-Item -ItemType Directory -Path $LOG_DIR -Force | Out-Null
        Write-Host "✓ Directorio de logs creado: $LOG_DIR" -ForegroundColor Green
    }
    
    if (-not (Test-Path $BACKUP_DIR)) {
        New-Item -ItemType Directory -Path $BACKUP_DIR -Force | Out-Null
        Write-Host "✓ Directorio de backup creado: $BACKUP_DIR" -ForegroundColor Green
    }
    
    Pause-Script -Message "Directorios preparados. Presiona una tecla..."
    
    # Verificar si Excel existe
    Write-Host "`n[PASO 2/7] Verificando si el archivo Excel ya existe..." -ForegroundColor Cyan
    if (Test-Path $EXCEL_FILE) {
        Write-Host "✗ Archivo Excel ya existe: $EXCEL_FILE" -ForegroundColor Yellow
        
        if ($Force) {
            Write-Host "Forzando recreación (parámetro -Force)" -ForegroundColor Magenta
            
            # Crear backup antes de sobrescribir
            $backup = Create-BackupFile -SourceFile $EXCEL_FILE
            Remove-Item -Path $EXCEL_FILE -Force -ErrorAction SilentlyContinue
            Write-Host "✓ Archivo anterior eliminado" -ForegroundColor Green
        } else {
            Write-Host "Use -Force para recrear el archivo" -ForegroundColor Yellow
            Pause-Script -Message "Archivo ya existe. Presiona una tecla para salir..."
            return
        }
    } else {
        Write-Host "✓ Archivo Excel no existe, se procederá a crear" -ForegroundColor Green
    }
    
    Pause-Script -Message "Verificación de archivos completada. Presiona una tecla..."
    
    # Verificar si Excel está instalado
    $script:EXCEL_AVAILABLE = Test-ExcelInstalled
    
    if (-not $EXCEL_AVAILABLE) {
        Write-Host "`n[ALTERNATIVA] Creando estructura CSV..." -ForegroundColor Cyan
        Pause-Script -Message "Excel no disponible. Creando CSV alternativo. Presiona una tecla..."
        Create-CSVAlternative
        return
    }
    
    # Crear archivo Excel
    Write-Host "`n[INICIANDO CREACION DE EXCEL]" -ForegroundColor Cyan
    Write-Host "================================" -ForegroundColor Cyan
    
    try {
        Write-Host "Inicializando Excel COM Object..." -ForegroundColor Gray
        Pause-Script -Message "Inicializando Excel. Esto puede tardar unos segundos..."
        
        # Crear aplicación Excel
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        $excel.AskToUpdateLinks = $false
        
        Write-Host "✓ Excel inicializado" -ForegroundColor Green
        
        # Crear nuevo libro
        Write-Host "Creando nuevo libro de trabajo..." -ForegroundColor Gray
        $workbook = $excel.Workbooks.Add()
        Write-Host "✓ Libro creado" -ForegroundColor Green
        
        Pause-Script -Message "Excel listo. Presiona una tecla para crear la estructura..."
        
        # Crear estructura de hojas
        Create-ExcelStructure -Excel $excel -Workbook $workbook
        
        Pause-Script -Message "Estructura creada. Presiona una tecla para agregar fórmulas..."
        
        # Agregar fórmulas y validaciones
        Add-FormulasAndValidations -Workbook $workbook
        
        Pause-Script -Message "Fórmulas agregadas. Presiona una tecla para crear tablas dinámicas..."
        
        # Crear tablas dinámicas
        Create-PivotTables -Workbook $workbook
        
        Pause-Script -Message "Tablas dinámicas creadas. Presiona una tecla para proteger hojas..."
        
        # Guardar archivo - MODIFICADO: Guardar sin protección temporal
        Write-Host "`n[PASO 6/7] Guardando archivo Excel..." -ForegroundColor Cyan
        Pause-Script -Message "Guardando archivo. Esto puede tardar unos segundos..."
        
        Write-Host "Guardando en: $EXCEL_FILE" -ForegroundColor Yellow
        
        # Intentar guardar con diferentes métodos si falla
        try {
            # Método 1: Guardar como .xlsm
            $workbook.SaveAs($EXCEL_FILE, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled
            Write-Host "✓ Archivo guardado exitosamente" -ForegroundColor Green
        } catch {
            Write-Host "✗ Error al guardar, intentando método alternativo..." -ForegroundColor Yellow
            try {
                # Método 2: Guardar sin formato específico
                $workbook.SaveAs($EXCEL_FILE)
                Write-Host "✓ Archivo guardado con método alternativo" -ForegroundColor Green
            } catch {
                Write-Host "✗ Error crítico al guardar: $($_.Exception.Message)" -ForegroundColor Red
                throw
            }
        }
        
        # Desbloquear archivo inmediatamente después de guardar
        Write-Host "Desbloqueando archivo para edición..." -ForegroundColor Gray
        $unlockResult = Unlock-ExcelFile -FilePath $EXCEL_FILE
        
        if (-not $unlockResult) {
            Write-Host "✗ Advertencia: No se pudo desbloquear completamente el archivo" -ForegroundColor Yellow
            Write-Host "  Puede que necesites habilitar manualmente la edición" -ForegroundColor Yellow
        }
        
        # Crear backup inicial
        Write-Host "`n[PASO 7/7] Creando copia de seguridad..." -ForegroundColor Cyan
        Create-BackupFile -SourceFile $EXCEL_FILE
        
        # Estadísticas
        Write-Host "`n[ESTADISTICAS]" -ForegroundColor Cyan
        Write-Host "===============" -ForegroundColor Cyan
        
        $fileSize = (Get-Item $EXCEL_FILE).Length / 1MB
        $sheetCount = $workbook.Worksheets.Count
        
        Write-Host "Tamaño del archivo: $($fileSize.ToString('0.00')) MB" -ForegroundColor White
        Write-Host "Número de hojas: $sheetCount" -ForegroundColor White
        
        # Mostrar lista de hojas creadas
        Write-Host "`nHojas creadas:" -ForegroundColor Yellow
        foreach ($ws in $workbook.Worksheets) {
            Write-Host "  • $($ws.Name)" -ForegroundColor White
        }
        
        # Información adicional sobre el desbloqueo
        Write-Host "`n[INFORMACION DE DESBLOQUEO]" -ForegroundColor Cyan
        Write-Host "=============================" -ForegroundColor Cyan
        Write-Host "El archivo ha sido desbloqueado para edición." -ForegroundColor White
        Write-Host "Si aún ves 'solo lectura' al abrir:" -ForegroundColor Yellow
        Write-Host "1. Haz clic en 'Habilitar edición' en la barra amarilla" -ForegroundColor White
        Write-Host "2. O guarda una copia local desde Archivo → Guardar como" -ForegroundColor White
        
        Pause-Script -Message "Estadísticas mostradas. Presiona una tecla para cerrar Excel..."
        
        # Cerrar Excel
        Write-Host "Cerrando Excel..." -ForegroundColor Gray
        $workbook.Close($true)
        $excel.Quit()
        
        # Liberar objetos COM
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        Remove-Variable excel, workbook
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Host "✓ Excel cerrado correctamente" -ForegroundColor Green
        
    } catch {
        Write-Host "`n✗✗✗ ERROR CRITICO ✗✗✗" -ForegroundColor Red
        Write-Host "Error al crear Excel: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Ubicación del error: $($_.ScriptStackTrace)" -ForegroundColor Yellow
        
        Write-Log "Error crítico al crear Excel: $($_.Exception.Message)" -Level "ERROR"
        Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
        $script:GLOBAL_ERRORS++
        
        # Intentar cerrar Excel si está abierto
        try {
            if ($workbook) { $workbook.Close($false) }
            if ($excel) { $excel.Quit() }
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            Remove-Variable excel, workbook -ErrorAction SilentlyContinue
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        } catch {}
        
        Pause-Script -Message "Error crítico. Presiona una tecla para crear alternativa CSV..." -ForcePause $true
        
        # Crear alternativa CSV
        Create-CSVAlternative
    }
}

# ===================================================
# FUNCIÓN ALTERNATIVA CSV
# ===================================================

function Create-CSVAlternative {
    Write-Host "`n[CREANDO ESTRUCTURA CSV ALTERNATIVA]" -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    
    $csvDir = Join-Path $PROJECT_ROOT "CSV_Backup"
    Write-Host "Creando directorio: $csvDir" -ForegroundColor Yellow
    
    New-Item -ItemType Directory -Path $csvDir -Force | Out-Null
    
    Pause-Script -Message "Directorio CSV creado. Presiona una tecla para crear archivos..."
    
    # Definir estructura CSV completa
    $csvStructures = @{
        "USUARIOS.csv" = @"
UserID,Nombre,Email,Teléfono,Dirección,Ciudad,CP,Coord_Lat,Coord_Lon,Radio_Búsqueda_KM,Pref_Transporte,Pref_Marcas,Pref_Categorías,Restricciones,Presupuesto_Mensual,Historial_Búsqueda,Fecha_Registro,Último_Acceso,Activo,Nivel_Usuario
USR001,Juan Pérez,juan.perez@email.com,+34 600111222,Calle Mayor 1 1ºA,Madrid,28013,40.416775,-3.703790,5,Coche,"Nestlé,Danone,Kellogg's","Alimentación,Limpieza","Sin lactosa, Sin gluten",450.00,"[{""producto"":""leche"",""fecha"":""2024-01-15""}]",2024-01-15,2024-01-20 10:30:00,TRUE,Básico
"@
        
        "PRODUCTOS.csv" = @"
ProductID,Nombre,Nombre_Científico,Categoría,Subcategoría,Marca,Descripción,Características,Unidad_Medida,Tamaño_Paquete,Unidades_Paquete,Peso_Bruto,Peso_Neto,Dimensiones,UPC/EAN,Código_Interno,URL_Imagen,URL_Info,URL_Nutricional,Alérgenos,Caducidad_Mínima,Refrigerado,Congelado,Orgánico,Comercio_Justo,Fecha_Alta,Activo
PROD001,Leche Entera UHT,Lactis liquidum,Alimentación,Lácteos,Pascual,Leche entera UHT tratamiento térmico 1L,"Enriquecida con calcio y vitaminas A y D",litro,1.000,1,1050.000,1000.000,"6.5x6.5x18.5 cm",8410100001234,LEC-ENT-UHT-1L,http://example.com/leche.jpg,http://example.com/info_leche,http://example.com/nutri_leche,Lactosa,90,FALSE,FALSE,FALSE,FALSE,2024-01-15,TRUE
"@
        
        "TIENDAS.csv" = @"
StoreID,Nombre_Tienda,Cadena,Dirección,Ciudad,CP,Provincia,País,Coord_Lat,Coord_Lon,Horario,Teléfono,Email,Web,Tipo_Tienda,Tamaño_Tienda,Servicios,Parking,Acceso_Discapacitados,Wifi_Gratis,Cajeros_Automáticos,Farmacia,Valoración_Media,N_Opiniones,Fecha_Valoración,Distancia_Usuario,Tiempo_Desplazamiento,Coste_Desplazamiento,Activo
TND001,Mercadona Alcalá,Mercadona,Calle Alcalá 10,Madrid,28013,Madrid,España,40.417000,-3.703000,"09:00-21:00",912345678,info@mercadona.es,http://www.mercadona.es,Supermercado,Grande,"Delivery,Recogida en tienda,Parking",TRUE,TRUE,FALSE,TRUE,FALSE,4.2,150,2024-01-15,2.5,0:15:00,1.50,TRUE
"@
        
        "PRECIOS.csv" = @"
PriceID,ProductID,StoreID,Precio_Unitario,Precio_Paquete,Unidad_Medida,Precio_x_KG,Precio_x_Litro,Precio_x_Unidad,Oferta,Descuento_%,Precio_Original,Tipo_Oferta,Fecha_Inicio_Oferta,Fecha_Fin_Oferta,Stock,Cantidad_Stock,Unidades_Mínimas,Unidades_Máximas,Fecha_Actualización,Fuente_Datos,URL_Oferta,Confianza_Datos,Historial_Precios
PRC001-PROD001-TND001,PROD001,TND001,1.20,1.20,litro,,1.2000,,TRUE,10.00,1.33,"2x1",2024-01-15,2024-01-31,Alto,50,1,10,2024-01-15 10:30:00,Manual,http://oferta.com/leche,0.95,"[{""fecha"":""2024-01-01"",""precio"":1.33}]"
"@
        
        "COMPARATIVA.csv" = @"
ComparativaID,UserID,ProductID,Lista_Productos,Fecha_Comparación,Mejor_Precio,Tienda_Mejor_Precio,Precio_Medio,Precio_Máximo,Precio_Mínimo,Desviación_Estándar,Distancia_Mejor,Tiempo_Mejor,Coste_Desplazamiento,Ahorro_Estimado,Ahorro_Porcentual,N_Tiendas_Comparadas,Ruta_Recomendada,Tiendas_Ruta,Distancia_Total_Ruta,Tiempo_Total_Ruta,Coste_Total_Ruta,Puntuación_Global,Puntuación_Precio,Puntuación_Distancia,Puntuación_Calidad,Recomendación,Notas
CMP001-USR001,USR001,PROD001,"[""PROD001""]",2024-01-15 14:30:00,1.15,TND003,1.22,1.30,1.15,0.075,1.8,0:10:00,0.80,0.07,5.74,3,"[{""tienda"":""TND003"",""orden"":1}]","TND003",1.8,0:10:00,0.80,85.50,92.00,78.00,75.00,Comprar,"Mejor precio en tienda cercana"
"@
        
        "HISTORIAL_COMPRAS.csv" = @"
CompraID,UserID,StoreID,Fecha_Compra,Total_Compra,Total_Descuentos,Total_Sin_Descuentos,N_Productos,N_Items,Lista_Productos,Método_Pago,Tipo_Compra,Ticket_Image,Ticket_PDF,Valoración_Compra,Valoración_Productos,Valoración_Atención,Valoración_Tienda,Comentarios,Problemas,Sugerencias,Fecha_Registro
CMP001-USR001,USR001,TND003,2024-01-15 16:20:00,45.60,5.40,51.00,15,18,"[{""producto"":""PROD001"",""cantidad"":2,""precio_unitario"":1.15,""total"":2.30}]",Tarjeta,Presencial,C:\Tickets\ticket001.jpg,C:\Tickets\ticket001.pdf,4.5,4.2,4.8,4.3,"Todo correcto, buen servicio","Ninguno","Mejor señalización en pasillos",2024-01-15 16:30:00
"@
        
        "PREFERENCIAS_IA.csv" = @"
PrefID,UserID,Categoría_Favorita,Subcategoría_Favorita,Marca_Favorita,Tienda_Favorita,Gasto_Promedio_Mes,Frecuencia_Compra,Día_Preferido_Compra,Hora_Preferida,Sensibilidad_Precio,Sensibilidad_Calidad,Sensibilidad_Distancia,Sensibilidad_Tiempo,Sensibilidad_Marca,Tolerancia_Desplazamiento,Presupuesto_Máx_Producto,Preferencia_Ofertas,Preferencia_Ecológico,Preferencia_Local,Historial_Recomendaciones,Acierto_Recomendaciones,Última_Actualización,Modelo_IA,Versión_Modelo
PREF001-USR001,USR001,Alimentación,Lácteos,Nestlé,TND003,200.00,4,Sábado,10:00:00,0.80,0.60,0.40,0.50,0.30,5.00,10.00,TRUE,FALSE,TRUE,"[{""fecha"":""2024-01-15"",""producto"":""PROD001"",""aceptada"":true}]",75.50,2024-01-20 10:30:00,Modelo_Colaborativo_Basico,1.0
"@
    }
    
    Write-Host "`nCreando archivos CSV:" -ForegroundColor Yellow
    
    # Crear archivos CSV
    $fileCount = 0
    foreach ($file in $csvStructures.Keys) {
        $filePath = Join-Path $csvDir $file
        Write-Host "  Creando: $file" -ForegroundColor Gray
        $csvStructures[$file] | Out-File -FilePath $filePath -Encoding UTF8 -Force
        $fileCount++
        Write-Host "    ✓ $file creado" -ForegroundColor Green
    }
    
    Write-Host "`n✓ $fileCount archivos CSV creados" -ForegroundColor Green
    
    Pause-Script -Message "Archivos CSV creados. Presiona una tecla para crear instrucciones..."
    
    # Crear archivo de instrucciones
    $instructions = @"
# SISTEMA COMPARADOR DE COMPRAS IA - ESTRUCTURA CSV
# =================================================

ESTRUCTURA DE ARCHIVOS CSV:
$(($csvStructures.Keys | ForEach-Object { "• $_" }) -join "`n")

INSTRUCCIONES PARA IMPORTAR A EXCEL:

1. ABRIR MICROSOFT EXCEL
2. PARA CADA ARCHIVO CSV:
   a. Ir a Datos → Desde archivo de texto/CSV
   b. Seleccionar el archivo CSV
   c. Configurar:
      - Origen del archivo: 65001 : Unicode (UTF-8)
      - Delimitador: Coma
      - Calificación de texto: "
   d. Hacer clic en Cargar
   e. Cambiar nombre de la hoja al nombre del archivo (sin .csv)

3. GUARDAR COMO LIBRO HABILITADO PARA MACROS:
   a. Archivo → Guardar como
   b. Tipo: Libro de Excel habilitado para macros (*.xlsm)
   c. Nombre: Comparador_Compras_IA_Completo.xlsm

4. SI EL ARCHIVO SE ABRE COMO SOLO LECTURA:
   a. Cierra el archivo
   b. Haz clic derecho sobre el archivo → Propiedades
   c. Desmarca "Solo lectura" si está marcado
   d. Haz clic en "Desbloquear" en la sección de seguridad
   e. Aplica los cambios

UBICACIóN DE ARCHIVOS: $csvDir

Fecha de creación: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")
Versión del sistema: $VERSION
"@
    
    $instructions | Out-File -FilePath (Join-Path $csvDir "INSTRUCCIONES_IMPORTACION.txt") -Encoding UTF8 -Force
    
    Write-Host "✓ Instrucciones creadas" -ForegroundColor Green
    Write-Host "`nEstructura CSV alternativa creada en: $csvDir" -ForegroundColor Cyan
}

# ===================================================
# EJECUCIÓN PRINCIPAL
# ===================================================

try {
    Write-Host "`n===================================================" -ForegroundColor Cyan
    Write-Host "  EJECUTANDO CREAR_EXCEL.PS1" -ForegroundColor Cyan
    Write-Host "===================================================" -ForegroundColor Cyan
    
    Main
    
    # Resumen final
    $END_TIME = Get-Date
    $DURATION = ($END_TIME - $START_TIME).TotalSeconds
    
    Write-Host "`n"
    Write-Host "===================================================" -ForegroundColor Green
    Write-Host "  PROCESO COMPLETADO" -ForegroundColor Green
    Write-Host "===================================================" -ForegroundColor Green
    Write-Host "`n"
    
    Write-Host "RESUMEN:" -ForegroundColor Yellow
    Write-Host "• Tiempo total: $($DURATION.ToString('0.00')) segundos" -ForegroundColor White
    Write-Host "• Errores encontrados: $GLOBAL_ERRORS" -ForegroundColor White
    
    if ($EXCEL_AVAILABLE) {
        if (Test-Path $EXCEL_FILE) {
            $size = (Get-Item $EXCEL_FILE).Length / 1MB
            Write-Host "• Archivo creado: $EXCEL_FILE" -ForegroundColor White
            Write-Host "• Tamaño del archivo: $($size.ToString('0.00')) MB" -ForegroundColor White
            
            # Verificación final
            Write-Host "`n[VERIFICACION FINAL]" -ForegroundColor Cyan
            $isReadOnly = (Get-Item $EXCEL_FILE).IsReadOnly
            if ($isReadOnly) {
                Write-Host "✗ ADVERTENCIA: El archivo aún está marcado como solo lectura" -ForegroundColor Red
                Write-Host "  Por favor, desmarca manualmente en Propiedades del archivo" -ForegroundColor Yellow
            } else {
                Write-Host "✓ El archivo está listo para editar" -ForegroundColor Green
            }
        } else {
            Write-Host "• Archivo Excel NO creado" -ForegroundColor Red
        }
    } else {
        Write-Host "• Archivos CSV creados en: $PROJECT_ROOT\CSV_Backup" -ForegroundColor White
    }
    
    Write-Host "• Registro de actividad: $LOG_FILE" -ForegroundColor White
    Write-Host "`n"
    
    if ($GLOBAL_ERRORS -eq 0) {
        Write-Host "¡Excel creado exitosamente!" -ForegroundColor Green
    } else {
        Write-Host "Proceso completado con advertencias" -ForegroundColor Yellow
    }
    
    Write-Host "`n"
    
    Pause-Script -Message "Proceso finalizado. Presiona una tecla para salir..."
    
    # Código de salida
    exit $GLOBAL_ERRORS
    
} catch {
    Write-Host "`n✗✗✗ ERROR FATAL NO CONTROLADO ✗✗✗" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Yellow
    
    Pause-Script -Message "Error fatal. Presiona una tecla para salir..." -ForcePause $true
    
    exit 99
}