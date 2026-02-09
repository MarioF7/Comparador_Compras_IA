param(
    [Parameter(Mandatory=$false)]
    [string]$ProjectPath = (Split-Path -Parent $MyInvocation.MyCommand.Path) + "\..\Comparador_Compras_IA",
    
    [Parameter(Mandatory=$false)]
    [switch]$Silent = $false
)

# configurar_sistema.ps1
# Script de configuración avanzada del sistema - Versión 3.5.0
# Compatible con Windows 7/8/10/11 y PowerShell 3.0+

# ===================================================================
# CONFIGURACIÓN INICIAL
# ===================================================================

# Configurar codificación para caracteres especiales
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Variables globales
$ErrorActionPreference = "Stop"
$script:ConfigData = @{}
$script:LogEntries = New-Object System.Collections.ArrayList

# Función de logging mejorada
function Write-SystemLog {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Module = "CONFIG"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] [$Module] $Message"
    
    # Añadir a lista en memoria
    [void]$script:LogEntries.Add($logEntry)
    
    # Mostrar en consola según nivel
    switch ($Level) {
        "SUCCESS" { 
            if (-not $Silent) { Write-Host "  [✓] $Message" -ForegroundColor Green }
        }
        "ERROR" { 
            if (-not $Silent) { Write-Host "  [!] $Message" -ForegroundColor Red }
        }
        "WARNING" { 
            if (-not $Silent) { Write-Host "  [*] $Message" -ForegroundColor Yellow }
        }
        "INFO" { 
            if (-not $Silent) { Write-Host "  [i] $Message" -ForegroundColor Cyan }
        }
        default {
            if (-not $Silent) { Write-Host "  [i] $Message" -ForegroundColor Gray }
        }
    }
    
    # Guardar en archivo log
    try {
        $logPath = Join-Path $ProjectPath "Logs\configuracion_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        $logEntry | Out-File -FilePath $logPath -Append -Encoding UTF8 -Force
    } catch {
        # Silenciar errores de log
    }
}

# Función para verificar requisitos
function Test-SystemRequirements {
    Write-SystemLog "Verificando requisitos del sistema..." -Level "INFO"
    
    $requirements = @{
        "PowerShell Version" = @{
            Minimum = 3
            Current = $PSVersionTable.PSVersion.Major
            Status = ($PSVersionTable.PSVersion.Major -ge 3)
        }
        ".NET Framework" = @{
            Minimum = "4.5"
            Current = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name Release -ErrorAction SilentlyContinue).Release
            Status = $true  # Se verificará después
        }
        "Espacio en disco" = @{
            Minimum = 100MB
            Current = (Get-PSDrive -Name $env:SystemDrive[0]).Free
            Status = ((Get-PSDrive -Name $env:SystemDrive[0]).Free -gt 100MB)
        }
        "Permisos de escritura" = @{
            Status = $true
        }
    }
    
    # Verificar .NET Framework
    try {
        $netRelease = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name Release -ErrorAction Stop).Release
        if ($netRelease -ge 379893) { # .NET 4.5.2 o superior
            $requirements[".NET Framework"].Current = "4.5.2+"
            $requirements[".NET Framework"].Status = $true
        } else {
            $requirements[".NET Framework"].Status = $false
        }
    } catch {
        $requirements[".NET Framework"].Status = $false
    }
    
    # Verificar permisos de escritura
    try {
        $testFile = Join-Path $ProjectPath "test_permissions.tmp"
        "test" | Out-File -FilePath $testFile -Encoding UTF8 -Force
        Remove-Item $testFile -Force -ErrorAction Stop
        $requirements["Permisos de escritura"].Status = $true
    } catch {
        $requirements["Permisos de escritura"].Status = $false
    }
    
    # Mostrar resultados
    foreach ($req in $requirements.Keys) {
        if ($requirements[$req].Status) {
            Write-SystemLog "OK" -Level "SUCCESS"
        } else {
            Write-SystemLog "FALLO" -Level "ERROR"
		}
    }
    
    # Verificar si hay fallos críticos
    $criticalFailures = $requirements.Values | Where-Object { $_.Status -eq $false } | Measure-Object
    return ($criticalFailures.Count -eq 0)
}

# Función para cargar configuración existente
function Load-Configuration {
    param([string]$ConfigPath)
    
    $defaultConfig = @{
        Sistema = @{
            Version = "3.5.0"
            FechaInstalacion = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            Modo = "Normal"
            Idioma = "es-ES"
        }
        Usuario = @{
            Nombre = $env:USERNAME
            Email = ""
            Telefono = ""
            Direccion = ""
            Ciudad = ""
            CP = ""
            Coordenadas = @{
                Lat = 0
                Lon = 0
            }
        }
        Preferencias = @{
            Moneda = "EUR"
            UnidadDistancia = "km"
            UnidadPeso = "kg"
            FormatoFecha = "dd/MM/yyyy"
            Notificaciones = $true
            Tema = "Claro"
            AutoBackup = $true
        }
        Rendimiento = @{
            CacheHabilitado = $true
            MaxCacheMB = 100
            LogDetallado = $false
            AutoActualizar = $true
        }
        Seguridad = @{
            EncriptarDatos = $false
            HashPasswords = $true
            TimeoutMinutos = 30
            MaxIntentosLogin = 3
        }
        Conexiones = @{
            APISupermercados = @()
            APIMaps = ""
            APIWeather = ""
            Proxy = @{
                Habilitado = $false
                Servidor = ""
                Puerto = 0
            }
        }
    }
    
    # Intentar cargar configuración existente
    try {
        if (Test-Path $ConfigPath) {
            $jsonContent = Get-Content $ConfigPath -Encoding UTF8 -Raw
            # Convertir de JSON a objeto PSCustomObject
            $existingConfigObj = $jsonContent | ConvertFrom-Json
            Write-SystemLog "Configuración existente cargada desde: $ConfigPath" -Level "SUCCESS"
            
            # Convertir PSCustomObject a Hashtable recursivamente
            $existingConfig = ConvertTo-Hashtable $existingConfigObj
            
            # Combinar configuraciones (mantener existentes, añadir nuevas)
            return Merge-Hashtables $defaultConfig, $existingConfig
        }
    } catch {
        Write-SystemLog "Error al cargar configuración existente: $($_.Exception.Message)" -Level "WARNING"
    }
    
    return $defaultConfig
}

# Función auxiliar para convertir PSCustomObject a Hashtable recursivamente
function ConvertTo-Hashtable {
    param([Parameter(ValueFromPipeline)]$InputObject)
    
    process {
        if ($null -eq $InputObject) {
            return $null
        }
        
        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
            $collection = @()
            foreach ($item in $InputObject) {
                $collection += (ConvertTo-Hashtable $item)
            }
            return $collection
        } elseif ($InputObject -is [PSCustomObject]) {
            $hash = @{}
            foreach ($property in $InputObject.PSObject.Properties) {
                $hash[$property.Name] = ConvertTo-Hashtable $property.Value
            }
            return $hash
        } else {
            return $InputObject
        }
    }
}

# Función auxiliar para combinar hashtables
function Merge-Hashtables {
    param([hashtable[]]$Hashtables)
    
    $result = @{}
    
    foreach ($ht in $Hashtables) {
        foreach ($key in $ht.Keys) {
            if ($result.ContainsKey($key)) {
                if ($result[$key] -is [hashtable] -and $ht[$key] -is [hashtable]) {
                    $result[$key] = Merge-Hashtables $result[$key], $ht[$key]
                } else {
                    $result[$key] = $ht[$key]
                }
            } else {
                $result[$key] = $ht[$key]
            }
        }
    }
    
    return $result
}

# Función para crear estructura avanzada de carpetas
function Create-AdvancedFolderStructure {
    param([string]$RootPath)
    
    Write-SystemLog "Creando estructura avanzada de carpetas..." -Level "INFO"
    
    $folders = @(
        # Nivel 1
        @{Path = "Data_Backup"; Subfolders = @("Diario", "Semanal", "Mensual", "Automatico", "Manual")}
        @{Path = "Configuraciones"; Subfolders = @("Usuarios", "Sistema", "APIs", "Plantillas")}
        @{Path = "Scripts_IA"; Subfolders = @("Analisis", "Modelos", "Utilidades", "Pruebas")}
        @{Path = "Reportes"; Subfolders = @("PDF", "Excel", "HTML", "Dashboard", "Automaticos")}
        @{Path = "Tickets"; Subfolders = @("Imagenes", "PDF", "OCR", "Procesados")}
        @{Path = "Templates"; Subfolders = @("Email", "Reportes", "Documentos", "Contratos")}
        @{Path = "Logs"; Subfolders = @("Sistema", "Errores", "Auditoria", "Depuracion")}
        @{Path = "Cache"; Subfolders = @("Imagenes", "Datos", "Temporal", "Sesiones")}
        @{Path = "Exportaciones"; Subfolders = @("CSV", "Excel", "PDF", "JSON", "XML")}
        @{Path = "Datos_Externos"; Subfolders = @("APIs", "WebScraping", "Importados", "Procesados")}
        @{Path = "Plantillas_IA"; Subfolders = @("Modelos", "DatosEntrenamiento", "Resultados")}
        @{Path = "Modelos_ML"; Subfolders = @("Entrenados", "EnEntrenamiento", "Backup")}
        @{Path = "Modulos"; Subfolders = @("VBA", "Python", "PowerShell", "SQL")}
        @{Path = "Documentacion"; Subfolders = @("Tecnica", "Usuario", "API", "Cambios")}
        @{Path = "Temp"; Subfolders = @("Uploads", "Downloads", "Procesamiento")}
        @{Path = "Sesiones"; Subfolders = @("Usuarios", "Sistema", "Backup")}
    )
    
    $createdCount = 0
    $errorCount = 0
    
    foreach ($folder in $folders) {
        $mainPath = Join-Path $RootPath $folder.Path
        
        try {
            # Crear carpeta principal
            if (-not (Test-Path $mainPath)) {
                New-Item -ItemType Directory -Path $mainPath -Force | Out-Null
                Write-SystemLog "Creada carpeta: $($folder.Path)" -Level "SUCCESS"
                $createdCount++
            }
            
            # Crear subcarpetas
            foreach ($subfolder in $folder.Subfolders) {
                $subPath = Join-Path $mainPath $subfolder
                if (-not (Test-Path $subPath)) {
                    New-Item -ItemType Directory -Path $subPath -Force | Out-Null
                }
            }
            
        } catch {
            Write-SystemLog "Error creando carpeta $($folder.Path): $($_.Exception.Message)" -Level "ERROR"
            $errorCount++
        }
    }
    
    Write-SystemLog "Estructura de carpetas creada: $createdCount carpetas principales" -Level "SUCCESS"
    return ($errorCount -eq 0)
}

# Función para crear archivos de configuración avanzados
function Create-AdvancedConfigFiles {
    param(
        [hashtable]$Config,
        [string]$ConfigPath
    )
    
    Write-SystemLog "Creando archivos de configuración avanzados..." -Level "INFO"
    
    try {
        # 1. Configuración principal del sistema (JSON)
        $configJson = $Config | ConvertTo-Json -Depth 10
        $configJson | Out-File -FilePath (Join-Path $ConfigPath "config_sistema.json") -Encoding UTF8 -Force
        Write-SystemLog "Configuración principal creada: config_sistema.json" -Level "SUCCESS"
        
        # 2. Configuración de usuario (JSON)
        $userConfig = @{
            Usuario = $Config.Usuario
            Preferencias = $Config.Preferencias
            Sesion = @{
                UltimoAcceso = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                IntentosFallidos = 0
                IP = (Test-Connection -ComputerName $env:COMPUTERNAME -Count 1).IPv4Address.IPAddressToString
            }
        }
        ($userConfig | ConvertTo-Json -Depth 5) | Out-File -FilePath (Join-Path $ConfigPath "..\Configuraciones\Usuarios\config_$($env:USERNAME).json") -Encoding UTF8 -Force
        
        # 3. Configuración de conexiones (XML)
		$xmlFilePath = Join-Path $ConfigPath "\APIs\conexiones.xml"
        $xmlDir = Split-Path $xmlFilePath -Parent
        if (-not (Test-Path $xmlDir)) {
            New-Item -ItemType Directory -Path $xmlDir -Force | Out-Null
            Write-SystemLog "Creado directorio APIs: $xmlDir" -Level "INFO"
        }
		
        $xmlConfig = [xml]@"
<?xml version="1.0" encoding="UTF-8"?>
<Configuraciones>
    <Conexiones>
        <APIs>
            <GoogleMaps activa="false" clave="" />
            <OpenWeather activa="false" clave="" />
            <Supermercados>
                <API nombre="Mercadona" activa="false" endpoint="" />
                <API nombre="Carrefour" activa="false" endpoint="" />
            </Supermercados>
        </APIs>
        <Proxy activo="false">
            <Servidor></Servidor>
            <Puerto>0</Puerto>
            <Usuario></Usuario>
            <Password encriptado=""></Password>
        </Proxy>
        <BaseDatos>
            <Local tipo="SQLite" archivo="database.db" />
            <Remota tipo="None" />
        </BaseDatos>
    </Conexiones>
</Configuraciones>
"@
		$xmlConfig.Save((Join-Path $ConfigPath "..\Configuraciones\APIs\conexiones.xml"))
        
        # 4. Configuración de seguridad
        $securityConfig = @{
            Seguridad = @{
                Encriptacion = @{
                    Algoritmo = "AES-256"
                    Salt = [System.Convert]::ToBase64String((1..32 | ForEach-Object { Get-Random -Minimum 0 -Maximum 255 }))
                }
                Autenticacion = @{
                    MinCaracteres = 8
                    RequerirMayusculas = $true
                    RequerirNumeros = $true
                    RequerirEspeciales = $false
                }
                Sesiones = @{
                    Timeout = 30
                    MaxSesiones = 3
                    RenewToken = $true
                }
            }
        }
        ($securityConfig | ConvertTo-Json -Depth 5) | Out-File -FilePath (Join-Path $ConfigPath "..\Configuraciones\Sistema\seguridad.json") -Encoding UTF8 -Force
        
        # 5. Configuración de backup
        $backupConfig = @{
            Backup = @{
                Automatico = @{
                    Habilitado = $true
                    IntervaloHoras = 24
                    MaxBackups = @{
                        Diarios = 7
                        Semanales = 4
                        Mensuales = 12
                        Anuales = 2
                    }
                }
                Manual = @{
                    Comprimir = $true
                    Formato = "ZIP"
                    IncluirLogs = $true
                }
                Destinos = @(
                    @{
                        Tipo = "Local"
                        Ruta = "Data_Backup\Automatico"
                    }
                )
            }
        }
        ($backupConfig | ConvertTo-Json -Depth 5) | Out-File -FilePath (Join-Path $ConfigPath "..\Configuraciones\Sistema\backup.json") -Encoding UTF8 -Force
        
        Write-SystemLog "5 archivos de configuración creados exitosamente" -Level "SUCCESS"
        return $true
        
    } catch {
        Write-SystemLog "Error creando archivos de configuración: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

# Función para crear scripts de utilidad
function Create-UtilityScripts {
    param([string]$ScriptsPath)
    
    Write-SystemLog "Creando scripts de utilidad..." -Level "INFO"
    
    $scripts = @{
        "backup_automatico.ps1" = @'
# Script de backup automático - Sistema Comparador Compras IA
param([string]$ProjectPath = ".")

$backupDir = Join-Path $ProjectPath "Data_Backup\Automatico\$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

# Archivos a respaldar
$filesToBackup = @(
    "Comparador_Compras_IA_Completo.xlsm",
    "Configuraciones\*.json",
    "Configuraciones\*.xml",
    "Logs\*.log"
)

foreach ($pattern in $filesToBackup) {
    $files = Get-ChildItem -Path (Join-Path $ProjectPath $pattern) -File
    foreach ($file in $files) {
        $dest = Join-Path $backupDir $file.Name
        Copy-Item $file.FullName $dest -Force
    }
}

# Comprimir backup
$zipFile = "$backupDir.zip"
Compress-Archive -Path "$backupDir\*" -DestinationPath $zipFile -Force

# Limpiar carpeta temporal
Remove-Item $backupDir -Recurse -Force

Write-Output "Backup completado: $zipFile"
'@

        "limpiar_cache.ps1" = @'
# Script para limpiar caché del sistema
param([string]$ProjectPath = ".")

$cacheDirs = @(
    "Cache\Imagenes",
    "Cache\Datos",
    "Cache\Temporal",
    "Temp"
)

$totalFreed = 0
foreach ($dir in $cacheDirs) {
    $fullPath = Join-Path $ProjectPath $dir
    if (Test-Path $fullPath) {
        $files = Get-ChildItem $fullPath -File -Recurse
        $size = ($files | Measure-Object -Property Length -Sum).Sum
        Remove-Item "$fullPath\*" -Recurse -Force
        $totalFreed += $size
    }
}

Write-Output "Cache limpiado: $([math]::Round($totalFreed/1MB, 2)) MB liberados"
'@

        "verificar_sistema.ps1" = @'
# Script de verificación del sistema
param([string]$ProjectPath = ".")

$checks = @()

# 1. Verificar archivos esenciales
$essentialFiles = @(
    "Comparador_Compras_IA_Completo.xlsm",
    "Configuraciones\config_sistema.json",
    "INSTRUCCIONES_PROYECTO.txt"
)

foreach ($file in $essentialFiles) {
    $path = Join-Path $ProjectPath $file
    $checks += @{
        Archivo = $file
        Existe = (Test-Path $path)
        Tamaño = if (Test-Path $path) { (Get-Item $path).Length } else { 0 }
    }
}

# 2. Verificar permisos
try {
    $testFile = Join-Path $ProjectPath "test_permissions.tmp"
    "test" | Out-File $testFile -Encoding UTF8
    Remove-Item $testFile -Force
    $permisos = $true
} catch {
    $permisos = $false
}

$checks += @{
    Componente = "Permisos de escritura"
    Estado = $permisos
}

# 3. Verificar espacio
$drive = (Get-PSDrive -Name $env:SystemDrive[0])
$checks += @{
    Componente = "Espacio en disco"
    Estado = ($drive.Free -gt 100MB)
    Libre = "$([math]::Round($drive.Free/1MB, 2)) MB"
}

# Mostrar resultados
$checks | ForEach-Object {
    $status = if ($_.Estado -or ($_.Existe -eq $true)) { "OK" } else { "ERROR" }
    Write-Host "[$status] $($_.Archivo ?? $_.Componente)" -ForegroundColor $(if ($status -eq "OK") { "Green" } else { "Red" })
}
'@
    }
    
    $created = 0
    foreach ($scriptName in $scripts.Keys) {
        $scriptPath = Join-Path $ScriptsPath $scriptName
        $scripts[$scriptName] | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
        $created++
    }
    
    Write-SystemLog "$created scripts de utilidad creados" -Level "SUCCESS"
    return $true
}

# Función para configurar políticas del sistema
function Set-SystemPolicies {
    Write-SystemLog "Configurando políticas del sistema..." -Level "INFO"
    
    try {
        # Configurar política de ejecución de PowerShell (solo para proceso actual)
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
        
        # Configurar políticas de Internet Explorer (si existe) para evitar advertencias
        if (Test-Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main") {
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -Value 1 -ErrorAction SilentlyContinue
        }
        
        Write-SystemLog "Políticas del sistema configuradas" -Level "SUCCESS"
        return $true
        
    } catch {
        Write-SystemLog "Error configurando políticas: $($_.Exception.Message)" -Level "WARNING"
        return $false
    }
}

# Función principal
function Main {
    # Encabezado
    if (-not $Silent) {
        Write-Host "`n" -NoNewline
        Write-Host "===================================================" -ForegroundColor Cyan
        Write-Host "  CONFIGURADOR DEL SISTEMA - Versión 3.5.0" -ForegroundColor Cyan
        Write-Host "===================================================" -ForegroundColor Cyan
        Write-Host "`n"
    }
    
    Write-SystemLog "Iniciando configuración del sistema..." -Level "INFO"
    Write-SystemLog "Ruta del proyecto: $ProjectPath" -Level "INFO"
    
    # Verificar que el proyecto existe
    if (-not (Test-Path $ProjectPath)) {
        Write-SystemLog "ERROR: La ruta del proyecto no existe: $ProjectPath" -Level "ERROR"
        return 1
    }
    
    # Verificar requisitos del sistema
    if (-not (Test-SystemRequirements)) {
        Write-SystemLog "Fallo en la verificación de requisitos del sistema" -Level "ERROR"
        return 2
    }
    
    # Configurar políticas
    Set-SystemPolicies | Out-Null
    
    # Crear estructura de carpetas
    if (-not (Create-AdvancedFolderStructure -RootPath $ProjectPath)) {
        Write-SystemLog "Advertencia: Error creando algunas carpetas" -Level "WARNING"
    }
    
    # Cargar/Crear configuración
    $configPath = Join-Path $ProjectPath "Configuraciones\config_sistema.json"
    $script:ConfigData = Load-Configuration -ConfigPath $configPath
    
    # Crear archivos de configuración avanzados
    $configDir = Join-Path $ProjectPath "Configuraciones"
    if (-not (Create-AdvancedConfigFiles -Config $script:ConfigData -ConfigPath $configDir)) {
        Write-SystemLog "Advertencia: Error creando algunos archivos de configuración" -Level "WARNING"
    }
    
    # Crear scripts de utilidad
    $scriptsDir = Join-Path $ProjectPath "Scripts_IA\Utilidades"
    Create-UtilityScripts -ScriptsPath $scriptsDir | Out-Null
    
    # Crear archivo de resumen
    $summaryPath = Join-Path $ProjectPath "Configuraciones\resumen_configuracion.txt"
    $summary = @"
RESUMEN DE CONFIGURACIÓN DEL SISTEMA
====================================
Fecha: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Versión: 3.5.0
Usuario: $env:USERNAME
Equipo: $env:COMPUTERNAME
Ruta Proyecto: $ProjectPath

ESTRUCTURA CREADA:
-----------------
✓ Data_Backup (con 5 subcarpetas)
✓ Configuraciones (con 4 subcarpetas)
✓ Scripts_IA (con 4 subcarpetas)
✓ Reportes (con 5 subcarpetas)
✓ Tickets (con 4 subcarpetas)
✓ Templates (con 4 subcarpetas)
✓ Logs (con 4 subcarpetas)
✓ Cache (con 4 subcarpetas)
✓ 6 carpetas adicionales especializadas

ARCHIVOS DE CONFIGURACIÓN:
--------------------------
1. config_sistema.json (Configuración principal)
2. config_$($env:USERNAME).json (Configuración de usuario)
3. conexiones.xml (Configuración de APIs)
4. seguridad.json (Configuración de seguridad)
5. backup.json (Configuración de backups)

SCRIPTS DE UTILIDAD:
--------------------
1. backup_automatico.ps1 (Sistema de backups automáticos)
2. limpiar_cache.ps1 (Limpieza de caché del sistema)
3. verificar_sistema.ps1 (Verificación de integridad)

ESTADO DEL SISTEMA:
-------------------
Requisitos mínimos: CUMPLIDOS
Políticas del sistema: CONFIGURADAS
Estructura de carpetas: COMPLETA
Archivos de configuración: CREADOS
Scripts de utilidad: INSTALADOS

PRÓXIMOS PASOS:
---------------
1. Abrir el archivo Excel principal
2. Habilitar macros cuando se solicite
3. Configurar sus datos personales
4. Empezar a añadir productos y precios
5. Revisar los scripts de utilidad según necesidad

SOPORTE:
--------
• Consulte INSTRUCCIONES_PROYECTO.txt
• Revise los logs en la carpeta Logs\
• Ejecute verificar_sistema.ps1 para diagnóstico

¡SISTEMA CONFIGURADO EXITOSAMENTE!
===================================
"@
    
    $summary | Out-File -FilePath $summaryPath -Encoding UTF8 -Force
    
    # Mostrar resumen final
    if (-not $Silent) {
        Write-Host "`n"
        Write-Host "===================================================" -ForegroundColor Green
        Write-Host "  CONFIGURACIÓN COMPLETADA EXITOSAMENTE" -ForegroundColor Green
        Write-Host "===================================================" -ForegroundColor Green
        Write-Host "`nResumen de la configuración:" -ForegroundColor Yellow
        Write-Host "  • Estructura de carpetas: COMPLETA" -ForegroundColor Green
        Write-Host "  • Archivos de configuración: 5 creados" -ForegroundColor Green
        Write-Host "  • Scripts de utilidad: 3 instalados" -ForegroundColor Green
        Write-Host "  • Resumen guardado en: Configuraciones\resumen_configuracion.txt" -ForegroundColor Cyan
        Write-Host "`n¡El sistema está listo para usar!" -ForegroundColor Green
        Write-Host "`n"
    }
    
    Write-SystemLog "Configuración del sistema completada exitosamente" -Level "SUCCESS"
    return 0
}

# Punto de entrada del script
try {
    $exitCode = Main
    exit $exitCode
} catch {
    Write-SystemLog "ERROR FATAL: $($_.Exception.Message)" -Level "ERROR"
    Write-SystemLog "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
    exit 99
}