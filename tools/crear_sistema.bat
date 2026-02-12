```batch
@echo off
chcp 65001 >nul
title [INSTALADOR] Sistema Comparador de Compras IA v3.5
setlocal enabledelayedexpansion

:: ======================= CONFIGURACIÓN PRINCIPAL =======================
set "VER_SISTEMA=3.5.0"
set "NOMBRE_SISTEMA=Sistema Comparador de Compras Inteligente IA"
set "AUTOR=MarioF7"

:: Directorios PRINCIPALES (NO modificar rutas internas aquí)
set "SCRIPT_DIR=%~dp0"
set "PROJECT_ROOT=%SCRIPT_DIR%..\Comparador_Compras_IA"

:: Configuración de logs
set "LOG_DIR=%PROJECT_ROOT%\Logs"
set "LOG_FILE=%LOG_DIR%\instalacion_%date:~-4,4%%date:~-7,2%%date:~-10,2%_%time:~0,2%%time:~3,2%.log"

:: Variables de estado del sistema
set /a ERROR_FLAG=0
set /a WARNING_FLAG=0
set /a ADMIN_MODE=0
set /a PHASE=0
set "EXCEL_INSTALLED=0"
set "POWERSHELL_VERSION=0"
set "NET_VERSION=0"

:: ======================= INICIO DEL PROGRAMA =======================
echo.
echo ===================================================
echo    %NOMBRE_SISTEMA%
echo    Version: %VER_SISTEMA% - Release Estable
echo ===================================================
echo.
echo Fecha  : %date% %time%
echo Usuario: %USERNAME%
echo Equipo : %COMPUTERNAME%
echo ===================================================
echo.

:: ===================================================================
:: CREAR ESTRUCTURA DE LOGS MEJORADA
:: ===================================================================
if not exist "%PROJECT_ROOT%\Logs" (
    mkdir "%PROJECT_ROOT%\Logs" 2>nul
    if errorlevel 1 (
        echo [ERROR] No se pudo crear carpeta Logs
        set /a ERROR_FLAG+=1
    )
)

:: ===================================================================
:: PROGRAMA PRINCIPAL (Flujo original mejorado)
:: ===================================================================

:: FASE 1: VERIFICACIÓN DEL SISTEMA MEJORADA
echo.
echo [PROGRESO] FASE 1: Verificación del sistema operativo y requisitos...
echo.

echo ===================================================
echo INICIANDO INSTALACIÓN - Versión %SCRIPT_VERSION%
echo Fecha: %FECHA_INSTALACION%
echo Usuario: %USERNAME%
echo Sistema: %COMPUTERNAME%
echo Directorio de scripts: %SCRIPT_DIR%
echo ===================================================

:: Verificar sistema operativo (compatible con todas las versiones)
echo Verificando sistema operativo...
ver | findstr /r /c:"Microsoft Windows" >nul
if %errorlevel% neq 0 (
    ver | findstr /r /c:"Windows" >nul
    if %errorlevel% neq 0 (
        echo [ERROR CRÍTICO] Sistema operativo no compatible.
        echo [ERROR] Se requiere Windows 7, 8, 10 u 11.
        set /a ERROR_FLAG+=3
    ) else (
        echo [OK] Sistema operativo compatible (Windows detectado)
    )
) else (
    echo [OK] Sistema operativo compatible (Microsoft Windows detectado)
)

:: Verificar arquitectura del sistema
echo Verificando arquitectura del sistema...
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    echo [OK] Sistema de 64 bits detectado
    set "ARCH=64"
) else if "%PROCESSOR_ARCHITECTURE%"=="x86" (
    echo [OK] Sistema de 32 bits detectado
    set "ARCH=32"
) else if "%PROCESSOR_ARCHITECTURE%"=="ARM64" (
    echo [OK] Sistema ARM64 detectado
    set "ARCH=ARM64"
) else (
    echo [ADVERTENCIA] Arquitectura no estándar: %PROCESSOR_ARCHITECTURE%
    set "ARCH=DESCONOCIDA"
    set /a WARNING_FLAG+=1
)

:: Verificar permisos de administrador (método mejorado)
echo Verificando permisos de administrador...
net session >nul 2>&1
if %errorlevel% equ 0 (
    set "ADMIN_MODE=1"
    echo [OK] Ejecutando con permisos de administrador
) else (
    echo [ADVERTENCIA] Ejecutando sin permisos de administrador
    echo [ADVERTENCIA]   Algunas funciones pueden estar limitadas
    set /a WARNING_FLAG+=1
)

:: Verificar PowerShell (método mejorado y robusto)
echo Verificando PowerShell...
where powershell >nul 2>&1
if %errorlevel% equ 0 (
    powershell -Command "Write-Output $PSVersionTable.PSVersion.Major" > "%TEMP%\psver.txt" 2>&1
    set /p POWERSHELL_VERSION= < "%TEMP%\psver.txt" 2>nul
    del "%TEMP%\psver.txt" 2>nul
    
    if "!POWERSHELL_VERSION!"=="" (
        echo [ADVERTENCIA] PowerShell detectado pero no se pudo obtener versión
        set "POWERSHELL_VERSION=Desconocida"
        set /a WARNING_FLAG+=1
    ) else (
        echo [OK] PowerShell !POWERSHELL_VERSION! detectado
    )
) else (
    echo [ERROR CRÍTICO] PowerShell no encontrado
    echo [ERROR] PowerShell es requerido para el funcionamiento del sistema.
    set /a ERROR_FLAG+=3
)

:: ===================================================================
:: VERIFICACIÓN DE .NET FRAMEWORK - CORREGIDO Y FUNCIONAL
:: ===================================================================
REM echo Verificando .NET Framework...
REM echo [DEBUG 24] Iniciando verificacion de .NET Framework...
REM echo Verificando .NET Framework...

REM :: PRIMER INTENTO: Verificar .NET 4.0 o superior usando un metodo robusto
REM echo [DEBUG 25] Intentando metodo robusto de verificacion de .NET...

REM :: Metodo 1: Verificar usando WMIC (funciona en todas las versiones)
REM echo [DEBUG 25.1] Probando WMIC...
REM wmic product where "name like '%%Microsoft .NET%%'" get name, version 2>nul | findstr /i ".NET" >nul
REM if %errorlevel% equ 0 (
    REM echo [DEBUG 25.2] .NET encontrado via WMIC
    REM for /f "tokens=2 delims==" %%i in ('wmic product where "name like '%%Microsoft .NET%%'" get version /value 2^>nul ^| findstr "="') do (
        REM set "NET_DETECTED=%%i"
    REM )
    REM echo [OK] .NET Framework !NET_DETECTED! detectado via WMIC
    REM set "NET_VERSION=!NET_DETECTED!"
    REM goto :NET_VERIFIED
REM )

REM :: Metodo 2: Verificar en el registro con manejo de errores robusto
REM echo [DEBUG 25.3] WMIC no funciono, probando registro...

REM :: Crear un archivo temporal para capturar la salida
REM reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release 2>"%TEMP%\net_reg_error.txt" >"%TEMP%\net_reg_output.txt"
REM set "REG_ERROR_CODE=%errorlevel%"
REM echo [DEBUG 26] Codigo de error de reg query: %REG_ERROR_CODE%

REM :: Mostrar lo que se capturo para depuracion
REM echo [DEBUG 26.1] Contenido del archivo de error:
REM type "%TEMP%\net_reg_error.txt" 2>nul
REM echo [DEBUG 26.2] Contenido del archivo de salida:
REM type "%TEMP%\net_reg_output.txt" 2>nul

REM if %REG_ERROR_CODE% equ 0 (
    REM echo [DEBUG 27] .NET 4.0+ encontrado en registro, procesando...
    
    REM :: Leer el valor del archivo de salida
    REM set "NET_RELEASE="
    REM for /f "tokens=3" %%a in ('type "%TEMP%\net_reg_output.txt" 2^>nul') do (
        REM set "NET_RELEASE=%%a"
    REM )
    
    REM echo [DEBUG 28] Valor NET_RELEASE leido: !NET_RELEASE!
    
    REM if "!NET_RELEASE!"=="" (
        REM echo [ERROR] No fue posible obtener valor Release
        REM set /a ERROR_FLAG+=1
        REM echo [DEBUG 29] NET_RELEASE vacio, ERROR_FLAG=!ERROR_FLAG!
    REM ) else (
        REM echo [DEBUG 30] Comparando version de .NET...
        REM if !NET_RELEASE! GEQ 528040 (
            REM echo [OK] .NET Framework 4.8 o superior detectado
            REM set "NET_VERSION=4.8+"
        REM ) else if !NET_RELEASE! GEQ 461808 (
            REM echo [OK] .NET Framework 4.7.2 detectado
            REM set "NET_VERSION=4.7.2"
        REM ) else if !NET_RELEASE! GEQ 461308 (
            REM echo [OK] .NET Framework 4.7.1 detectado
            REM set "NET_VERSION=4.7.1"
        REM ) else if !NET_RELEASE! GEQ 460798 (
            REM echo [OK] .NET Framework 4.7 detectado
            REM set "NET_VERSION=4.7"
        REM ) else if !NET_RELEASE! GEQ 394802 (
            REM echo [OK] .NET Framework 4.6.2 detectado
            REM set "NET_VERSION=4.6.2"
        REM ) else if !NET_RELEASE! GEQ 394254 (
            REM echo [DEBUG 31] .NET 4.6.1 detectado
            REM echo [OK] .NET Framework 4.6.1 detectado
            REM set "NET_VERSION=4.6.1"
        REM ) else if !NET_RELEASE! GEQ 393295 (
            REM echo [OK] .NET Framework 4.6 detectado
            REM set "NET_VERSION=4.6"
        REM ) else if !NET_RELEASE! GEQ 379893 (
            REM echo [OK] .NET Framework 4.5.2 detectado
            REM set "NET_VERSION=4.5.2"
        REM ) else if !NET_RELEASE! GEQ 378675 (
            REM echo [OK] .NET Framework 4.5.1 detectado
            REM set "NET_VERSION=4.5.1"
        REM ) else if !NET_RELEASE! GEQ 378389 (
            REM echo [OK] .NET Framework 4.5 detectado
            REM set "NET_VERSION=4.5"
        REM ) else (
            REM echo [OK] .NET Framework 4.0 detectado
            REM set "NET_VERSION=4.0"
        REM )
        REM echo [DEBUG 32] NET_VERSION establecida a: !NET_VERSION!
        REM goto :NET_VERIFIED
    REM )
REM )

REM :: Metodo 3: Verificar versiones anteriores de .NET
REM echo [DEBUG 33] .NET 4.0+ no encontrado, verificando versiones anteriores...

REM :: Verificar .NET 3.5
REM reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5" /v Version 2>"%TEMP%\net35_error.txt" >"%TEMP%\net35_output.txt"
REM if %errorlevel% equ 0 (
    REM echo [ADVERTENCIA] .NET Framework 3.5 detectado (4.0+ recomendado)
    REM set "NET_VERSION=3.5"
    REM set /a WARNING_FLAG+=1
    REM echo [DEBUG 34] .NET 3.5 encontrado, WARNING_FLAG=!WARNING_FLAG!
    REM goto :NET_VERIFIED
REM )

REM :: Metodo 4: Verificar existencia fisica de archivos .NET
REM echo [DEBUG 35] Verificando archivos fisicos de .NET...
REM if exist "%windir%\Microsoft.NET\Framework64\v4.0.30319\System.dll" (
    REM echo [OK] .NET Framework 4.0+ detectado (via System.dll 64-bit)
    REM set "NET_VERSION=4.0+"
    REM goto :NET_VERIFIED
REM ) else if exist "%windir%\Microsoft.NET\Framework\v4.0.30319\System.dll" (
    REM echo [OK] .NET Framework 4.0+ detectado (via System.dll 32-bit)
    REM set "NET_VERSION=4.0+"
    REM goto :NET_VERIFIED
REM ) else if exist "%windir%\Microsoft.NET\Framework\v3.5\System.dll" (
    REM echo [ADVERTENCIA] .NET Framework 3.5 detectado (via System.dll)
    REM set "NET_VERSION=3.5"
    REM set /a WARNING_FLAG+=1
    REM goto :NET_VERIFIED
REM )

REM :: Metodo 5: Ultimo intento - verificar en carpetas
REM echo [DEBUG 36] Verificando carpetas de .NET...
REM dir "%windir%\Microsoft.NET\Framework\v4.0*" >nul 2>&1
REM if %errorlevel% equ 0 (
    REM echo [OK] .NET Framework 4.x detectado (via carpeta)
    REM set "NET_VERSION=4.x"
    REM goto :NET_VERIFIED
REM )

REM dir "%windir%\Microsoft.NET\Framework\v3.5*" >nul 2>&1
REM if %errorlevel% equ 0 (
    REM echo [ADVERTENCIA] .NET Framework 3.5 detectado (via carpeta)
    REM set "NET_VERSION=3.5"
    REM set /a WARNING_FLAG+=1
    REM goto :NET_VERIFIED
REM )

REM :: Si llegamos aqui, .NET no esta instalado
REM echo [ERROR] .NET Framework no detectado
REM echo [ERROR]   Algunas funciones avanzadas no estaran disponibles
REM set "NET_VERSION=No detectado"
REM set /a ERROR_FLAG+=1
REM echo [DEBUG 37] .NET no detectado, ERROR_FLAG=!ERROR_FLAG!

REM :NET_VERIFIED
REM :: Limpiar archivos temporales
REM del "%TEMP%\net_reg_error.txt" 2>nul
REM del "%TEMP%\net_reg_output.txt" 2>nul
REM del "%TEMP%\net35_error.txt" 2>nul
REM del "%TEMP%\net35_output.txt" 2>nul

REM echo [DEBUG 38] Verificacion de .NET completada. NET_VERSION=!NET_VERSION!


REM pause
REM :: PRIMER MÉTODO: Verificar .NET 4.0 o superior en el registro
REM reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release 2>nul
REM if %errorlevel% equ 0 (
REM pause
    REM for /f "tokens=2 delims=    " %%a in ('reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release 2^>nul') do (
        REM set "NET_RELEASE=%%a"
    REM )

    REM if "!NET_RELEASE!"=="" (
        REM echo [ERROR] No fue posible obtener valor Release
        REM set /a ERROR_FLAG+=1
    REM ) else (
        REM if !NET_RELEASE! GEQ 528040 (
            REM echo [OK] .NET Framework 4.8 o superior detectado
            REM set "NET_VERSION=4.8+"
        REM ) else if !NET_RELEASE! GEQ 461808 (
            REM echo [OK] .NET Framework 4.7.2 detectado
            REM set "NET_VERSION=4.7.2"
        REM ) else if !NET_RELEASE! GEQ 461308 (
            REM echo [OK] .NET Framework 4.7.1 detectado
            REM set "NET_VERSION=4.7.1"
        REM ) else if !NET_RELEASE! GEQ 460798 (
            REM echo [OK] .NET Framework 4.7 detectado
            REM set "NET_VERSION=4.7"
        REM ) else if !NET_RELEASE! GEQ 394802 (
            REM echo [OK] .NET Framework 4.6.2 detectado
            REM set "NET_VERSION=4.6.2"
        REM ) else if !NET_RELEASE! GEQ 394254 (
            REM echo [OK] .NET Framework 4.6.1 detectado
            REM set "NET_VERSION=4.6.1"
        REM ) else if !NET_RELEASE! GEQ 393295 (
            REM echo [OK] .NET Framework 4.6 detectado
            REM set "NET_VERSION=4.6"
        REM ) else if !NET_RELEASE! GEQ 379893 (
            REM echo [OK] .NET Framework 4.5.2 detectado
            REM set "NET_VERSION=4.5.2"
        REM ) else if !NET_RELEASE! GEQ 378675 (
            REM echo [OK] .NET Framework 4.5.1 detectado
            REM set "NET_VERSION=4.5.1"
        REM ) else if !NET_RELEASE! GEQ 378389 (
            REM echo [OK] .NET Framework 4.5 detectado
            REM set "NET_VERSION=4.5"
        REM ) else (
            REM echo [OK] .NET Framework 4.0 detectado
            REM set "NET_VERSION=4.0"
        REM )
    REM )
REM ) else (
REM pause
    REM :: SEGUNDO MÉTODO: Verificar .NET 3.5
    REM reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5" 2>nul
    REM if %errorlevel% equ 0 (
        REM echo [ADVERTENCIA] .NET Framework 3.5 detectado (4.0+ recomendado)
        REM set "NET_VERSION=3.5"
        REM set /a WARNING_FLAG+=1
    REM ) else (
        REM :: TERCER MÉTODO: Verificar archivos físicos de .NET
        REM if exist "%windir%\Microsoft.NET\Framework64\v4.0.30319\System.dll" (
            REM echo [OK] .NET Framework 4.0+ detectado (via archivos del sistema)
            REM set "NET_VERSION=4.0+"
        REM ) else if exist "%windir%\Microsoft.NET\Framework\v4.0.30319\System.dll" (
            REM echo [OK] .NET Framework 4.0+ detectado (via archivos del sistema)
            REM set "NET_VERSION=4.0+"
        REM ) else if exist "%windir%\Microsoft.NET\Framework\v3.5\System.dll" (
            REM echo [ADVERTENCIA] .NET Framework 3.5 detectado (via archivos del sistema)
            REM set "NET_VERSION=3.5"
            REM set /a WARNING_FLAG+=1
        REM ) else (
            REM :: CUARTO MÉTODO: Verificar carpetas de .NET
            REM dir "%windir%\Microsoft.NET\Framework\v4.0*" >nul 2>&1
            REM if %errorlevel% equ 0 (
                REM echo [OK] .NET Framework 4.x detectado (via carpeta)
                REM set "NET_VERSION=4.x"
            REM ) else (
                REM dir "%windir%\Microsoft.NET\Framework\v3.5*" >nul 2>&1
                REM if %errorlevel% equ 0 (
                    REM echo [ADVERTENCIA] .NET Framework 3.5 detectado (via carpeta)
                    REM set "NET_VERSION=3.5"
                    REM set /a WARNING_FLAG+=1
                REM ) else (
                    REM echo [ERROR] .NET Framework no detectado
                    REM echo [ERROR]   Algunas funciones avanzadas no estarán disponibles
                    REM set "NET_VERSION=No detectado"
                    REM set /a ERROR_FLAG+=1
                REM )
            REM )
        REM )
    REM )
REM )
:: ===================================================================
:: VERIFICACIÓN DE .NET FRAMEWORK - SIMPLIFICADA Y NO CRÍTICA
:: ===================================================================
:: Verificar .NET Framework (método simple y no crítico)
echo Verificando .NET Framework...
set "NET_VERSION=No requerido"

:: Solo una verificación simple sin lógica compleja
reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release 2>nul >nul
if !errorlevel! equ 0 (
    echo [OK] .NET Framework detectado
    set "NET_VERSION=4.0+"
) else (
    echo [INFO] .NET Framework no detectado
    echo [INFO]   No afecta al funcionamiento básico del sistema
)

REM :: Verificar Microsoft Excel (método mejorado)
REM echo Verificando Microsoft Excel...
REM set "EXCEL_FOUND=0"

REM :: Buscar en registro de 64 bits - USANDO FIND en lugar de FINDSTR
REM reg query "HKLM\SOFTWARE\Microsoft\Office" /s 2>nul | find /i "Excel" >nul
REM if !errorlevel! equ 0 set "EXCEL_FOUND=1"

REM :: Buscar en registro de 32 bits (en sistema de 64 bits)
REM reg query "HKLM\SOFTWARE\Wow6432Node\Microsoft\Office" /s 2>nul | find /i "Excel" >nul
REM if !errorlevel! equ 0 set "EXCEL_FOUND=1"

REM :: Buscar en registro de usuario
REM reg query "HKCU\SOFTWARE\Microsoft\Office" /s 2>nul | find /i "Excel" >nul
REM if !errorlevel! equ 0 set "EXCEL_FOUND=1"

REM if !EXCEL_FOUND! equ 1 (
    REM set "EXCEL_INSTALLED=1"
    REM echo [OK] Microsoft Excel detectado
REM ) else (
    REM echo [ADVERTENCIA] Microsoft Excel no detectado
    REM echo [ADVERTENCIA]   Se crearán archivos CSV como alternativa
    REM echo [ADVERTENCIA]   Se recomienda instalar Excel para todas las funciones
    REM set /a WARNING_FLAG+=2
REM )

:: Verificar espacio en disco (método directo y confiable)
echo Verificando espacio en disco...
set "FREE_SPACE_MB=0"

:: Método 1: Usar fsutil (más directo en Windows 10/11)
fsutil volume diskfree %SystemDrive% > "%TEMP%\fsinfo.txt" 2>nul
if !errorlevel! equ 0 (
    for /f "tokens=3" %%a in ('type "%TEMP%\fsinfo.txt" ^| find "Disponible"') do (
        set "FREE_SPACE_BYTES=%%a"
    )
    
    if "!FREE_SPACE_BYTES!" neq "" (
        :: Convertir a MB (1 MB = 1048576 bytes)
        set /a FREE_SPACE_MB=!FREE_SPACE_BYTES! / 1048576 2>nul
    )
    del "%TEMP%\fsinfo.txt" 2>nul
)

:: Si aún no tenemos el valor, usar PowerShell
if "!FREE_SPACE_MB!"=="0" (
    for /f "delims=" %%m in ('powershell -Command "(Get-PSDrive -Name %SystemDrive:~0,1%).Free / 1MB" 2^>nul') do (
        set "FREE_SPACE_MB=%%m"
    )
)

:: Si aún no, usar wmic de otra forma
if "!FREE_SPACE_MB!"=="0" (
    for /f "skip=1 tokens=3" %%a in ('wmic logicaldisk where "DeviceID='%SystemDrive%'" get FreeSpace^,Size^,DeviceID /format:csv 2^>nul') do (
        set "FREE_SPACE_BYTES=%%a"
    )
    if "!FREE_SPACE_BYTES!" neq "" (
        set /a FREE_SPACE_MB=!FREE_SPACE_BYTES! / 1048576 2>nul
    )
)

:: Mostrar resultado
if !FREE_SPACE_MB! LSS 100 (
    echo [ADVERTENCIA CRÍTICA] Espacio libre en disco bajo: !FREE_SPACE_MB! MB
    echo [ADVERTENCIA]   Se recomienda al menos 100MB de espacio libre
    set /a WARNING_FLAG+=3
) else if !FREE_SPACE_MB! GTR 0 (
    echo [OK] Espacio en disco suficiente: !FREE_SPACE_MB! MB libres
) else (
    echo [ADVERTENCIA] No se pudo verificar el espacio en disco
    set /a WARNING_FLAG+=1
)

:: Verificar memoria RAM disponible (método robusto)
echo Verificando memoria RAM...
set "RAM_MB=0"

:: Método 1: Usar wmic
wmic OS get FreePhysicalMemory /value > "%TEMP%\raminfo.txt" 2>nul
if %errorlevel% equ 0 (
    for /f "tokens=2 delims==" %%a in ('type "%TEMP%\raminfo.txt" ^| find "FreePhysicalMemory"') do (
        set "RAM_KB=%%a"
    )
    
    if "!RAM_KB!" neq "" (
        set /a "RAM_MB=!RAM_KB! / 1024" 2>nul
        echo [OK] Memoria RAM disponible: !RAM_MB! MB
    ) else (
        echo [INFO] Memoria RAM: Información no disponible
    )
    del "%TEMP%\raminfo.txt" 2>nul
) else (
    :: Método 2: Usar PowerShell
    powershell -Command "Get-WmiObject Win32_OperatingSystem | Select-Object -ExpandProperty FreePhysicalMemory" > "%TEMP%\ram.txt" 2>&1
    if %errorlevel% equ 0 (
        set /p RAM_KB= < "%TEMP%\ram.txt" 2>nul
        if "!RAM_KB!" neq "" (
            set /a "RAM_MB=!RAM_KB! / 1024" 2>nul
            echo [OK] Memoria RAM disponible: !RAM_MB! MB
        ) else (
            echo [INFO] Memoria RAM: Información no disponible
        )
    ) else (
        echo [INFO] Memoria RAM: Verificación no disponible
    )
    del "%TEMP%\ram.txt" 2>nul
)

:: Resumen de verificación
echo.
echo ===================================================
echo RESUMEN DE VERIFICACIÓN:
echo ===================================================
if !ERROR_FLAG! EQU 0 (
    echo Errores críticos: NINGUNO
) else (
    echo Errores críticos: !ERROR_FLAG!
)
echo Advertencias: !WARNING_FLAG!
echo PowerShell: !POWERSHELL_VERSION!
echo .NET Framework: !NET_VERSION!
echo Excel: !EXCEL_INSTALLED! (1=Sí, 0=No)
if !FREE_SPACE_MB! GTR 0 echo Espacio libre: !FREE_SPACE_MB! MB
if "!RAM_MB!" neq "" echo RAM disponible: !RAM_MB! MB
echo ===================================================

if !ERROR_FLAG! GEQ 3 (
    echo.
    echo [ERROR] Demasiados errores críticos. Abortando instalación.
    timeout /t 10 >nul
    exit /b 1
)

if !WARNING_FLAG! GEQ 5 (
    echo.
    echo [ADVERTENCIA] Muchas advertencias detectadas.
    echo [ADVERTENCIA] El sistema puede no funcionar correctamente.
)

echo.
set /p CONTINUAR="¿Desea continuar con la instalación? (S/N): "
if /i "!CONTINUAR!" NEQ "S" (
    echo [INFO] Instalación cancelada por el usuario.
    timeout /t 3 >nul
    exit /b 0
)

:: ===================================================================
:: FASE 2: PREPARACIÓN DEL ENTORNO MEJORADA
:: ===================================================================
echo.
echo [PROGRESO] FASE 2: Preparando entorno de instalación...
echo.

:: Verificar si el proyecto ya existe
if exist "!PROJECT_ROOT!" (
    echo [ATENCIÓN] El proyecto ya existe en: !PROJECT_ROOT!
    
    :: Crear backup con timestamp
    set "BACKUP_DIR=!PROJECT_ROOT!\_backup_%date:~-4,4%%date:~-7,2%%date:~-10,2%_%time:~0,2%%time:~3,2%"
    echo Creando backup en: !BACKUP_DIR!
    
    :: Copiar con robocopy (más robusto que xcopy)
    robocopy "!PROJECT_ROOT!" "!BACKUP_DIR!" /E /COPYALL /R:3 /W:5 /LOG:"%TEMP%\backup_log.txt" >nul
    if %errorlevel% LSS 8 (
        echo [OK] Backup creado exitosamente
        echo [INFO] Log de backup: %TEMP%\backup_log.txt
    ) else (
        echo [ERROR] No se pudo crear backup completo
        echo [INFO] Se intentó continuar con la instalación...
    )
    
    :: Preguntar confirmación
    echo.
    set /p CONFIRM_OVERWRITE="¿Desea reinstalar el sistema? (S/N): "
    if /i "!CONFIRM_OVERWRITE!" NEQ "S" (
        echo [INFO] Instalación cancelada por el usuario
        echo.
        echo Instalación cancelada. El sistema existente no ha sido modificado.
        echo Backup disponible en: !BACKUP_DIR!
        timeout /t 5 >nul
        exit /b 0
    )
    
    :: Limpiar instalación anterior de forma segura
    echo Eliminando instalación anterior...
    
    :: Primero eliminar archivos individuales
    del /q "!PROJECT_ROOT!\*.*" >nul 2>&1
    
    :: Luego eliminar carpetas vacías
    for /d %%d in ("!PROJECT_ROOT!\*") do (
        rmdir "%%d" /s /q >nul 2>&1
    )
    
    :: Esperar a que se liberen los recursos
    timeout /t 2 /nobreak >nul
)

:: ===================================================================
:: FASE 3: CREACIÓN DE ESTRUCTURA DE CARPETAS MEJORADA
:: ===================================================================
echo.
echo [PROGRESO] FASE 3: Creando estructura de carpetas...
echo.

:: Crear carpeta principal con verificación
mkdir "!PROJECT_ROOT!" 2>nul
if not exist "!PROJECT_ROOT!" (
    echo [ERROR CRÍTICO] No se pudo crear la carpeta principal
    echo [ERROR] Verifique permisos y espacio en disco.
    timeout /t 5 >nul
    exit /b 1
)

echo [OK] Carpeta principal creada: !PROJECT_ROOT!

:: Lista completa de carpetas principales
set "MAIN_FOLDERS=Data_Backup Configuraciones Scripts_IA Reportes Tickets Templates Logs Cache Exportaciones Datos_Externos Plantillas_IA Modelos_ML Modulos Documentacion Temp Sesiones"

echo Creando carpetas principales...
for %%f in (!MAIN_FOLDERS!) do (
    mkdir "!PROJECT_ROOT!\%%f" 2>nul
    if exist "!PROJECT_ROOT!\%%f" (
        echo   [?] !PROJECT_ROOT!\%%f
    ) else (
        echo   [?] Error creando: !PROJECT_ROOT!\%%f
        set /a ERROR_FLAG+=1
    )
)

:: Crear subcarpetas especializadas
echo Creando subcarpetas especializadas...

:: Data_Backup
mkdir "!PROJECT_ROOT!\Data_Backup\Diario" 2>nul
mkdir "!PROJECT_ROOT!\Data_Backup\Semanal" 2>nul
mkdir "!PROJECT_ROOT!\Data_Backup\Mensual" 2>nul
mkdir "!PROJECT_ROOT!\Data_Backup\Automatico" 2>nul
mkdir "!PROJECT_ROOT!\Data_Backup\Manual" 2>nul

:: Configuraciones
mkdir "!PROJECT_ROOT!\Configuraciones\Usuarios" 2>nul
mkdir "!PROJECT_ROOT!\Configuraciones\Sistema" 2>nul
mkdir "!PROJECT_ROOT!\Configuraciones\APIs" 2>nul
mkdir "!PROJECT_ROOT!\Configuraciones\Plantillas" 2>nul

:: Scripts_IA
mkdir "!PROJECT_ROOT!\Scripts_IA\Analisis" 2>nul
mkdir "!PROJECT_ROOT!\Scripts_IA\Modelos" 2>nul
mkdir "!PROJECT_ROOT!\Scripts_IA\Utilidades" 2>nul
mkdir "!PROJECT_ROOT!\Scripts_IA\Pruebas" 2>nul

:: Reportes
mkdir "!PROJECT_ROOT!\Reportes\PDF" 2>nul
mkdir "!PROJECT_ROOT!\Reportes\Excel" 2>nul
mkdir "!PROJECT_ROOT!\Reportes\HTML" 2>nul
mkdir "!PROJECT_ROOT!\Reportes\Dashboard" 2>nul
mkdir "!PROJECT_ROOT!\Reportes\Automaticos" 2>nul

:: Tickets
mkdir "!PROJECT_ROOT!\Tickets\Imagenes" 2>nul
mkdir "!PROJECT_ROOT!\Tickets\PDF" 2>nul
mkdir "!PROJECT_ROOT!\Tickets\OCR" 2>nul
mkdir "!PROJECT_ROOT!\Tickets\Procesados" 2>nul

:: Templates
mkdir "!PROJECT_ROOT!\Templates\Email" 2>nul
mkdir "!PROJECT_ROOT!\Templates\Reportes" 2>nul
mkdir "!PROJECT_ROOT!\Templates\Documentos" 2>nul
mkdir "!PROJECT_ROOT!\Templates\Contratos" 2>nul

:: Logs
mkdir "!PROJECT_ROOT!\Logs\Sistema" 2>nul
mkdir "!PROJECT_ROOT!\Logs\Errores" 2>nul
mkdir "!PROJECT_ROOT!\Logs\Auditoria" 2>nul
mkdir "!PROJECT_ROOT!\Logs\Depuracion" 2>nul

:: Cache
mkdir "!PROJECT_ROOT!\Cache\Imagenes" 2>nul
mkdir "!PROJECT_ROOT!\Cache\Datos" 2>nul
mkdir "!PROJECT_ROOT!\Cache\Temporal" 2>nul
mkdir "!PROJECT_ROOT!\Cache\Sesiones" 2>nul

:: Exportaciones
mkdir "!PROJECT_ROOT!\Exportaciones\CSV" 2>nul
mkdir "!PROJECT_ROOT!\Exportaciones\Excel" 2>nul
mkdir "!PROJECT_ROOT!\Exportaciones\PDF" 2>nul
mkdir "!PROJECT_ROOT!\Exportaciones\JSON" 2>nul
mkdir "!PROJECT_ROOT!\Exportaciones\XML" 2>nul

:: Datos_Externos
mkdir "!PROJECT_ROOT!\Datos_Externos\APIs" 2>nul
mkdir "!PROJECT_ROOT!\Datos_Externos\WebScraping" 2>nul
mkdir "!PROJECT_ROOT!\Datos_Externos\Importados" 2>nul
mkdir "!PROJECT_ROOT!\Datos_Externos\Procesados" 2>nul

:: Plantillas_IA
mkdir "!PROJECT_ROOT!\Plantillas_IA\Modelos" 2>nul
mkdir "!PROJECT_ROOT!\Plantillas_IA\DatosEntrenamiento" 2>nul
mkdir "!PROJECT_ROOT!\Plantillas_IA\Resultados" 2>nul

:: Modelos_ML
mkdir "!PROJECT_ROOT!\Modelos_ML\Entrenados" 2>nul
mkdir "!PROJECT_ROOT!\Modelos_ML\EnEntrenamiento" 2>nul
mkdir "!PROJECT_ROOT!\Modelos_ML\Backup" 2>nul

:: Modulos
mkdir "!PROJECT_ROOT!\Modulos\VBA" 2>nul
mkdir "!PROJECT_ROOT!\Modulos\Python" 2>nul
mkdir "!PROJECT_ROOT!\Modulos\PowerShell" 2>nul
mkdir "!PROJECT_ROOT!\Modulos\SQL" 2>nul

:: Documentacion
mkdir "!PROJECT_ROOT!\Documentacion\Tecnica" 2>nul
mkdir "!PROJECT_ROOT!\Documentacion\Usuario" 2>nul
mkdir "!PROJECT_ROOT!\Documentacion\API" 2>nul
mkdir "!PROJECT_ROOT!\Documentacion\Cambios" 2>nul

:: Temp
mkdir "!PROJECT_ROOT!\Temp\Uploads" 2>nul
mkdir "!PROJECT_ROOT!\Temp\Downloads" 2>nul
mkdir "!PROJECT_ROOT!\Temp\Procesamiento" 2>nul

:: Sesiones
mkdir "!PROJECT_ROOT!\Sesiones\Usuarios" 2>nul
mkdir "!PROJECT_ROOT!\Sesiones\Sistema" 2>nul
mkdir "!PROJECT_ROOT!\Sesiones\Backup" 2>nul

echo [OK] Estructura de carpetas creada exitosamente
echo [INFO] Total: 15 carpetas principales con 58 subcarpetas

if !ERROR_FLAG! GTR 0 (
    echo [ADVERTENCIA] Se produjeron !ERROR_FLAG! errores creando carpetas
)

echo.
set /p CONTINUAR="Presione S y Enter para continuar con la FASE 4... "
if /i "!CONTINUAR!" NEQ "S" (
    echo [INFO] Instalación pausada por el usuario.
    timeout /t 3 >nul
    exit /b 0
)

:: ===================================================================
:: FASE 4: EJECUCIÓN DE SCRIPTS DE CONFIGURACIÓN MEJORADA - CORREGIDO
:: ===================================================================
echo.
echo [PROGRESO] FASE 4: Ejecutando scripts de configuración...
echo.

:: Configurar política de ejecución de PowerShell de forma segura
echo Configurando política de ejecución de PowerShell...
powershell -Command "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force" >nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] Política de ejecución configurada
) else (
    echo [ADVERTENCIA] No se pudo configurar política de ejecución
    set /a WARNING_FLAG+=1
)

:: Ejecutar scripts en orden con mejor manejo de errores
echo.
echo Ejecutando scripts de configuración...

:: Lista de scripts a ejecutar 
set "SCRIPTS=crear_excel.ps1 cargar_datos.ps1 configurar_sistema.ps1"

set "SCRIPT_SUCCESS=0"
set "SCRIPT_TOTAL=0"

:: DEBUG: Mostrar información sobre los scripts
echo [DEBUG] Scripts a ejecutar: !SCRIPTS!
echo [DEBUG] Directorio de scripts: !SCRIPT_DIR!
echo [DEBUG] Directorio del proyecto: !PROJECT_ROOT!
echo [DEBUG] Contenido exacto: %SCRIPTS%

for %%s in (%SCRIPTS%) do (
    set /a SCRIPT_TOTAL+=1
    echo.
    echo --------------------------------------
    echo Ejecutando script !SCRIPT_TOTAL!: %%s
    echo --------------------------------------
    
    :: Verificar si el script existe
    if exist "!SCRIPT_DIR!\%%s" (
		echo [INFO] Script encontrado: !SCRIPT_DIR!\%%s
        
        :: Ejecutar script con timeout y captura de errores
        echo [INFO] Ejecutando PowerShell script...
        :: Crear un archivo temporal para capturar la salida
        set "PS_OUTPUT_FILE=%TEMP%\ps_output_%%s_%time:~0,2%%time:~3,2%%time:~6,2%.txt"
		::Ejemplo de como crear bien el if en bash, se supone que usando un call no falla el for
		::if ["%%s" == "crear_excel.ps1"]; then
		::
		::else
		::
		::fi
		:: Ejecutar PowerShell script y capturar salida en modo silent
		powershell -NoProfile -ExecutionPolicy Bypass -File "!SCRIPT_DIR!\%%s" -ProjectPath "!PROJECT_ROOT!" -Silent > "!PS_OUTPUT_FILE!" 2>&1
		:: Ejecutar PowerShell script y capturar salida
		::powershell -NoProfile -ExecutionPolicy Bypass -File "!SCRIPT_DIR!\%%s" -ProjectPath "!PROJECT_ROOT!" > "!PS_OUTPUT_FILE!" 2>&1
		:: Ejecutar PowerShell script y capturar salida abriendo ventana
		::start /wait powershell -NoProfile -ExecutionPolicy Bypass -File "!SCRIPT_DIR!\%%s" -ProjectPath "!PROJECT_ROOT!" > "!PS_OUTPUT_FILE!" 2>&1
		:: Ejecutar PowerShell script y sin capturar salida abiendo ventana
		::start /wait powershell -NoProfile -ExecutionPolicy Bypass -File "!SCRIPT_DIR!\%%s" -ProjectPath "!PROJECT_ROOT!"
		set "SCRIPT_EXITCODE=!errorlevel!"
        
        :: Mostrar las primeras líneas de la salida
        echo [INFO] Mostrando salida del script:
        echo --------------------------------------
		if exist "!PS_OUTPUT_FILE!" (
            echo [INFO] Resumen de salida:
            type "!PS_OUTPUT_FILE!"
        )
        echo --------------------------------------
        
        :: Evaluar el código de salida
        if !SCRIPT_EXITCODE! equ 0 (
            echo [OK] %%s ejecutado exitosamente - Código: 0
            set /a SCRIPT_SUCCESS+=1
        ) else if !SCRIPT_EXITCODE! equ 1 (
            echo [ADVERTENCIA] %%s completado con advertencias - Código: 1
            set /a SCRIPT_SUCCESS+=1
            set /a WARNING_FLAG+=1
        ) else (
            echo [ERROR] Fallo al ejecutar: %%s - Código: !SCRIPT_EXITCODE!
            echo [INFO] Revisar archivo de log: !PS_OUTPUT_FILE!
            set /a ERROR_FLAG+=1
        )
        
        :: Limpiar archivo temporal si no hay errores graves
        if !SCRIPT_EXITCODE! leq 1 (
            del "!PS_OUTPUT_FILE!" 2>nul
        )
    ) else (
        echo [ERROR CRÍTICO] Script no encontrado: !SCRIPT_DIR!\%%s
        echo [INFO] Verifica que el archivo exista en la ubicación correcta.
        set /a ERROR_FLAG+=1
    )
    
    :: Pausa breve entre scripts
    timeout /t 1 /nobreak >nul
)

:: Si no hay scripts ejecutados, crear estructura básica
if !SCRIPT_TOTAL! equ 0 (
    echo [ADVERTENCIA] No se encontraron scripts para ejecutar
    echo [INFO] Creando estructura básica del proyecto...
    
    :: Crear archivo Excel básico si no existe
    if not exist "!PROJECT_ROOT!\Comparador_Compras_IA_Completo.xlsm" (
        echo [INFO] Creando archivo Excel básico...
        copy /y "!SCRIPT_DIR!\plantilla_excel.xlsm" "!PROJECT_ROOT!\Comparador_Compras_IA_Completo.xlsm" >nul 2>&1
        if errorlevel 1 (
            :: Si no hay plantilla, crear un archivo vacío
            echo. > "!PROJECT_ROOT!\Comparador_Compras_IA_Completo.xlsm"
        )
    )
)

:: Ejecutar script VBScript para macros (opcional)
if exist "!SCRIPT_DIR!\agregar_macros.vbs" (
    echo.
    echo --------------------------------------
    echo Ejecutando: agregar_macros.vbs
    echo --------------------------------------
    
    cscript //nologo "!SCRIPT_DIR!\agregar_macros.vbs" "!PROJECT_ROOT!"
    if !errorlevel! neq 0 (
        echo [ADVERTENCIA] Fallo al agregar macros (Código: !errorlevel!)
        set /a WARNING_FLAG+=1
    ) else (
        echo [OK] Macros agregadas exitosamente
    )
)

:: Resumen de ejecución de scripts
echo.
echo ===================================================
echo RESUMEN DE EJECUCIÓN DE SCRIPTS:
echo ===================================================
echo Scripts encontrados: !SCRIPT_TOTAL!
echo Scripts ejecutados exitosamente: !SCRIPT_SUCCESS!
echo Errores en esta fase: !ERROR_FLAG!
echo Advertencias en esta fase: !WARNING_FLAG!
echo ===================================================

if !SCRIPT_SUCCESS! equ 0 (
    echo [ADVERTENCIA CRÍTICA] Ningún script se ejecutó correctamente
    echo [INFO] Continuando con instalación básica...
) else if !SCRIPT_SUCCESS! LSS !SCRIPT_TOTAL! (
    echo [ADVERTENCIA] No todos los scripts se ejecutaron correctamente
    echo [INFO] Algunas funciones pueden estar limitadas
)

echo.
echo Presione cualquier tecla para continuar con la FASE 5...
pause >nul

:: ===================================================================
:: FASE 5: CREACIÓN DE ARCHIVOS DE CONFIGURACIÓN MEJORADA
:: ===================================================================
echo.
echo [PROGRESO] FASE 5: Creando archivos de configuración...
echo.

:: El archivo config_sistema.json ahora es creado por configurar_sistema.ps1
:: Verificamos que se haya creado correctamente
if exist "!PROJECT_ROOT!\Configuraciones\config_sistema.json" (
    echo [OK] Archivo de configuración principal creado por configurar_sistema.ps1
) else (
    echo [ADVERTENCIA] No se encontró config_sistema.json
    echo [INFO] Creando versión básica...
    
    (
    echo {
    echo   "sistema": {
    echo     "version": "!SCRIPT_VERSION!",
    echo     "fecha_instalacion": "!FECHA_INSTALACION!",
    echo     "sistema_operativo": "%OS%",
    echo     "arquitectura": "%PROCESSOR_ARCHITECTURE%",
    echo     "usuario": "%USERNAME%",
    echo     "equipo": "%COMPUTERNAME%",
    echo     "powershell_version": "!POWERSHELL_VERSION!",
    echo     "net_version": "!NET_VERSION!",
    echo     "excel_instalado": !EXCEL_INSTALLED!
    echo   }
    echo }
    ) > "!PROJECT_ROOT!\Configuraciones\config_sistema.json" 2>nul
    
    if exist "!PROJECT_ROOT!\Configuraciones\config_sistema.json" (
        echo [OK] Configuración básica creada
    ) else (
        echo [ERROR] No se pudo crear configuración básica
        set /a ERROR_FLAG+=1
    )
)

:: Archivo de instrucciones mejorado (actualizado)
echo Creando INSTRUCCIONES_PROYECTO.txt...
(
echo ===================================================
echo    SISTEMA COMPARADOR DE COMPRAS INTELIGENTE IA
echo    Versión: !SCRIPT_VERSION! - Edición Empresarial
echo ===================================================
echo.
echo ?? CONFIGURACIÓN DEL SISTEMA
echo ----------------------------------------------------
echo.
echo FECHA DE INSTALACIÓN: !FECHA_INSTALACION!
echo USUARIO: %USERNAME%
echo EQUIPO: %COMPUTERNAME%
echo SISTEMA: %OS% !ARCH! bits
echo POWERSHELL: !POWERSHELL_VERSION!
echo .NET FRAMEWORK: !NET_VERSION!
echo EXCEL: !EXCEL_INSTALLED! (1=Instalado, 0=No instalado)
echo.
echo ?? UBICACIÓN DEL PROYECTO: !PROJECT_ROOT!
echo.
echo ?? SCRIPTS DE CONFIGURACIÓN EJECUTADOS: !SCRIPT_SUCCESS!/!SCRIPT_TOTAL!
echo.
echo ??  ADVERTENCIAS: !WARNING_FLAG!
echo ? ERRORES: !ERROR_FLAG!
echo.
echo ----------------------------------------------------
echo ?? INICIO RÁPIDO
echo ----------------------------------------------------
echo.
echo 1. ?? ACCESO DIRECTO: Busque "Comparador Compras IA" en su escritorio
echo 2. ?? EXCEL PRINCIPAL: Abra Comparador_Compras_IA_Completo.xlsm
echo 3. ? HABILITAR MACROS: Permita la ejecución cuando se le solicite
echo 4. ?? MENÚ PRINCIPAL: Use el menú "Comparador IA" en Excel
echo 5. ?? CONFIGURACIÓN: Complete sus datos en la hoja USUARIOS
echo.
echo ----------------------------------------------------
echo ?? ESTRUCTURA DEL PROYECTO
echo ----------------------------------------------------
echo.
echo ?? Data_Backup/        - Sistema de backups automáticos
echo ?? Configuraciones/    - Archivos de configuración JSON/XML
echo ?? Scripts_IA/         - Scripts PowerShell y Python
echo ?? Reportes/           - Reportes PDF, Excel y HTML
echo ?? Tickets/            - Tickets escaneados y procesados
echo ?? Templates/          - Plantillas de email y documentos
echo ?? Logs/               - Registros del sistema
echo ?? Cache/              - Datos temporales en caché
echo ?? Exportaciones/      - Datos para exportar
echo ?? Datos_Externos/     - Datos de APIs y web scraping
echo ?? Plantillas_IA/      - Modelos de IA
echo ?? Modelos_ML/         - Modelos de machine learning
echo ?? Modulos/            - Módulos VBA, Python, PowerShell
echo ?? Documentacion/      - Documentación técnica y de usuario
echo ?? Temp/               - Archivos temporales
echo ?? Sesiones/           - Datos de sesiones de usuario
echo.
echo ----------------------------------------------------
echo ???  HERRAMIENTAS Y UTILIDADES
echo ----------------------------------------------------
echo.
echo ?? Scripts de utilidad incluidos:
echo   • backup_automatico.ps1    - Sistema de backups programados
echo   • limpiar_cache.ps1        - Limpieza de caché del sistema
echo   • verificar_sistema.ps1    - Diagnóstico del sistema
echo.
echo ?? Archivos de configuración:
echo   • config_sistema.json      - Configuración principal
echo   • config_%USERNAME%.json   - Configuración de usuario
echo   • conexiones.xml           - Configuración de APIs
echo   • seguridad.json           - Configuración de seguridad
echo   • backup.json              - Configuración de backups
echo.
echo ----------------------------------------------------
echo ?? SOLUCIÓN DE PROBLEMAS
echo ----------------------------------------------------
echo.
echo ? Si Excel no abre o da errores:
echo   1. Verifique que tenga Microsoft Excel 2016 o superior
echo   2. Asegúrese de habilitar macros
echo   3. Ejecute verificar_sistema.ps1 para diagnóstico
echo.
echo ??  Si aparecen errores de PowerShell:
echo   1. Ejecute PowerShell como administrador
echo   2. Ejecute: Set-ExecutionPolicy RemoteSigned
echo   3. Reinstale el sistema si es necesario
echo.
echo ?? Si los datos no se cargan:
echo   1. Verifique los archivos CSV en Datos_Externos\
echo   2. Revise los logs en Logs\Errores\
echo   3. Ejecute cargar_datos.ps1 manualmente
echo.
echo ----------------------------------------------------
echo ?? SOPORTE Y MANTENIMIENTO
echo ----------------------------------------------------
echo.
echo ?? Actualizaciones automáticas: Habilitadas
echo ?? Backup automático: Cada 24 horas
echo ?? Logs detallados: En carpeta Logs\
echo ???  Seguridad: Validación de datos y hashing
echo.
echo ----------------------------------------------------
echo ?? PRÓXIMOS PASOS RECOMENDADOS
echo ----------------------------------------------------
echo.
echo 1. ?? COMPLETAR CONFIGURACIÓN INICIAL (HOY)
echo    • Complete sus datos en USUARIOS
echo    • Añada al menos 3 tiendas locales
echo    • Registre 5 productos frecuentes
echo.
echo 2. ?? PRIMER ANÁLISIS (PRÓXIMA SEMANA)
echo    • Ingrese precios de 2-3 supermercados
echo    • Genere su primera comparación
echo    • Revise el reporte automático
echo.
echo 3. ?? AUTOMATIZACIÓN (EN 2 SEMANAS)
echo    • Configure alertas de precio
echo    • Programe backups automáticos
echo    • Explore scripts de IA avanzados
echo.
echo ----------------------------------------------------
echo ?? FUNCIONALIDADES PRINCIPALES
echo ----------------------------------------------------
echo.
echo ? COMPARACIÓN INTELIGENTE
echo    • Análisis de precios en tiempo real
echo    • Histórico de precios y tendencias
echo    • Alertas automáticas de ofertas
echo.
echo ???  OPTIMIZACIÓN DE RUTAS
echo    • Cálculo de rutas más eficientes
echo    • Consideración de tráfico y horarios
echo    • Multi-destino inteligente
echo.
echo ?? INTELIGENCIA ARTIFICIAL
echo    • Recomendaciones personalizadas
echo    • Predicción de precios futuros
echo    • Detección de patrones de compra
echo.
echo ?? REPORTES AVANZADOS
echo    • Dashboards interactivos
echo    • Exportación a múltiples formatos
echo    • Análisis estadístico completo
echo.
echo ===================================================
echo    ¡SISTEMA INSTALADO Y CONFIGURADO EXITOSAMENTE!
echo ===================================================
echo.
echo ?? CONSEJO FINAL: Revise regularmente los logs y
echo    realice backups manuales antes de cambios grandes.
echo.
) > "!PROJECT_ROOT!\INSTRUCCIONES_PROYECTO.txt"

if exist "!PROJECT_ROOT!\INSTRUCCIONES_PROYECTO.txt" (
    echo [OK] Instrucciones del proyecto creadas
) else (
    echo [ERROR] No se pudo crear archivo de instrucciones
    set /a ERROR_FLAG+=1
)

:: Crear archivo de licencia actualizado
echo Creando LICENCIA.txt...
(
echo LICENCIA DE USO - SISTEMA COMPARADOR DE COMPRAS IA
echo ===================================================
echo.
echo Versión del sistema: !SCRIPT_VERSION!
echo Fecha de instalación: !FECHA_INSTALACION!
echo Usuario licenciado: %USERNAME%
echo Equipo: %COMPUTERNAME%
echo.
echo ----------------------------------------------------
echo TÉRMINOS DE USO Y LICENCIA
echo ----------------------------------------------------
echo.
echo 1. LICENCIA DE USO
echo   1.1. Esta licencia permite el uso personal y empresarial.
echo   1.2. Se permite la instalación en hasta 3 dispositivos.
echo   1.3. No se permite la reventa o distribución comercial.
echo.
echo 2. RESPONSABILIDADES DEL USUARIO
echo   2.1. El usuario es responsable de la veracidad de los datos.
echo   2.2. Debe realizar copias de seguridad regularmente.
echo   2.3. Debe mantener el sistema actualizado.
echo.
echo 3. LIMITACIONES DE GARANTÍA
echo   3.1. El software se proporciona "TAL CUAL".
echo   3.2. No hay garantías de funcionamiento ininterrumpido.
echo   3.3. El desarrollador no se hace responsable por pérdidas.
echo.
echo 4. PROPIEDAD INTELECTUAL
echo   4.1. Todos los derechos de autor son reservados.
echo   4.2. El código fuente permanece propiedad del desarrollador.
echo   4.3. Se permite la modificación para uso personal.
echo.
echo 5. DISTRIBUCIÓN
echo   5.1. Puede distribuirse libremente manteniendo esta licencia.
echo   5.2. Debe incluirse completa la documentación.
echo   5.3. No se permite la distribución modificada sin autorización.
echo.
echo ----------------------------------------------------
echo ACEPTACIÓN DE TÉRMINOS
echo ----------------------------------------------------
echo.
echo Al utilizar este software, usted acepta:
echo • Los términos de esta licencia.
echo • Las limitaciones de garantía establecidas.
echo • Ser responsable del uso adecuado del sistema.
echo.
echo ----------------------------------------------------
echo INFORMACIÓN DE CONTACTO
echo ----------------------------------------------------
echo.
echo Para soporte técnico o preguntas sobre la licencia:
echo • Consulte la documentación incluida.
echo • Revise los archivos de log para diagnóstico.
echo • Contacte al desarrollador si es necesario.
echo.
echo ===================================================
echo © 2024 Sistema Comparador de Compras IA v!SCRIPT_VERSION!
echo Todos los derechos reservados.
echo ===================================================
) > "!PROJECT_ROOT!\LICENCIA.txt"

if exist "!PROJECT_ROOT!\LICENCIA.txt" (
    echo [OK] Archivo de licencia creado
) else (
    echo [ERROR] No se pudo crear archivo de licencia
    set /a ERROR_FLAG+=1
)

echo.
echo [OK] Archivos de documentación creados exitosamente

echo.
set /p CONTINUAR="Presione S y Enter para continuar con la FASE 6... "
if /i "!CONTINUAR!" NEQ "S" (
    echo [INFO] Instalación pausada por el usuario.
    timeout /t 3 >nul
    exit /b 0
)

:: ===================================================================
:: FASE 6: CREACIÓN DE ACCESOS DIRECTOS MEJORADA
:: ===================================================================
echo.
echo [PROGRESO] FASE 6: Creando accesos directos...
echo.

:: Acceso directo en escritorio (mejorado)
set "DESKTOP_SHORTCUT=%USERPROFILE%\Desktop\Comparador Compras IA.lnk"
set "DESKTOP_SHORTCUT2=%USERPROFILE%\Desktop\Comparador IA - Abrir Carpeta.lnk"

echo Creando accesos directos en el escritorio...

:: Acceso directo 1: Archivo Excel principal
if not exist "!DESKTOP_SHORTCUT!" (
    (
    echo Set oWS = WScript.CreateObject("WScript.Shell")
    echo sLinkFile = "!DESKTOP_SHORTCUT!"
    echo Set oLink = oWS.CreateShortcut(sLinkFile)
    echo oLink.TargetPath = "!PROJECT_ROOT!\Comparador_Compras_IA_Completo.xlsm"
    echo oLink.WorkingDirectory = "!PROJECT_ROOT!"
    echo oLink.Description = "Sistema Comparador de Compras Inteligente IA v!SCRIPT_VERSION!"
    echo oLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll,165"
    echo oLink.Save
    ) > "%TEMP%\crear_acceso_excel.vbs"
    
    cscript //nologo "%TEMP%\crear_acceso_excel.vbs" >nul 2>&1
    del "%TEMP%\crear_acceso_excel.vbs" 2>nul
    
    if exist "!DESKTOP_SHORTCUT!" (
        echo [OK] Acceso directo creado: Comparador Compras IA.lnk
    ) else (
        echo [ADVERTENCIA] No se pudo crear acceso directo principal
        set /a WARNING_FLAG+=1
    )
) else (
    echo [INFO] Acceso directo principal ya existe
)

:: Acceso directo 2: Carpeta del proyecto
if not exist "!DESKTOP_SHORTCUT2!" (
    (
    echo Set oWS = WScript.CreateObject("WScript.Shell")
    echo sLinkFile = "!DESKTOP_SHORTCUT2!"
    echo Set oLink = oWS.CreateShortcut(sLinkFile)
    echo oLink.TargetPath = "!PROJECT_ROOT!"
    echo oLink.WorkingDirectory = "!PROJECT_ROOT!"
    echo oLink.Description = "Abrir carpeta del proyecto - Sistema Comparador IA"
    echo oLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll,4"
    echo oLink.Save
    ) > "%TEMP%\crear_acceso_carpeta.vbs"
    
    cscript //nologo "%TEMP%\crear_acceso_carpeta.vbs" >nul 2>&1
    del "%TEMP%\crear_acceso_carpeta.vbs" 2>nul
    
    if exist "!DESKTOP_SHORTCUT2!" (
        echo [OK] Acceso directo creado: Comparador IA - Abrir Carpeta.lnk
    )
)

:: Acceso directo en menú inicio (solo con permisos de admin)
if !ADMIN_MODE! equ 1 (
    echo Creando acceso directo en menú Inicio...
    
    set "START_MENU_DIR=%ProgramData%\Microsoft\Windows\Start Menu\Programs\Comparador Compras IA"
    mkdir "!START_MENU_DIR!" 2>nul
    
    if exist "!START_MENU_DIR!" (
        (
        echo Set oWS = WScript.CreateObject("WScript.Shell")
        echo sLinkFile = "!START_MENU_DIR!\Comparador Compras IA.lnk"
        echo Set oLink = oWS.CreateShortcut(sLinkFile)
        echo oLink.TargetPath = "!PROJECT_ROOT!\Comparador_Compras_IA_Completo.xlsm"
        echo oLink.WorkingDirectory = "!PROJECT_ROOT!"
        echo oLink.Description = "Sistema Comparador de Compras Inteligente IA"
        echo oLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll,165"
        echo oLink.Save
        ) > "%TEMP%\crear_acceso_startmenu.vbs"
        
        cscript //nologo "%TEMP%\crear_acceso_startmenu.vbs" >nul 2>&1
        del "%TEMP%\crear_acceso_startmenu.vbs" 2>nul
        
        if exist "!START_MENU_DIR!\Comparador Compras IA.lnk" (
            echo [OK] Acceso directo creado en el menú Inicio
        )
    )
) else (
    echo [INFO] Acceso directo en menú Inicio omitido (sin permisos de admin)
)

echo.
echo [OK] Accesos directos configurados

echo.
set /p CONTINUAR="Presione S y Enter para continuar con la FASE 7... "
if /i "!CONTINUAR!" NEQ "S" (
    echo [INFO] Instalación pausada por el usuario.
    timeout /t 3 >nul
    exit /b 0
)

:: ===================================================================
:: FASE 7: VERIFICACIÓN FINAL MEJORADA
:: ===================================================================
echo.
echo [PROGRESO] FASE 7: Realizando verificación final del sistema...
echo.

:: Verificar archivos esenciales creados
echo Verificando archivos esenciales...
set "ESSENTIAL_FILES=Comparador_Compras_IA_Completo.xlsm INSTRUCCIONES_PROYECTO.txt LICENCIA.txt"
set "ESSENTIAL_CONFIGS=Configuraciones\config_sistema.json Configuraciones\Sistema\seguridad.json Configuraciones\Sistema\backup.json"

set "FILES_FOUND=0"
set "FILES_TOTAL=0"

:: Contar archivos esenciales
for %%f in (!ESSENTIAL_FILES!) do set /a FILES_TOTAL+=1
for %%f in (!ESSENTIAL_CONFIGS!) do set /a FILES_TOTAL+=1

:: Verificar archivos
for %%f in (!ESSENTIAL_FILES!) do (
    if exist "!PROJECT_ROOT!\%%f" (
        set /a FILES_FOUND+=1
        echo   [?] %%f encontrado
    ) else (
        echo   [?] %%f NO encontrado
        set /a ERROR_FLAG+=1
    )
)

for %%f in (!ESSENTIAL_CONFIGS!) do (
    if exist "!PROJECT_ROOT!\%%f" (
        set /a FILES_FOUND+=1
        echo   [?] %%f encontrado
    ) else (
        echo   [?] %%f NO encontrado
        set /a WARNING_FLAG+=1
    )
)

if !FILES_FOUND! equ !FILES_TOTAL! (
    echo [OK] Todos los archivos esenciales están presentes (!FILES_FOUND!/!FILES_TOTAL!)
) else (
    echo [ADVERTENCIA] Faltan algunos archivos: !FILES_FOUND!/!FILES_TOTAL!
)

:: Verificar permisos de escritura exhaustivos
echo Verificando permisos de escritura...
set "TEST_PATHS=!PROJECT_ROOT! !PROJECT_ROOT!\Logs !PROJECT_ROOT!\Cache !PROJECT_ROOT!\Temp"
set "WRITE_TEST_PASSED=0"
set "WRITE_TEST_TOTAL=0"

for %%p in (!TEST_PATHS!) do (
    set /a WRITE_TEST_TOTAL+=1
    echo test > "%%p\test_write_!WRITE_TEST_TOTAL!.tmp" 2>nul
    if exist "%%p\test_write_!WRITE_TEST_TOTAL!.tmp" (
        del "%%p\test_write_!WRITE_TEST_TOTAL!.tmp" 2>nul
        set /a WRITE_TEST_PASSED+=1
        echo   [?] Permisos en: %%p
    ) else (
        echo   [?] Sin permisos en: %%p
        set /a ERROR_FLAG+=1
    )
)

if !WRITE_TEST_PASSED! equ !WRITE_TEST_TOTAL! (
    echo [OK] Permisos de escritura verificados (!WRITE_TEST_PASSED!/!WRITE_TEST_TOTAL!)
) else (
    echo [ERROR] Problemas con permisos de escritura (!WRITE_TEST_PASSED!/!WRITE_TEST_TOTAL!)
)

:: Verificar integridad del Excel
echo Verificando integridad del archivo Excel...
if exist "!PROJECT_ROOT!\Comparador_Compras_IA_Completo.xlsm" (
    for %%a in ("!PROJECT_ROOT!\Comparador_Compras_IA_Completo.xlsm") do (
        set "EXCEL_SIZE=%%~za"
    )
    
    if !EXCEL_SIZE! GTR 10240 (
        echo [OK] Archivo Excel válido (!EXCEL_SIZE! bytes)
    ) else (
        echo [ERROR] Archivo Excel sospechosamente pequeño (!EXCEL_SIZE! bytes)
        set /a ERROR_FLAG+=1
    )
) else (
    echo [ERROR CRÍTICO] Archivo Excel principal no encontrado
    set /a ERROR_FLAG+=2
)

:: Verificar scripts de utilidad
echo Verificando scripts de utilidad...
if exist "!PROJECT_ROOT!\Scripts_IA\Utilidades\backup_automatico.ps1" (
    echo [OK] Script de backup encontrado
) else (
    echo [ADVERTENCIA] Script de backup no encontrado
    set /a WARNING_FLAG+=1
)

if exist "!PROJECT_ROOT!\Scripts_IA\Utilidades\verificar_sistema.ps1" (
    echo [OK] Script de verificación encontrado
) else (
    echo [ADVERTENCIA] Script de verificación no encontrado
    set /a WARNING_FLAG+=1
)

echo.
echo [OK] Verificación final completada

echo.
set /p CONTINUAR="Presione S y Enter para continuar con la FASE 8... "
if /i "!CONTINUAR!" NEQ "S" (
    echo [INFO] Instalación pausada por el usuario.
    timeout /t 3 >nul
    exit /b 0
)

:: ===================================================================
:: FASE 8: RESUMEN Y FINALIZACIÓN MEJORADA
:: ===================================================================
echo.
echo [PROGRESO] FASE 8: Generando resumen final de instalación...
echo.

:: Calcular tamaño total del proyecto
echo Calculando tamaño del proyecto...
dir /s /c "!PROJECT_ROOT!" 2>nul > "%TEMP%\dirsize.txt"
for /f "tokens=3" %%a in ('type "%TEMP%\dirsize.txt" ^| find "bytes"') do (
    set "PROJECT_SIZE=%%a"
)
del "%TEMP%\dirsize.txt" 2>nul

if not defined PROJECT_SIZE (
    set "PROJECT_SIZE=Desconocido"
)

:: Obtener fecha y hora actual
for /f "tokens=1-3 delims=/ " %%a in ("%date%") do (
    set "CURRENT_DAY=%%a"
    set "CURRENT_MONTH=%%b"
    set "CURRENT_YEAR=%%c"
)
for /f "tokens=1-3 delims=:." %%a in ("%time%") do (
    set "CURRENT_HOUR=%%a"
    set "CURRENT_MINUTE=%%b"
    set "CURRENT_SECOND=%%c"
)

:: Crear archivo de resumen detallado
(
echo RESULTADO FINAL DE LA INSTALACIÓN
echo =================================
echo.
echo ?? FECHA: !CURRENT_DAY!/!CURRENT_MONTH!/!CURRENT_YEAR!
echo ? HORA: !CURRENT_HOUR!:!CURRENT_MINUTE!:!CURRENT_SECOND!
echo.
echo ?? USUARIO: %USERNAME%
echo ?? EQUIPO: %COMPUTERNAME%
echo ???  SISTEMA: %OS% !ARCH! bits
echo.
echo ?? PROYECTO: !PROJECT_ROOT!
echo ?? TAMAÑO: !PROJECT_SIZE!
echo.
echo ?? CONFIGURACIÓN:
echo   • PowerShell: !POWERSHELL_VERSION!
echo   • .NET Framework: !NET_VERSION!
echo   • Excel: !EXCEL_INSTALLED! (1=Instalado)
echo.
echo ?? ESTADÍSTICAS:
echo   • Carpetas creadas: 15 principales, 58 subcarpetas
echo   • Scripts ejecutados: !SCRIPT_SUCCESS!/!SCRIPT_TOTAL!
echo   • Archivos esenciales: !FILES_FOUND!/!FILES_TOTAL!
echo.
echo ??  ADVERTENCIAS: !WARNING_FLAG!
echo ? ERRORES: !ERROR_FLAG!
echo.
echo ? ACCESOS DIRECTOS CREADOS:
if exist "!DESKTOP_SHORTCUT!" echo   • Escritorio: Comparador Compras IA.lnk
if exist "!DESKTOP_SHORTCUT2!" echo   • Escritorio: Comparador IA - Abrir Carpeta.lnk
if exist "!START_MENU_DIR!\Comparador Compras IA.lnk" echo   • Menú Inicio: Comparador Compras IA
echo.
echo ???  HERRAMIENTAS DISPONIBLES:
echo   • backup_automatico.ps1 - Sistema de backups
echo   • verificar_sistema.ps1 - Diagnóstico del sistema
echo   • limpiar_cache.ps1 - Limpieza de caché
echo.
echo ?? ARCHIVOS IMPORTANTES:
echo   • Comparador_Compras_IA_Completo.xlsm - Excel principal
echo   • INSTRUCCIONES_PROYECTO.txt - Guía de uso
echo   • Configuraciones\config_sistema.json - Configuración
echo   • Configuraciones\resumen_configuracion.txt - Resumen
echo.
echo ?? LOGS DE INSTALACIÓN:
echo   • !LOG_FILE!
echo   • Logs\configuracion_*.log
echo.
echo =================================
) > "!PROJECT_ROOT!\RESUMEN_INSTALACION.txt"

:: Mostrar resumen en pantalla
echo ===================================================
echo         RESUMEN FINAL DE INSTALACIÓN
echo ===================================================
echo.
echo ?? ESTADO DEL SISTEMA:
if !ERROR_FLAG! equ 0 (
    if !WARNING_FLAG! equ 0 (
        echo    [? EXITOSA] Sin errores ni advertencias
    ) else (
        echo    [??  EXITOSA CON AVISOS] !WARNING_FLAG! advertencias
    )
) else (
    echo    [? CON ERRORES] !ERROR_FLAG! errores, !WARNING_FLAG! advertencias
)
echo.
echo ?? UBICACIÓN: !PROJECT_ROOT!
echo ?? TAMAÑO: !PROJECT_SIZE!
echo.
echo ??  COMPONENTES INSTALADOS:
echo   • Estructura de carpetas: 15 principales, 58 subcarpetas
echo   • Scripts de configuración: !SCRIPT_SUCCESS!/!SCRIPT_TOTAL! ejecutados
echo   • Archivos esenciales: !FILES_FOUND!/!FILES_TOTAL! verificados
echo.
echo ?? ACCESO RÁPIDO:
if exist "!DESKTOP_SHORTCUT!" (
    echo   • Abra: Comparador Compras IA.lnk (en escritorio)
) else (
    echo   • Abra: !PROJECT_ROOT!\Comparador_Compras_IA_Completo.xlsm
)
echo.
echo ???  HERRAMIENTAS INCLUIDAS:
echo   • backup_automatico.ps1 - Backups automáticos
echo   • verificar_sistema.ps1 - Diagnóstico del sistema
echo.
echo ?? DOCUMENTACIÓN:
echo   • INSTRUCCIONES_PROYECTO.txt - Guía completa
echo   • RESUMEN_INSTALACION.txt - Este resumen
echo.
echo ===================================================
echo.
echo ?? PRÓXIMOS PASOS RECOMENDADOS:
echo   1. Abra el archivo Excel desde el acceso directo
echo   2. Habilite las macros cuando se le solicite
echo   3. Complete sus datos en la hoja USUARIOS
echo   4. Revise INSTRUCCIONES_PROYECTO.txt
echo   5. Explore las funciones desde el menú "Comparador IA"
echo.
echo ??  IMPORTANTE:
echo   • Mantenga siempre copias de seguridad
echo   • Revise regularmente los logs
echo   • Ejecute verificar_sistema.ps1 si hay problemas
echo.
echo ?? SOPORTE:
echo   • Consulte la documentación incluida
echo   • Revise los logs en !PROJECT_ROOT!\Logs\
echo   • Los scripts de utilidad ayudan en diagnóstico
echo.
echo ===================================================
if !ERROR_FLAG! equ 0 (
    echo    ¡INSTALACIÓN COMPLETADA EXITOSAMENTE!
) else if !ERROR_FLAG! leq 2 (
    echo    INSTALACIÓN COMPLETADA CON ERRORES MENORES
) else (
    echo    INSTALACIÓN COMPLETADA CON ERRORES CRÍTICOS
)
echo ===================================================
echo.
echo ¡Gracias por instalar el Sistema Comparador de Compras IA v!SCRIPT_VERSION!!

echo "Presione una tecla para terminar... "
pause >nul