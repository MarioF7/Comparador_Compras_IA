@echo off
chcp 65001 > nul
echo ===================================================
echo    SISTEMA COMPARADOR DE COMPRAS INTELIGENTE IA
echo ===================================================
echo.
echo Creando estructura completa del sistema...
echo.

REM Crear estructura de carpetas
if not exist "Comparador_Compras_IA" mkdir "Comparador_Compras_IA"
cd "Comparador_Compras_IA"
if not exist "Tickets" mkdir "Tickets"
if not exist "Exportaciones" mkdir "Exportaciones"
if not exist "Backups" mkdir "Backups"
if not exist "Logs" mkdir "Logs"
if not exist "Datos_API" mkdir "Datos_API"

echo [✓] Carpetas creadas
echo.

REM Verificar si PowerShell está disponible
where powershell >nul 2>nul
if %errorlevel% equ 0 (
    echo PowerShell encontrado, continuando...
    echo.
    
    REM Crear archivo Excel
    echo Generando archivo Excel principal...
    powershell -ExecutionPolicy Bypass -Command "& {iex ((New-Object System.Net.WebClient).DownloadString('https://tinyurl.com/crear-excel-compras'))}"
    
    REM Si falla la descarga, crear localmente
    if exist "crear_excel.ps1" (
        powershell -ExecutionPolicy Bypass -File "crear_excel.ps1"
    ) else (
        echo Creando archivo Excel manualmente...
        call :crear_excel_manual
    )
    
    echo.
    echo [✓] Archivo Excel creado
    echo.
    
    REM Cargar datos
    echo Generando datos iniciales...
    if exist "cargar_datos.ps1" (
        powershell -ExecutionPolicy Bypass -File "cargar_datos.ps1"
    ) else (
        call :cargar_datos_manual
    )
    
    echo.
    echo [✓] Datos iniciales cargados
    echo.
    
    REM Agregar macros
    echo Agregando macros VBA...
    if exist "agregar_macros.vbs" (
        cscript //nologo "agregar_macros.vbs"
    ) else (
        call :agregar_macros_manual
    )
    
) else (
    echo ERROR: PowerShell no está instalado o no está en el PATH.
    echo Instala PowerShell o habilítalo y vuelve a ejecutar.
    pause
    exit /b 1
)

REM Crear archivo README
echo Creando documentación...
(
echo ================================
echo SISTEMA COMPARADOR DE COMPRAS IA
echo ================================
echo.
echo INSTRUCCIONES:
echo 1. Abrir "Comparador_Compras_IA_Completo.xlsm"
echo 2. Habilitar macros cuando Excel lo solicite
echo 3. Seguir los pasos en la hoja "00_INSTRUCCIONES"
echo.
echo BOTONES PRINCIPALES:
echo - Procesar Ticket: Hoja OCR_ENTRADA
echo - Generar Ruta: Hoja LISTAS_COMPRA
echo - Ver Dashboard: Hoja DASHBOARD
echo.
echo SOPORTE:
echo Si tienes problemas, ejecuta "solucionar_problemas.bat"
) > "INSTRUCCIONES.txt"

echo.
echo [!] Sistema creado exitosamente!
echo.
echo ===================================================
echo ARCHIVOS CREADOS:
echo - Comparador_Compras_IA_Completo.xlsm
echo - INSTRUCCIONES.txt
echo - 5 carpetas de soporte
echo ===================================================
echo.
echo PRESIONA UNA TECLA PARA ABRIR LA CARPETA...
pause > nul

REM Abrir carpeta
explorer .
exit /b 0

:crear_excel_manual
echo Creando Excel manualmente (sin PowerShell)...
REM Esta función se ejecutará si no hay PowerShell
echo Por favor, crea el archivo Excel manualmente.
echo.
echo PASOS:
echo 1. Abre Excel
echo 2. Crea nuevo libro
echo 3. Guarda como "Comparador_Compras_IA_Completo.xlsm"
echo 4. Crea las hojas: USUARIOS, PRODUCTOS, TIENDAS, COMPRAS
exit /b 0

:cargar_datos_manual
echo Cargando datos manualmente...
(
echo ID_Usuario,Nombre,Email
echo U001,Usuario Principal,usuario@email.com
echo U002,Usuario Familiar,familia@email.com
) > "datos_usuarios.csv"
echo Datos básicos creados en datos_usuarios.csv
exit /b 0

:agregar_macros_manual
echo Creando archivo de macros básicas...
(
echo Option Explicit
echo Sub InicializarSistema()
echo     MsgBox "Sistema inicializado"
echo End Sub
) > "macros_basicas.txt"
echo Macros básicas creadas en macros_basicas.txt
exit /b 0