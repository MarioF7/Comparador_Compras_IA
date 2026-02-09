@echo off
echo ==========================================
echo   SOLUCIONADOR DE PROBLEMAS - COMPRAS IA
echo ==========================================
echo.
echo Selecciona el problema:
echo.
echo 1. PowerShell no ejecuta scripts
echo 2. Excel no permite macros
echo 3. Archivo no se crea
echo 4. Error al abrir el archivo
echo 5. Restaurar sistema desde cero
echo.
set /p opcion="Elige opción (1-5): "

if "%opcion%"=="1" goto problema1
if "%opcion%"=="2" goto problema2
if "%opcion%"=="3" goto problema3
if "%opcion%"=="4" goto problema4
if "%opcion%"=="5" goto problema5

:problema1
echo.
echo SOLUCIÓN 1: Permitir ejecución de PowerShell
echo.
echo Ejecutando comando para permitir scripts...
powershell -Command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force"
echo.
echo ✅ Hecho. Intenta ejecutar crear_sistema.bat de nuevo.
pause
exit

:problema2
echo.
echo SOLUCIÓN 2: Habilitar macros en Excel
echo.
echo Sigue estos pasos:
echo 1. Abre Excel
echo 2. Ve a Archivo -> Opciones -> Centro de confianza
echo 3. Configuración del Centro de confianza
echo 4. Configuración de macros -> Habilitar todas las macros
echo 5. Aceptar y reiniciar Excel
echo.
pause
exit

:problema3
echo.
echo SOLUCIÓN 3: Archivo no se crea
echo.
echo Eliminando y recreando archivos...
del /f /q "Comparador_Compras_IA_Completo.xlsm" 2>nul
del /f /q "crear_excel.ps1" 2>nul
echo Recreando archivos...
copy crear_excel.ps1.bak crear_excel.ps1 2>nul
echo.
echo ✅ Intenta ejecutar de nuevo.
pause
exit

:problema4
echo.
echo SOLUCIÓN 4: Error al abrir archivo
echo.
echo Verificando archivo...
if exist "Comparador_Compras_IA_Completo.xlsm" (
    echo El archivo existe. Probablemente está dañado.
    echo.
    echo Creando copia nueva...
    copy "Comparador_Compras_IA_Completo.xlsm" "Comparador_Compras_IA_Completo_BAK.xlsm"
    del "Comparador_Compras_IA_Completo.xlsm"
    call crear_excel.ps1
) else (
    echo El archivo no existe. Creando nuevo...
    call crear_excel.ps1
)
pause
exit

:problema5
echo.
echo SOLUCIÓN 5: Restaurar sistema desde cero
echo.
set /p confirm="¿Estás seguro? Se eliminarán todos los datos. (S/N): "
if /i "%confirm%" neq "S" exit

echo Eliminando archivos anteriores...
del /f /q *.xlsm 2>nul
del /f /q *.csv 2>nul
del /f /q *.txt 2>nul
rmdir /s /q Tickets 2>nul
rmdir /s /q Backups 2>nul
rmdir /s /q Exportaciones 2>nul
rmdir /s /q Logs 2>nul
rmdir /s /q Datos_API 2>nul

echo Creando sistema nuevo...
call crear_sistema.bat
pause
exit