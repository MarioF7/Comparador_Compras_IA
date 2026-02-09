@echo off
echo Habilitando PowerShell para ejecutar scripts...
echo.
powershell -Command "Set-ExecutionPolicy -ExecutionPolicy Restricted -Scope CurrentUser"
echo.
echo ✅ PowerShell configurado correctamente.
echo.
pause