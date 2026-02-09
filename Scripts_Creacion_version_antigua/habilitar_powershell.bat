@echo off
echo Habilitando PowerShell para ejecutar scripts...
echo.
powershell -Command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force"
echo.
echo ✅ PowerShell configurado correctamente.
echo.
pause