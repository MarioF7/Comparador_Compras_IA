@echo off
echo Sistema Simple de Comparador de Compras
echo.

REM Crear Excel básico con VBScript (funciona en TODOS Windows)
echo Creando Excel básico...
(
echo Set objExcel = CreateObject("Excel.Application")
echo objExcel.Visible = False
echo objExcel.DisplayAlerts = False
echo Set objWorkbook = objExcel.Workbooks.Add()
echo 
echo ' Crear hojas
echo objWorkbook.Sheets.Add ,,5
echo objWorkbook.Sheets(1).Name = "USUARIOS"
echo objWorkbook.Sheets(2).Name = "PRODUCTOS" 
echo objWorkbook.Sheets(3).Name = "TIENDAS"
echo objWorkbook.Sheets(4).Name = "COMPRAS"
echo objWorkbook.Sheets(5).Name = "INSTRUCCIONES"
echo 
echo ' Datos básicos en USUARIOS
echo Set ws = objWorkbook.Sheets("USUARIOS")
echo ws.Cells(1,1) = "ID_Usuario"
echo ws.Cells(1,2) = "Nombre"
echo ws.Cells(1,3) = "Email"
echo ws.Cells(2,1) = "U001"
echo ws.Cells(2,2) = "Tu Nombre"
echo ws.Cells(2,3) = "tu@email.com"
echo 
echo ' Guardar
echo objWorkbook.SaveAs "Comparador_Simple.xlsx"
echo objWorkbook.Close
echo objExcel.Quit
echo MsgBox "Sistema creado!"
) > crear_excel.vbs

cscript //nologo crear_excel.vbs

echo.
echo [✓] Sistema simple creado: Comparador_Simple.xlsx
echo.
echo Ahora tienes un sistema básico para empezar.
echo.
pause