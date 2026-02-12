Attribute VB_Name = "modUtilidades"
Option Explicit

' =============================================
' FUNCIONES AUXILIARES PARA TODO EL SISTEMA
' =============================================

' -------------------------------------------------
' Obtener el siguiente ID disponible en una columna
' -------------------------------------------------
Function GenerarNuevoID(ByVal sHoja As String, ByVal sPrefijo As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sHoja)
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If ultimaFila = 1 Then
        GenerarNuevoID = sPrefijo & "001"
        Exit Function
    End If
    
    Dim ultimoID As String
    ultimoID = ws.Cells(ultimaFila, 1).Value
    Dim numPart As Long
    numPart = Val(Mid(ultimoID, Len(sPrefijo) + 1)) + 1
    GenerarNuevoID = sPrefijo & Format(numPart, "000")
End Function

' -------------------------------------------------
' Obtener índice de columna por nombre de encabezado
' -------------------------------------------------
Function ObtenerColumna(ByVal ws As Worksheet, ByVal sHeader As String) As Long
    Dim i As Long
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If UCase(Trim(ws.Cells(1, i).Value)) = UCase(Trim(sHeader)) Then
            ObtenerColumna = i
            Exit Function
        End If
    Next i
    ObtenerColumna = 0
End Function

' -------------------------------------------------
' Validar que una celda contiene un número positivo
' -------------------------------------------------
Function EsNumeroPositivo(ByVal valor As Variant) As Boolean
    If IsNumeric(valor) Then
        EsNumeroPositivo = (valor > 0)
    Else
        EsNumeroPositivo = False
    End If
End Function

' -------------------------------------------------
' Mostrar mensaje de error con formato
' -------------------------------------------------
Sub MostrarError(ByVal sMensaje As String)
    MsgBox "❌ " & sMensaje, vbCritical, APP_NAME
End Sub

' -------------------------------------------------
' Mostrar mensaje de éxito
' -------------------------------------------------
Sub MostrarExito(ByVal sMensaje As String)
    MsgBox "✅ " & sMensaje, vbInformation, APP_NAME
End Sub

' -------------------------------------------------
' Limpiar contenido de un rango (excepto encabezados)
' -------------------------------------------------
Sub LimpiarDatosHoja(ByVal ws As Worksheet)
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultimaFila > 1 Then
        ws.Rows("2:" & ultimaFila).Delete Shift:=xlUp
    End If
End Sub

' -------------------------------------------------
' Formatear moneda en euros
' -------------------------------------------------
Function FormatearMoneda(ByVal valor As Double) As String
    FormatearMoneda = Format(valor, "#,##0.00 €")
End Function

' -------------------------------------------------
' Registrar acción en hoja de log (si existe)
' -------------------------------------------------
Sub RegistrarLog(ByVal sAccion As String, ByVal sDetalle As String)
    On Error Resume Next
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Worksheets("LOG")
    If wsLog Is Nothing Then Exit Sub
    
    Dim ultFila As Long
    ultFila = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    wsLog.Cells(ultFila, 1).Value = Now
    wsLog.Cells(ultFila, 2).Value = sAccion
    wsLog.Cells(ultFila, 3).Value = sDetalle
    wsLog.Cells(ultFila, 4).Value = Environ("Username")
End Sub