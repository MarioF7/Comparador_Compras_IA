Option Explicit

' ===================================================
' MÓDULO MATEMÁTICO - FUNCIONES DE CÁLCULO
' ===================================================

Public Function CalcularDistanciaHaversine(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double
    Dim dLat As Double, dLon As Double
    Dim a As Double, c As Double

    ' Convertir grados a radianes
    dLat = GradosARadianes(lat2 - lat1)
    dLon = GradosARadianes(lon2 - lon1)

    ' Fórmula de Haversine
    a = Sin(dLat / 2) * Sin(dLat / 2) + _
        Cos(GradosARadianes(lat1)) * Cos(GradosARadianes(lat2)) * _
        Sin(dLon / 2) * Sin(dLon / 2)

    c = 2 * WorksheetFunction.Atan2(Sqr(a), Sqr(1 - a))

    ' Distancia en kilómetros
    CalcularDistanciaHaversine = EARTH_RADIUS_KM * c
End Function

Public Function GradosARadianes(grados As Double) As Double
    GradosARadianes = grados * PI / 180
End Function

Public Function CalcularPrecioPorUnidad(precioTotal As Double, cantidad As Double, unidadOrigen As String) As Double
    Dim factorConversion As Double

    Select Case LCase(unidadOrigen)
        Case "kg", "litro", "unidad"
            factorConversion = 1
        Case "g"
            factorConversion = 1000
        Case "ml"
            factorConversion = 1000
        Case "mg"
            factorConversion = 1000000
        Case Else
            factorConversion = 1
    End Select

    If cantidad > 0 Then
        CalcularPrecioPorUnidad = (precioTotal / cantidad) * factorConversion
    Else
        CalcularPrecioPorUnidad = 0
    End If
End Function

Public Function CalcularAhorroPorcentual(precioOriginal As Double, precioOferta As Double) As Double
    If precioOriginal > 0 Then
        CalcularAhorroPorcentual = ((precioOriginal - precioOferta) / precioOriginal) * 100
    Else
        CalcularAhorroPorcentual = 0
    End If
End Function

Public Function ValidarEmail(email As String) As Boolean
    Dim regex As Object

    On Error GoTo ErrorHandler

    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
        .IgnoreCase = True
        .Global = False
    End With

    ValidarEmail = regex.Test(email)

    Exit Function

ErrorHandler:
    ValidarEmail = False
End Function

Public Function FormatearMoneda(valor As Double, Optional moneda As String = "EUR") As String
    Select Case UCase(moneda)
        Case "EUR", "€"
            FormatearMoneda = Format(valor, "0.00€")
        Case "USD", "$"
            FormatearMoneda = Format(valor, "$0.00")
        Case "GBP", "£"
            FormatearMoneda = Format(valor, "£0.00")
        Case Else
            FormatearMoneda = Format(valor, "0.00")
    End Select
End Function