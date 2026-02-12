Attribute VB_Name = "modComparador"
Option Explicit

' =============================================
' ALGORITMO DE COMPARACIÓN INTELIGENTE
' =============================================

' -------------------------------------------------
' Calcular puntuación de una tienda para un producto
' -------------------------------------------------
Function CalificarTienda(ByVal dblPrecio As Double, _
                         ByVal dblDistancia As Double, _
                         ByVal dblValoracion As Double, _
                         Optional ByVal dblPrecioMin As Double = 0, _
                         Optional ByVal dblPrecioMax As Double = 0, _
                         Optional ByVal dblDistMin As Double = 0, _
                         Optional ByVal dblDistMax As Double = 0) As Double
    
    Dim precioScore As Double, distScore As Double, valScore As Double
    Dim w_precio As Double, w_dist As Double, w_val As Double
    
    ' Cargar pesos desde configuración (se puede personalizar por usuario)
    w_precio = DEFAULT_WEIGHT_PRICE
    w_dist = DEFAULT_WEIGHT_DISTANCE
    w_val = DEFAULT_WEIGHT_RATING
    
    ' Normalizar precio (0 = caro, 1 = barato)
    If dblPrecioMax > dblPrecioMin Then
        precioScore = (dblPrecioMax - dblPrecio) / (dblPrecioMax - dblPrecioMin)
    Else
        precioScore = 1
    End If
    
    ' Normalizar distancia (0 = lejos, 1 = cerca)
    If dblDistMax > dblDistMin Then
        distScore = (dblDistMax - dblDistancia) / (dblDistMax - dblDistMin)
    Else
        distScore = 1
    End If
    
    ' Normalizar valoración (sobre 5)
    valScore = dblValoracion / 5
    
    ' Puntuación final ponderada
    CalificarTienda = (precioScore * w_precio) + (distScore * w_dist) + (valScore * w_val)
End Function

' -------------------------------------------------
' Generar comparativa para un producto específico
' -------------------------------------------------
Sub GenerarComparativaProducto(ByVal sProductoID As String, _
                               Optional ByVal sUsuarioID As String = "USR001")
    Dim wsPrecios As Worksheet, wsTiendas As Worksheet, wsComparativa As Worksheet
    Dim wsProductos As Worksheet
    
    Set wsPrecios = ThisWorkbook.Worksheets(SHEET_PRECIOS)
    Set wsTiendas = ThisWorkbook.Worksheets(SHEET_TIENDAS)
    Set wsProductos = ThisWorkbook.Worksheets(SHEET_PRODUCTOS)
    Set wsComparativa = ThisWorkbook.Worksheets(SHEET_COMPARATIVA)
    
    ' Obtener columnas dinámicamente
    Dim colPrecioProdID As Long, colPrecioTiendaID As Long, colPrecioValor As Long
    Dim colTiendaID As Long, colTiendaLat As Long, colTiendaLon As Long
    Dim colTiendaVal As Long, colTiendaDist As Long
    Dim colProdNombre As Long
    
    colPrecioProdID = ObtenerColumna(wsPrecios, "ProductID")
    colPrecioTiendaID = ObtenerColumna(wsPrecios, "StoreID")
    colPrecioValor = ObtenerColumna(wsPrecios, "Precio_Unitario")
    
    colTiendaID = ObtenerColumna(wsTiendas, "StoreID")
    colTiendaLat = ObtenerColumna(wsTiendas, "Coord_Lat")
    colTiendaLon = ObtenerColumna(wsTiendas, "Coord_Lon")
    colTiendaVal = ObtenerColumna(wsTiendas, "Valoracion_Media")
    colTiendaDist = ObtenerColumna(wsTiendas, "Distancia_Usuario")
    
    colProdNombre = ObtenerColumna(wsProductos, "Nombre")
    
    ' Limpiar comparativas anteriores de este producto (opcional)
    ' Aquí podrías filtrar
    
    ' Recopilar precios del producto en todas las tiendas
    Dim i As Long, ultFilaPrecios As Long
    Dim tiendaID As String, precio As Double
    Dim tiendaNombre As String, distancia As Double, valoracion As Double
    Dim precioMin As Double, precioMax As Double
    Dim distMin As Double, distMax As Double
    
    ultFilaPrecios = wsPrecios.Cells(wsPrecios.Rows.Count, colPrecioProdID).End(xlUp).Row
    
    ' Calcular mínimos y máximos para normalización
    precioMin = 9999999: precioMax = 0
    distMin = 9999999: distMax = 0
    
    For i = 2 To ultFilaPrecios
        If wsPrecios.Cells(i, colPrecioProdID).Value = sProductoID Then
            precio = wsPrecios.Cells(i, colPrecioValor).Value
            If precio < precioMin Then precioMin = precio
            If precio > precioMax Then precioMax = precio
        End If
    Next i
    
    ' Obtener distancias de todas las tiendas
    Dim j As Long, ultFilaTiendas As Long
    ultFilaTiendas = wsTiendas.Cells(wsTiendas.Rows.Count, colTiendaID).End(xlUp).Row
    For j = 2 To ultFilaTiendas
        distancia = wsTiendas.Cells(j, colTiendaDist).Value
        If distancia < distMin Then distMin = distancia
        If distancia > distMax Then distMax = distancia
    Next j
    
    ' Generar filas en COMPARATIVA
    Dim nuevaFila As Long
    nuevaFila = wsComparativa.Cells(wsComparativa.Rows.Count, 1).End(xlUp).Row + 1
    
    For i = 2 To ultFilaPrecios
        If wsPrecios.Cells(i, colPrecioProdID).Value = sProductoID Then
            tiendaID = wsPrecios.Cells(i, colPrecioTiendaID).Value
            precio = wsPrecios.Cells(i, colPrecioValor).Value
            
            ' Buscar datos de la tienda
            For j = 2 To ultFilaTiendas
                If wsTiendas.Cells(j, colTiendaID).Value = tiendaID Then
                    tiendaNombre = wsTiendas.Cells(j, ObtenerColumna(wsTiendas, "Nombre_Tienda")).Value
                    distancia = wsTiendas.Cells(j, colTiendaDist).Value
                    valoracion = wsTiendas.Cells(j, colTiendaVal).Value
                    Exit For
                End If
            Next j
            
            ' Calcular puntuación
            Dim puntuacion As Double
            puntuacion = CalificarTienda(precio, distancia, valoracion, precioMin, precioMax, distMin, distMax)
            
            ' Escribir fila
            wsComparativa.Cells(nuevaFila, ObtenerColumna(wsComparativa, "ComparativaID")).Value = GenerarNuevoID(SHEET_COMPARATIVA, ID_PREFIX_COMPARATIVA)
            wsComparativa.Cells(nuevaFila, ObtenerColumna(wsComparativa, "UserID")).Value = sUsuarioID
            wsComparativa.Cells(nuevaFila, ObtenerColumna(wsComparativa, "ProductID")).Value = sProductoID
            wsComparativa.Cells(nuevaFila, ObtenerColumna(wsComparativa, "Tienda_Mejor_Precio")).Value = tiendaNombre
            wsComparativa.Cells(nuevaFila, ObtenerColumna(wsComparativa, "Mejor_Precio")).Value = precio
            wsComparativa.Cells(nuevaFila, ObtenerColumna(wsComparativa, "Distancia_Mejor")).Value = distancia
            wsComparativa.Cells(nuevaFila, ObtenerColumna(wsComparativa, "Puntuación_Global")).Value = puntuacion
            wsComparativa.Cells(nuevaFila, ObtenerColumna(wsComparativa, "Fecha_Comparación")).Value = Now
            
            nuevaFila = nuevaFila + 1
        End If
    Next i
    
    MostrarExito "Comparativa generada para el producto."
End Sub

' -------------------------------------------------
' Comparar lista de productos (optimización de ruta)
' -------------------------------------------------
Sub GenerarRutaOptima(ByVal sUsuarioID As String, ParamArray productos() As Variant)
    ' Esta función requiere implementación de algoritmo de ruta
    ' Por simplicidad en esta fase, se deja esqueleto
    MsgBox "Función de ruta óptima disponible en Fase 3", vbInformation
End Sub