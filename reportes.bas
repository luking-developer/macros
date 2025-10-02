Sub AgregarFechaSiVacio()
    Dim hoja As Object
    Dim fila As Integer
    Dim celdaA As Object
    Dim celdaN As Object
    Dim ultimaFila As Integer

    hoja = ThisComponent.Sheets(0) ' Usa la primera hoja

    ' Comienza desde la fila 2
    fila = 2

    ' Recorre las filas hasta encontrar una fila vacía en la columna A
    Do While hoja.getCellByPosition(1, fila).String <> "" ' Mientras haya contenido en la columna A
        Set celdaA = hoja.getCellByPosition(0, fila) ' Columna A (índice 0)
        Set celdaN = hoja.getCellByPosition(13, fila) ' Columna N (índice 13)

        ' Si la columna A tiene contenido y la columna N está vacía
        If celdaA.String <> "" And celdaN.String = "" Then
            celdaN.Value = Date
            celdaN.NumberFormat = 37 ' Formato de fecha
        End If

        fila = fila + 1 ' Avanzar a la siguiente fila
    Loop
End Sub
