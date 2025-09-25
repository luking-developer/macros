Sub AgregarColumnaFechaNumero
    On Error GoTo ErrHandler

    Dim oDoc As Object
    Dim oSheet As Object
    Dim oCursor As Object
    Dim oRange As Object
    Dim oCell As Object

    Dim ultimaCol As Long
    Dim ultimaFila As Long
    Dim i As Long
    Dim fechaColIndex As Long
    Dim newColIndex As Long
    Dim fechaColLetter As String
    Dim dateColText As String
    Dim newColName As String

    oDoc = ThisComponent
    oSheet = oDoc.CurrentController.ActiveSheet

    ' Expected header text
    dateColText = "FECHA_ALTA"
    newColName = "FECHA"

    ' Detect last used column and row
    oCursor = oSheet.createCursor()
    oCursor.gotoEndOfUsedArea(True)
    ultimaCol = oCursor.RangeAddress.EndColumn
    ultimaFila = oCursor.RangeAddress.EndRow

    ' Find the index of the column with header = dateColText
    fechaColIndex = -1
    For i = 0 To ultimaCol
        oCell = oSheet.getCellByPosition(i, 0)
        If Trim(UCase(oCell.String)) = UCase(dateColText) Then
            fechaColIndex = i
            Exit For
        End If
    Next i

    If fechaColIndex = -1 Then
        MsgBox "No se encontró la columna '" & dateColText & "' en los encabezados.", 48, "AgregarColumnaFechaNumero"
        Exit Sub
    End If

    ' Insert new column
    newColIndex = ultimaCol + 1
    oSheet.Columns.insertByIndex(newColIndex, 1)

    ' Get column letter for formulas
    fechaColLetter = ColumnIndexToLetters(fechaColIndex)

    ' Fill new column with =VALUE(<fecha_alta_cell>)
    For i = 1 To ultimaFila
        oCell = oSheet.getCellByPosition(newColIndex, i)
        oCell.Formula = "=VALUE(" & fechaColLetter & (i + 1) & ")"
    Next i

    ' Apply a built-in date format (index 36 = DD/MM/YY)
    oRange = oSheet.getCellRangeByPosition(newColIndex, 1, newColIndex, ultimaFila)
    oRange.NumberFormat = 36

    ' Set new header
    oSheet.getCellByPosition(newColIndex, 0).String = newColName

	OrdenarPorFecha(newColName)

    Exit Sub

ErrHandler:
    MsgBox "Error BASIC: " & Err & vbCrLf & Error$ & vbCrLf & "Line: " & Erl, 16, "AgregarColumnaFechaNumero"
End Sub

' Convert 0-based column index to Excel-like letters
Function ColumnIndexToLetters(colIndex As Long) As String
    Dim n As Long
    Dim s As String
    n = colIndex
    s = ""
    Do
        s = Chr(65 + (n Mod 26)) & s
        n = (n \ 26) - 1
    Loop While n >= 0
    ColumnIndexToLetters = s
End Function

Sub OrdenarPorFecha(nombreColumna As String)
    On Error GoTo ErrHandler
    
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oRange As Object
    Dim oCell As Object
    Dim oSortFields(0) As New com.sun.star.util.SortField
    Dim oSortDesc(0) As New com.sun.star.beans.PropertyValue
    
    Dim columnaIndice As Long
    Dim i As Long
    
    oDoc = ThisComponent
    oSheet = oDoc.CurrentController.ActiveSheet
    
    ' Obtener el rango de datos que se va a ordenar
   	Dim oCursor As Object
	oCursor = oSheet.createCursor()
	oCursor.gotoEndOfUsedArea(False) ' Mueve el cursor al final de los datos
	oCursor.gotoStartOfUsedArea(True) ' Extiende la selección desde el final hasta el inicio
	oRange = oCursor
    
    ' Encontrar el índice de la columna de ordenamiento
    columnaIndice = -1
    For i = 0 To oRange.Columns.Count - 1
        ' Acceder a la celda por su posición relativa en el rango
        oCell = oRange.getCellByPosition(i, 0)
        
        If UCase(oCell.String) = UCase(nombreColumna) Then
            columnaIndice = i
            Exit For
        End If
    Next i
    
    If columnaIndice = -1 Then
        MsgBox "No se encontró la columna '" & nombreColumna & "' para ordenar.", 48, "Ordenar por Fecha"
        Exit Sub
    End If
    
    ' Configurar los campos de ordenamiento
    With oSortFields(0)
        .Field = columnaIndice ' El índice de la columna en el rango
        .SortAscending = False ' Orden descendente
    End With
    
    ' Configurar las propiedades del ordenamiento
    With oSortDesc(0)
        .Name = "SortFields"
        .Value = oSortFields()
    End With
    
    ' Realizar el ordenamiento en el rango de datos
    oRange.Sort(oSortDesc())
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error BASIC en OrdenarPorFecha: " & Err & vbCrLf & Error$ & vbCrLf & "Line: " & Erl, 16, "Ordenar por Fecha"
End Sub
