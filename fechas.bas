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
    fechaColLetter = UtilidadesFecha.ColumnIndexToLetters(fechaColIndex)

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


