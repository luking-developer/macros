Sub OrdenarPorFechaAlta()
    On Error GoTo ErrHandler
    Dim oDoc As Object : oDoc = StarDesktop.CurrentComponent
    Dim oSheet As Object : oSheet = oDoc.CurrentController.ActiveSheet
    Dim oCursor As Object : oCursor = oSheet.createCursor()
    
    oCursor.gotoEndOfUsedArea(False)
    Dim lastRow As Long : lastRow = oCursor.RangeAddress.EndRow
    
    ' 1. Localizar Columnas
    Dim i As Long, colAlta As Long : colAlta = -1
    Dim colFecha As Long : colFecha = -1
    For i = 0 To oCursor.RangeAddress.EndColumn
        Dim sH As String : sH = UCase(Trim(oSheet.getCellByPosition(i, 0).String))
        If sH = "FECHA_ALTA" Then colAlta = i
        If sH = "FECHA" Then colFecha = i
    Next i
    
    If colAlta = -1 Then Exit Sub ' No perder tiempo si no hay origen
       
    ' 3. Conversión de Datos
    For i = 1 To lastRow
        Dim sRaw As String : sRaw = LCase(Trim(oSheet.getCellByPosition(colAlta, i).String))
        Dim oCeldaDest As Object : oCeldaDest = oSheet.getCellByPosition(colAlta, i)
        
        If sRaw <> "" Then
            sRaw = Join(Split(sRaw, "sept."), "sep.") 
            On Error Resume Next
            oCeldaDest.Value = CDate(sRaw)
            On Error GoTo ErrHandler
        End If
    Next i

    ' 4. FORMATEO DE FUERZA BRUTA (Usando el Dispatcher)
    ' Seleccionamos toda la columna de datos
    Dim oRange As Object
    oRange = oSheet.getCellRangeByPosition(colAlta, 1, colAlta, lastRow)
    oDoc.CurrentController.select(oRange)
    
    Dim document   as object : document   = oDoc.CurrentController.Frame
    Dim dispatcher as object : dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    Dim args(0) as new com.sun.star.beans.PropertyValue
    args(0).Name = "NumberFormatValue"
    
    ' El valor 36 es el estándar interno para DD/MM/YYYY
    args(0).Value = 36
    dispatcher.executeDispatch(document, ".uno:NumberFormatValue", "", 0, args())

    ' 5. ORDENACIÓN
    Dim oSortCur As Object : oSortCur = oSheet.createCursor()
    oSortCur.gotoEndOfUsedArea(True)
    
    Dim oSortFields(0) As New com.sun.star.table.TableSortField
    oSortFields(0).Field = colAlta
    oSortFields(0).IsAscending = False

    Dim oSortProps(1) As New com.sun.star.beans.PropertyValue
    oSortProps(0).Name = "SortFields"
    oSortProps(0).Value = oSortFields()
    oSortProps(1).Name = "ContainsHeader"
    oSortProps(1).Value = True

    oSortCur.Sort(oSortProps())
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Error$ & " en línea " & Erl
End Sub
