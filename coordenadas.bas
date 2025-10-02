Sub GenerarNubeDePuntos()

    ' PASO 1: Parámetros y Variables iniciales (Nuevas variables resaltadas)
    ' ---------------------------------------
    
    ' Parámetros que puedes cambiar
    Dim NombreHoja As String
    NombreHoja = "EPE 057.1"  ' Nombre de la hoja/pestaña a procesar
    
    Dim ColumnaProcesar As String
    ColumnaProcesar = "D"     ' Letra de la columna con los datos a modificar
    
    Dim ColumnaDistritos As String
    ColumnaDistritos = "B"    ' Letra de la columna donde buscar el distrito
    
    Dim ColumnaUsuarios As String
    ColumnaUsuarios = "C"    ' Letra de la columna donde buscar el usuario
    
    Dim Distritos() As String
    ' Lista de distritos a buscar.
    Distritos = Array("Rafaela", "Bella Italia") 
    
    Dim RutaArchivoTexto As String
    RutaArchivoTexto = "C:\Users\operadorgis\Desktop\coordenadas.txt" ' Ruta para guardar el archivo de coordenadas**
    
    
    ' Variables internas
    Dim Documento As Object
    Dim Hojas As Object
    Dim Hoja As Object
    Dim CeldaProcesar As Object
    Dim CeldaDistrito As Object
    Dim CeldaUsuario As Object
    Dim IndiceColumnaProcesar As Long
    Dim IndiceColumnaDistritos As Long
    Dim i As Long
    Dim j As Long ' Para iterar sobre la lista de distritos
    Dim UltimaFila As Long
    Dim TextoOriginal As String
    Dim TextoProcesado As String
    Dim ArchivoNum As Integer
    Dim ContenidoSalida As String
    Dim DistritoCelda As String
    Dim CoincideDistrito As Boolean
    
    ' Obtener el documento actual de Calc
    Documento = ThisComponent
    Hojas = Documento.getSheets()
    
    
    ' PASO 2: Verificar la existencia de la Hoja y seleccionarla
    ' ---------------------------------------------------------
    
    If Hojas.hasByName(NombreHoja) Then
        Hoja = Hojas.getByName(NombreHoja)
        
        ' **MODIFICACIÓN:** Seleccionar/activar la hoja
        Documento.CurrentController.setActiveSheet(Hoja)
    Else
        MsgBox "ERROR: La hoja """ & NombreHoja & """ no existe en este libro.", 16, "Error de Hoja"
        Exit Sub ' Termina la ejecución si la hoja no existe
    End If
    
    
    ' PASO 3: Obtener el rango de datos y procesar registros con filtro
    ' ---------------------------------------------------------------
    
    ' Obtener los índices de columna (0=A, 1=B, etc.) para acceder a ellas
    IndiceColumnaProcesar = Hoja.getColumns().getByName(ColumnaProcesar).getRangeAddress().StartColumn
    IndiceColumnaDistritos = Hoja.getColumns().getByName(ColumnaDistritos).getRangeAddress().StartColumn
    IndiceColumnaUsuarios = Hoja.getColumns().getByName(ColumnaUsuarios).getRangeAddress().StartColumn
    
    ' Encontrar la última fila usada en la hoja
    Dim Cursor As Object
    Cursor = Hoja.createCursor()
    Cursor.gotoStartOfUsedArea(False)
    Cursor.gotoEndOfUsedArea(True)
    UltimaFila = Cursor.getRangeAddress().EndRow
    
    ContenidoSalida = "" ' Inicializa la cadena para el archivo de salida
    
    ' Recorrer las filas
    For i = 0 To UltimaFila
        
        ' 3.1. Obtener el valor de la columna de filtro (Distrito)
        CeldaDistrito = Hoja.getCellByPosition(IndiceColumnaDistritos, i)
        
        ' Limpiar y convertir a mayúsculas el texto de la celda para una comparación robusta
        DistritoCelda = Trim(UCase(CeldaDistrito.getString())) 
        
        CoincideDistrito = False
        ' 3.2. Comprobar si el distrito de la celda coincide con alguno de la lista
        For j = LBound(Distritos) To UBound(Distritos)
            If DistritoCelda = Trim(UCase(Distritos(j))) Then ' Compara también en mayúsculas y limpiando espacios
                CoincideDistrito = True
                Exit For ' Salir del bucle interno tan pronto como encuentre una coincidencia
            End If
        Next j
        
        ' 3.3. Si el distrito coincide, continuar con el procesamiento de la ColumnaProcesar
        If CoincideDistrito Then
            
            CeldaProcesar = Hoja.getCellByPosition(IndiceColumnaProcesar, i)
            CeldaUsuario = Hoja.getCellByPosition(IndiceColumnaUsuarios, i)
            
            If Not IsEmpty(CeldaProcesar.getString()) Then
                TextoOriginal = CeldaProcesar.getString()
                
                ' Reemplazos solicitados
                ' 1. Reemplazar comas "," por una tabulación Chr$(9)
                TextoProcesado = Replace(TextoOriginal, ",", Chr$(9))
                
                ' 2. Reemplazar puntos "." por comas ","
                TextoProcesado = Replace(TextoProcesado, ".", ",")
                
                ' Acumular el texto procesado y añadir un retorno de carro Chr$(10) para separar registros
				ContenidoSalida = ContenidoSalida & TextoProcesado & " ;" & CeldaUsuario.getString() & Chr$(10)
            End If
            
        End If
        
    Next i
    
    ' Verificar si se encontraron registros
    If ContenidoSalida = "" Then
        MsgBox "Proceso finalizado. No se encontraron registros que coincidan con los distritos especificados.", 48, "Proceso Incompleto"
        Exit Sub
    End If
    
    
    ' PASO 4: Volcar los registros en un archivo de texto
    ' -------------------------------------------------
    
    ArchivoNum = FreeFile()
    Open RutaArchivoTexto For Output As #ArchivoNum
    Print #ArchivoNum, ContenidoSalida
    Close #ArchivoNum
    
    
    ' PASO 5: Mensaje de éxito
    ' ------------------------
    
    MsgBox "Procesamiento completado con éxito." & Chr$(10) & "Registros filtrados guardados en:" & Chr$(10) & RutaArchivoTexto, 64, "Macro Finalizada"

End Sub
