Sub DRFiltrarCancelled()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col1 As ListColumn
    Dim col2 As ListColumn
    Dim filtro As String
    Dim i As Long
    Dim copyRows As Range
    Dim newRow As Range
    Dim newWs As Worksheet
    Dim foundTable As Boolean
    
    ' Establecer la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("Data_2")
    
    ' Buscar la tabla
    For Each tbl In ws.ListObjects
        If tbl.Name = "Clean_pop_dr" Then
            foundTable = True
            Exit For
        End If
    Next tbl
    
    ' Si no se encontró la tabla, mostrar un mensaje y salir
    If Not foundTable Then
        MsgBox "No se encontró la tabla 'TABLAF' en la hoja 'Data_2'.", vbExclamation
        Exit Sub
    End If
    
    ' Definir las columnas relevantes
    Set col1 = tbl.ListColumns("Risk_Rating_Outcome")
    Set col2 = tbl.ListColumns("Offboarding_Repository_Outcome")
    
    ' Establecer el criterio de filtro
    filtro = "*cancelled*"
    
    ' Limpiar el filtro actual
    If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
    
    ' Crear una nueva hoja para copiar las filas filtradas
    Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newWs.Name = "DR_Filtrado_Cancelled"
    
    ' Iterar a través de las filas y copiar las que cumplen con el criterio
    For i = 1 To tbl.ListRows.Count
        If LCase(tbl.DataBodyRange(i, col1.Index).Value) Like LCase(filtro) Or _
           LCase(tbl.DataBodyRange(i, col2.Index).Value) Like LCase(filtro) Then
            If copyRows Is Nothing Then
                Set copyRows = tbl.ListRows(i).Range
            Else
                Set copyRows = Union(copyRows, tbl.ListRows(i).Range)
            End If
        End If
    Next i
    
    ' Verificar si la primera fila también cumple con el criterio
    If LCase(tbl.HeaderRowRange.Cells(1, col1.Index).Value) Like LCase(filtro) Or _
       LCase(tbl.HeaderRowRange.Cells(1, col2.Index).Value) Like LCase(filtro) Then
        If copyRows Is Nothing Then
            Set copyRows = tbl.HeaderRowRange
        Else
            Set copyRows = Union(copyRows, tbl.HeaderRowRange)
        End If
    End If
    
    ' Copiar las filas en la nueva hoja
    If Not copyRows Is Nothing Then
        copyRows.Copy newWs.Range("A1")
        Application.CutCopyMode = False
    End If
    
    ' Ajustar el ancho de las columnas en la nueva hoja
    newWs.Columns.AutoFit
    
    ' Mensaje de finalización
    MsgBox "Filtro aplicado. Las filas filtradas se han copiado en una nueva hoja llamada 'DR_Filtrado_Cancelled'."
End Sub