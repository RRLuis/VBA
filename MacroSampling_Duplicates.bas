Sub LimpiarDatosBaskets()

    Dim wsOriginal As Worksheet
    Dim wsDuplicates As Worksheet
    Dim tblOriginal As ListObject
    Dim tblDuplicates As ListObject
    Dim rng As Range
    Dim dict As Object
    Dim cell As Range
    Dim key As Variant
    Dim duplicatesCount As Long
    
    ' Creamos diccionario para contar duplicados
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Referenciar la hoja y la tabla original
    Set wsOriginal = ThisWorkbook.Sheets("Baskets")
    Set tblOriginal = wsOriginal.ListObjects("TablaF")
    
    ' Creamos una copia de la hoja Baskets y la nombramos Duplicates
    wsOriginal.Copy After:=wsOriginal
    Set wsDuplicates = ActiveSheet
    On Error Resume Next
    wsDuplicates.Name = "Duplicates"
    On Error GoTo 0
    
    ' Obtener la tabla en la hoja Duplicates y renombrarla
    Set tblDuplicates = wsDuplicates.ListObjects(1)
    tblDuplicates.Name = "TablaSinDuplicados"
    
    ' Contar duplicados en la columna Caseid
    Set rng = tblDuplicates.ListColumns("Caseid").DataBodyRange
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            If dict.exists(cell.Value) Then
                dict(cell.Value) = dict(cell.Value) + 1
            Else
                dict.Add cell.Value, 1
            End If
        End If
    Next cell
    
    ' Contar valores duplicados
    duplicatesCount = 0
    For Each key In dict.keys
        If dict(key) > 1 Then
            duplicatesCount = duplicatesCount + (dict(key) - 1)
        End If
    Next key
    
    ' Mostrar mensaje con el número de casos duplicados
    MsgBox duplicatesCount & " casos duplicados encontrados en la columna Caseid", vbInformation
    
    ' Eliminar filas duplicadas en la hoja Duplicates
    rng.RemoveDuplicates Columns:=1, Header:=xlYes
    
    ' Mostrar mensaje indicando que los duplicados han sido removidos
    MsgBox "Duplicados Removidos", vbInformation
End Sub