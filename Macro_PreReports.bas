Sub ImportMHTFilesForAnalysts()

    Dim ws As Worksheet
    Dim wb As Workbook
    Dim FolderPath As String
    Dim Analysts As Variant
    Dim Analyst As Variant
    Dim MonthName As String
    Dim CurrentPath As String
    Dim File As Object
    Dim FSObject As Object
    Dim wsDest As Worksheet
    Dim LastRow As Long
    Dim QueryTable As QueryTable
    Dim ConnectionString As String

    Application.ScreenUpdating = False
    
    ' Definir la ruta principal
    FolderPath = "\\..."
    
    ' Lista de nombres de analistas
    Analysts = Array("Analyst1", "Analyst2", "Analyst3") ' Reemplazar con los nombres reales
    
    ' Mes específico a importar
    MonthName = "May" ' Cambiar al mes deseado
    
    ' Crear el objeto FileSystemObject
    Set FSObject = CreateObject("Scripting.FileSystemObject")
    
    ' Establecer el archivo de Excel de destino
    Set wb = ThisWorkbook
    Set wsDest = wb.Sheets("Sheet1") ' Cambiar "Sheet1" al nombre de la hoja de destino
    
    ' Limpiar la hoja de destino
    wsDest.Cells.Clear
    
    ' Añadir encabezados a la hoja de destino en la fila 4
    wsDest.Cells(4, 1).Value = "Analyst"
    wsDest.Cells(4, 2).Value = "File Name"
    
    ' Inicializar la fila de destino en la fila 5
    LastRow = 5
    
    ' Iterar a través de cada analista
    For Each Analyst In Analysts
    
        ' Construir la ruta del mes específico para el analista actual
        CurrentPath = FolderPath & Analyst & "\" & MonthName & "\"
        
        ' Verificar si la carpeta del mes específico existe
        If FSObject.FolderExists(CurrentPath) Then
            
            ' Iterar a través de cada archivo .mht en la carpeta del mes específico
            For Each File In FSObject.GetFolder(CurrentPath).Files
                If LCase(FSObject.GetExtensionName(File.Name)) = "mht" Then
                    
                    ' Agregar el nombre del analista y el nombre del archivo a la hoja de destino
                    wsDest.Cells(LastRow, 1).Value = Analyst
                    wsDest.Cells(LastRow, 2).Value = File.Name
                    
                    ' Importar el archivo .mht
                    ConnectionString = "TEXT;" & File.Path
                    Set QueryTable = wsDest.QueryTables.Add(Connection:=ConnectionString, Destination:=wsDest.Cells(LastRow, 3))
                    
                    With QueryTable
                        .TextFileParseType = xlDelimited
                        .TextFileConsecutiveDelimiter = False
                        .TextFileTabDelimiter = False
                        .TextFileSemicolonDelimiter = False
                        .TextFileCommaDelimiter = False
                        .TextFileSpaceDelimiter = False
                        .TextFileOtherDelimiter = "|"
                        .Refresh BackgroundQuery:=False
                    End With
                    
                    ' Incrementar la fila de destino
                    LastRow = LastRow + 1
                    
                End If
            Next File
        End If
    Next Analyst
    
    ' Eliminar filas desde la fila 6 hacia abajo
    wsDest.Rows("6:" & wsDest.Rows.Count).Delete
    
    ' Ajustar el ancho de todas las columnas
    wsDest.Columns("A:B").AutoFit

    Application.ScreenUpdating = True
    
    MsgBox "Archivos importados con éxito.", vbInformation

End Sub