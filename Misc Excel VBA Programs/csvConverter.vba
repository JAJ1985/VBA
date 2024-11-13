Public Sub csvToxls()
    Dim FSO As Object
    Dim folder As Object
    Dim wb As Object
    
    csvPath = "Input csv path"
    xlsPath = "Input xlx path"
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set csvFolder = FSO.GetFolder(csvPath)
    
    If FSO.FolderExists(xlsPath) = False Then
        FSO.CreateFolder (xlsPath)
    End If
    
    Set xlsFolder = FSO.GetFolder(xlsPath)
    
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
        
    For Each wb In csvFolder.Files
        If LCase(Right(wb.Name, 3)) = "csv" Then
            Set activeWB = Workbooks.Open(wb)
            activeWB.SaveAs Filename:=xlsPath & "\" & Left(activeWB.Name, Len(activeWB.Name) - 3) & "xlsx", FileFormat:=xlOpenXMLWorkbook
            activeWB.Close True
        End If
    Next
    
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
End Sub
