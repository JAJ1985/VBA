Sub SaveAsCsv()
Dim wb As Workbook
Dim sh As Worksheet
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog
'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then Exit Sub

'Target File Extension (must include wildcard "*")
  myExtension = "*.xlsx*"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
    Set wb = Workbooks.Open(Filename:=myPath & myFile)
    nameWb = myPath & Left(myFile, InStr(1, myFile, ".") - 1) & ".csv"
    ActiveWorkbook.SaveAs Filename:=nameWb, FileFormat:=xlCSV
    ActiveWorkbook.Close savechanges:=False
    'Get next file name
      myFile = Dir
  Loop
'Reset Macro Optimization Settings
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Application.Calculation = xlCalculationAutomatic
End Sub
