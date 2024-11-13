' Module 1
Sub MergeExcelFiles()
    Dim fnameList, fnameCurFile As Variant
    Dim countFiles, countSheets As Integer
    Dim wksCurSheet As Worksheet
    Dim wbkCurBook, wbkSrcBook As Workbook
 
    fnameList = Application.GetOpenFilename(FileFilter:="Microsoft Excel Workbooks (*.xls;*.xlsx;*.xlsm;*.csv),*.xls;*.xlsx;*.xlsm;*.csv", Title:="Choose Excel files to merge", MultiSelect:=True)
 
    If (vbBoolean <> VarType(fnameList)) Then
 
        If (UBound(fnameList) > 0) Then
            countFiles = 0
            countSheets = 0
 
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
 
            Set wbkCurBook = ActiveWorkbook
 
            For Each fnameCurFile In fnameList
                countFiles = countFiles + 1
 
                Set wbkSrcBook = Workbooks.Open(Filename:=fnameCurFile)
 
                For Each wksCurSheet In wbkSrcBook.Sheets
                    countSheets = countSheets + 1
                    wksCurSheet.Copy After:=wbkCurBook.Sheets(wbkCurBook.Sheets.Count)
                Next
 
                wbkSrcBook.Close SaveChanges:=False
 
            Next
 
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
 
            MsgBox "Processed " & countFiles & " files" & vbCrLf & "Merged " & countSheets & " worksheets", Title:="Merge Excel files"
        End If
 
    Else
        MsgBox "No files selected", Title:="Merge Excel files"
    End If
End Sub

' Module 2
Sub DeleteSheets1()
    Dim xWs As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.Name <> "Main" And xWs.Name <> "Combined" Then
            xWs.Delete
        End If
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub DeleteAll()

    Dim xWs As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If Xws

End Sub

' Module 3
Sub Combine()
    Dim J As Integer
    Dim s As Worksheet

    On Error Resume Next
    Sheets(1).Select
    Worksheets.Add ' add a sheet in first place
    Sheets(1).Name = "Combined"

    ' copy headings
    Sheets(2).Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets(1).Range("A1")

    For Each s In ActiveWorkbook.Sheets
        If s.Name <> "Combined" Then
            Application.GoTo Sheets(s.Name).[a1]
            Selection.CurrentRegion.Select
            ' Don't copy the headings
            Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
            Selection.Copy Destination:=Sheets("Combined"). _
              Cells(Rows.Count, 1).End(xlUp)(2)
        End If
    Next
End Sub

'Module 4
Sub Merge_Combine()

Dim Ws As Worksheet
Dim LR As Long
Dim EachCell As Variant
Dim Rng As Range, Cell As Range

TotalCount = 3

Counter = 1

Call MergeExcelFiles
Call Combine
Call DeleteSheets1

End Sub

