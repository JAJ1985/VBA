' Input
Dim myFile As String
Dim sName As String
Dim scrBook As Workbook
   
   
   Sub import_file(control As IRibbonControl)
   
   
   myFile = Application.GetOpenFilename(, , "Browse for Workbook")


    On Error Resume Next

    Set scrBook = Application.Workbooks.Open(myFile, _
                  UpdateLinks:=False, _
                  ReadOnly:=True, _
                  AddToMRU:=False)

    On Error GoTo 0

     If scrBook Is Nothing Then
     MsgBox "Sorry, the file was NOT found! Try again!"

     Else

     scrBook.ActiveSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

     scrBook.Close False

    sName = Application.InputBox _
              (Prompt:="Enter new worksheet name")

    ActiveSheet.Name = sName
     End If
   
   
   End Sub

------------------------------------------------------------------------------------------------

Dim CRow As Long
Dim RRow As Long
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim r As Integer
Dim find_error As Integer

Dim cell_value As Variant
Dim cell_value_2 As Variant

Dim case_number() As Variant
   
Function find_number(find_num As Variant) As Variant

On Error Resume Next
        Selection.Find(What:=find_num, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
                           :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                           False, SearchFormat:=False).Activate
                           
                    If Err.Number = 0 Then
                              
                              find_error = 0

                         Else
                         
                         '    MsgBox Err.Number
                          '  MsgBox "Number not found"
                            find_error = Err.Number

                     End If

End Function
   
   
Sub import_SF_report(control As IRibbonControl)

CRow = 0
'Stopping Application Alerts
Application.DisplayAlerts = False

  Worksheets("reimburs").Copy After:=Worksheets("reimburs")

  ActiveSheet.Name = "reimburs_check"

  Worksheets("cases").Copy After:=Worksheets("cases")

  ActiveSheet.Name = "Cases_check"



Worksheets("Cases_check").Copy After:=Worksheets("Cases_check")
ActiveSheet.Name = "Case Number"




Columns("A:A").Select
 ActiveSheet.Range("$A$1:$A$9000").RemoveDuplicates Columns:=Array(1), _
 Header:=xlYes
 Range("A1").Select


     ActiveSheet.Range("A2").Select

        Do While ActiveCell.Value <> Empty
        ActiveCell.Offset(1, 0).Select
        CRow = CRow + 1
        If ActiveCell.Value = "" Then
            Exit Do
        End If
    Loop

     'MsgBox CRow
  'CRow = Cells(Rows.Count, 1).End(xlUp).Row
     ActiveSheet.Range("A2").Select
   'MsgBox lRow

   i = 0

    For i = 0 To CRow

        'MsgBox i

        test = ActiveCell.Offset(i, 0).Value

        'MsgBox test

            ReDim Preserve case_number(i) ' Dynamic arrays
            case_number(i) = test

    Next i

'MsgBox i

Sheets("reimburs").Select

Sheets("Case Number").Delete


ActiveSheet.Range("A1").Select
RRow = Cells(Rows.Count, 1).End(xlUp).Row

'MsgBox RRow





 ActiveSheet.Range("D1").Select

   j = 0

  For j = 0 To RRow

    reim = ActiveCell.Offset(j, 0).Value

    'MsgBox reim

    If reim = "" Then

        'MsgBox "empty"

        ActiveCell.Offset(j, 0).Value = "'"
    End If

    'ActiveCell.Offset(1, 0).Select

  Next j


   ActiveSheet.Range("A1").Select
  j = 0


Columns("A:A").Insert Shift:=xlToRight


     For j = 0 To RRow

         ActiveCell.FormulaR1C1 = "=IF(RC[1]=R[-1]C[1],""Good"",""NL"")"
         ActiveCell.Offset(1, 0).Select

     Next j


 Columns("A:A").Copy
 Columns("A:A").PasteSpecial xlPasteValues


  ActiveSheet.Range("A2").Select

     j = 0
    r = 0

    For j = 0 To RRow

        cell_value = ActiveCell.Value

       ' MsgBox cell_value

        If cell_value = "NL" Then
                r = Selection.Row
                 'MsgBox r
                Cells(r, 1).EntireRow.Insert
                ActiveCell.Offset(1, 0).Select

        End If

        ActiveCell.Offset(1, 0).Select

     Next j


    Columns(1).EntireColumn.Delete




CRow = 0
i = 0
'I need to finish off the find function

Sheets("Cases_check").Select
ActiveSheet.Range("A1").Select
CRow = Cells(Rows.Count, 1).End(xlUp).Row
Columns("A:A").Insert Shift:=xlToRight

    For i = 0 To CRow

         ActiveCell.FormulaR1C1 = "=IF(RC[1]=R[-1]C[1],""Good"",""NL"")"
         ActiveCell.Offset(1, 0).Select

    Next i

     Columns("A:A").Copy
 Columns("A:A").PasteSpecial xlPasteValues

j = 0
r = 0

ActiveSheet.Range("A1").Select
    For j = 0 To CRow

        cell_value = ActiveCell.Value

       ' MsgBox cell_value

        If cell_value = "NL" Then
                r = Selection.Row
                 'MsgBox r
                'Cells(r, 1).EntireRow.Insert
                ActiveCell.EntireRow.Resize(3).Insert

                ActiveCell.Offset(3, 0).Select

        End If

        ActiveCell.Offset(1, 0).Select

     Next j


      Columns(1).EntireColumn.Delete

     ' Need to use array to find the case number

      j = 0

       For j = LBound(case_number) To UBound(case_number)

             Sheets("reimburs").Select
            Application.CutCopyMode = False
                Columns("A:A").Select

            array_output = case_number(j)

            find_number (array_output)

            'MsgBox array_output

           If array_output <> "" Then
            If find_error = 0 Then

                ActiveCell.Offset(0, 1).Select
                ActiveCell.CurrentRegion.Copy
                Sheets("Cases_check").Select
                Columns("A:A").Select
                find_number (array_output)
                ActiveCell.Offset(0, 17).Select
                ActiveCell.PasteSpecial

            ElseIf find_error = 91 Then

                'MsgBox "Not Find"
                   array_output = ""
            End If
           End If
        Next j

   Sheets("Cases_check").Select
   ActiveSheet.Range("A4").Select
   ActiveCell.Offset(0, 17).Select
   ActiveCell.Value = "Case Number"
   ActiveCell.Offset(0, 1).Select
   ActiveCell.Value = "Reimbursement Request: Reimbursement Request Number"
   ActiveCell.Offset(0, 1).Select
   ActiveCell.Value = "Type of Request"
   ActiveCell.Offset(0, 1).Select
   ActiveCell.Value = "Amount"
   ActiveCell.Offset(0, 1).Select
   ActiveCell.Value = "Check Number"
   ActiveCell.Offset(0, 1).Select
   ActiveCell.Value = "Consumer Name"
   
   'Remove worksheet
   Sheets("cases").Delete
   Sheets("reimburs").Delete
   Sheets("reimburs_check").Delete
   
   'Format columns
   Rows(1).EntireRow.Delete
   Rows(1).EntireRow.Delete
   Rows(1).EntireRow.Delete
   
   ActiveSheet.Range("Q:V").ColumnWidth = 18
   ActiveSheet.Range("A1").Select
   Range("A1").EntireRow.Font.Bold = True

   
End Sub


