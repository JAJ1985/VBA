Option Explicit

Dim myFile As String
Dim sName As String
Dim scrBook As Workbook


Dim lRow As Long
Dim LCol As Long
Dim Col As Long

Dim MyRow As Integer
Dim MyCol As Integer

Dim ProdCol As Integer
Dim SubCol As Integer
Dim SpecCol As Integer
Dim ProgCol As Integer


Dim cell_check As Variant
Dim a_cell As Variant


Dim join_string As String
    
Dim find_output() As Variant
Dim cellref As Variant

Dim mess As Integer
Dim array_output As Variant


Dim k As Integer
Dim i As Integer
Dim x As Integer
Dim l As Integer
Dim y As Integer

Dim j As Integer

Dim find_error As Integer
Dim find_word As Integer


Function find_text(text_find As Variant) As Variant
                      
    On Error Resume Next
            Cells.find(What:=text_find, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
            :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
             False, SearchFormat:=False).Activate
             
                   If Err.Number = 0 Then
                       
                              
                         Else
                                   
                            find_word = Err.Number
                         
                        End If


End Function

Function find_number(find_num As Variant) As Variant

On Error Resume Next
        Selection.find(What:=find_num, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
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

Sub Select_Range()

'Range(Cells(MyRow, MyCol), Cells(MyRow, MyCol + 7)).Select
Range(Cells(x, MyCol), Cells(x, MyCol + 7)).Select
End Sub


Sub findtext(control As IRibbonControl)

'----------------Import Data into Excel-----------------------------

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
    
     End If
'----------------------------------------------------------------

    Call find_text("WORKFLOW_BLOGPOSTTAGS")
    
        'MsgBox find_word

    If find_word = 0 Then
    
    '==========================Check for WORKFLOW_BLOGPOSTTAGS ===========
    
    cell_check = ActiveCell.Address
    'MsgBox cell_check
    
    a_cell = ActiveCell.Column
    'MsgBox a_cell

    
    If ActiveCell.Offset(0, 1) <> vbNullString Then
    
          Col = Cells(1, Columns.Count).End(xlToLeft).Column
         ' MsgBox Col
          ActiveCell.EntireColumn.Cut
        
          
          Columns(Col + 1).Insert Shift:=xlToRight
          
    End If
        
    '=====================================================================

     Range("A1").Select
     lRow = Cells(Rows.Count, 1).End(xlUp).Row
     'MsgBox lRow
    
    
     Range("A1").Select
       LCol = Cells(1, Columns.Count).End(xlToLeft).Column
       'MsgBox LCol
                                        
                ActiveCell.Offset(0, LCol + 7).Select
                ActiveCell.Value = "Product"
                
                ProdCol = ActiveCell.Column
                'MsgBox ProdCol
                
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Value = "Subject"
                
                SubCol = ActiveCell.Column
                'MsgBox SubCol
                
                
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Value = "Special Situations"
                
                SpecCol = ActiveCell.Column
                'MsgBox SpecCol
                
                
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Value = "Program Events"

                ProgCol = ActiveCell.Column
                'MsgBox ProgCol



    Call find_text("WORKFLOW_BLOGPOSTTAGS")
    
    MyRow = ActiveCell.Row
    MyCol = ActiveCell.Column
    
  '  MsgBox "This is Row number:" & MyRow
   ' MsgBox "This is Column number:" & MyCol
    
    
    Selection.EntireColumn.Select


     Selection.TextToColumns DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
                :=Array(Array(1, 1), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
                            
      For x = 1 To lRow ' this will be the rows
                            
               ' Range(Cells(x, 2), Cells(x, 8)).Select
             
                'Range(Cells(MyRow, MyCol), Cells(MyRow, MyCol + 7)).Select
                    
                 Call Select_Range
                    
                'MyRow = MyRow + 1
                    

                For i = 1 To 4 'this will search for the numbers
                
                Erase find_output
                k = 0
                
                    For j = 1 To 7
                   
                    
                    find_number (i)
                                            
                        'MsgBox find_error
                        
                        
                        If find_error = 91 Then
                        
                            'MsgBox "empty"
                            Exit For
                        
                        Else
                        
                            cellref = ActiveCell.Value
                            'MsgBox cellref
                        

                            If (Not Not find_output) <> 0 Then
                                For l = LBound(find_output) To UBound(find_output)
                                                
                                   ' MsgBox find_output(l)
                                    
    
                                    array_output = find_output(l)
                                    
                                     If array_output = cellref Then
                                    
                                            
                                          Exit For
                                                
                                    Else
                                        ReDim Preserve find_output(k) ' Dynamic arrays
                                        find_output(k) = cellref
                                        
                                    End If
                                    
                                Next l
                                
                                Else
                         ReDim Preserve find_output(k) ' Dynamic arrays
                        find_output(k) = cellref
                        k = k + 1
                        l = 0
                                
                            End If
                                          
                        End If
                            
                    If j = 7 Then

                        If i = 1 Then
                            join_string = Join(find_output, " , ")
                            'MsgBox join_string
                            ActiveSheet.Cells(x, ProdCol).Select
                            ActiveCell.Value = join_string
                            Call Select_Range
                            
                            ElseIf i = 2 Then
                            join_string = Join(find_output, " , ")
                            'MsgBox join_string
                            ActiveSheet.Cells(x, SubCol).Select
                            ActiveCell.Value = join_string
                            Call Select_Range
                            
                            
                            ElseIf i = 3 Then
                            join_string = Join(find_output, " , ")
                            'MsgBox join_string
                            ActiveSheet.Cells(x, SpecCol).Select
                            ActiveCell.Value = join_string
                            Call Select_Range
                        
                            ElseIf i = 4 Then
                            join_string = Join(find_output, " , ")
                            'MsgBox join_string
                            ActiveSheet.Cells(x, ProgCol).Select
                            ActiveCell.Value = join_string
                            Call Select_Range
                            
                            ElseIf i = 5 Then 'Need to implement 5 when country is in the data
                            join_string = Join(find_output, " , ")
                            MsgBox join_string

                        End If

                    End If
                                           
                            
                    Next j
                                       
                
                Next i
                
                If x = lRow Then
                   ' MsgBox "delete column"
                    
                    
                   Call find_text("WORKFLOW_BLOGPOSTTAGS")
                   
                   For y = 1 To 8

                        Columns(ActiveCell.Column).Delete

                    Next y
                    
                End If
    Next x
    
    Columns.AutoFit

    Else
        MsgBox "Field WORKFLOW_BLOGPOSTTAGS is not in spreadsheet."
        
    End If
    
End Sub

