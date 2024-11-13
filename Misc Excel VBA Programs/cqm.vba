Option Explicit
Option Base 1

Sub CopyRows()
Dim LastRow As Long
Dim NbRows As Long
Dim RowList()
Dim I As Long, J As Long, K As Long
Dim RowNb As Long
    Sheets("Data").Activate
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
    NbRows = IIf(LastRow < 200, LastRow * 0.2, 20)
    ReDim RowList(2 To NbRows)
    K = 2
    For I = 2 To NbRows
        RowNb = Rnd() * LastRow
        For J = 2 To K
            If (RowList(J) = RowNb) Then GoTo NextStep
        Next J
        RowList(K) = RowNb
        Rows(RowNb).Copy Destination:=Sheets("Results").Cells(K, "A")
        K = K + 1
NextStep:
    Next I
End Sub


'Put in Reset/Clear Option
