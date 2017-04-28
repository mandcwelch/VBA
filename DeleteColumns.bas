Attribute VB_Name = "DeleteColumns"
Sub DeleteColumn()
Dim ColNeeded As Boolean
Dim rCell As Range
Dim rRng As Range
Dim endCol As Integer
Dim endRow As Integer

ColNeeded = False

endCol = InputBox("Enter Column to Check")

Range("A1").Select
Range(Selection, Selection.End(xlDown)).Select
e = Selection.Rows.Count

    For i = 3 To e
    
        If Cells(i, endCol).Value <> 0 Then ColNeeded = True
        
    Next i

Cells(1, endCol).Activate


If ColNeeded = False Then ActiveCell.EntireColumn.Delete

End Sub
