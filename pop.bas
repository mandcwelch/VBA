Attribute VB_Name = "pop"
Sub color()

Selection.Interior.color = RGB(255, 100, 100)

End Sub

Sub populate_test()
Dim rng As Range
Dim r As Integer
On Error Resume Next
With Worksheets("Sheet1").Columns
 .ColumnWidth = .ColumnWidth / 3.45
End With
    
h = InputBox("How many iterations?")

r = InputBox("What size?")

Set rng = Range(Cells(1, 1), Cells(r, r))

For i = 1 To h

For Each rCell In rng.Cells

If rCell.Offset(1, 1).Interior.color = RGB(255, 100, 100) Then rCell.Interior.color = RGB(100, 100, 100)
If rCell.Offset(1, 0).Interior.color = RGB(100, 100, 100) Then

rCell.Interior.color = RGB(255, 255, 25)
rCell.Offset(1, 1).Interior.color = RGB(255, 255, 25)
rCell.Offset(2, 2).Interior.color = RGB(255, 255, 25)
rCell.Offset(-1, 1).Interior.color = RGB(255, 255, 25)

End If

If rCell.Interior.color = RGB(255, 255, 25) Then

rCell.Offset(1, 0).Interior.color = RGB(25, 255, 25)
rCell.Offset(0, 1).Interior.color = RGB(25, 255, 25)
rCell.Offset(-1, 0).Interior.color = RGB(25, 255, 25)
rCell.Offset(0, -1).Interior.color = RGB(25, 255, 25)

End If



Next rCell

Application.Wait (Now + TimeValue("00:00:01"))

Next i

End Sub

