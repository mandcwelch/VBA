Attribute VB_Name = "DeleteDuplicateInLine"
Sub InLine()

endRow = Cells(1, 1).End(xlDown).Row
chk = Cells(1, 4).End(xlDown).Row
For i = 2 To endRow

If Cells(i, 1) <> Cells(i, 4) Then
Cells(i, 4) = ""
Range(Cells(i + 1, 4), Cells(chk, 5)).Cut
Cells(i, 4).Insert
chk = chk - 1
i = i - 1
End If
Next i


End Sub

Sub nextline()
Dim look As String
Dim i As Integer
Dim endRow As Integer
Dim chkrow As Integer

endRow = Cells(1, 2).End(xlDown).Row
chkrow = Cells(1, 1).End(xlDown).Row

For i = 2 To endRow

look = Cells(i, 2).Value
Range(Cells(1, 1), Cells(chkrow, 1)).Find(look).Delete

Next i

End Sub
