Attribute VB_Name = "DuplicateFill"
Sub duplicate_name()

strt = InputBox("Plese enter the number of the first row to check")
endRow = InputBox("Please enter the number of the last row")
col = InputBox("Please enter the column to check")


For i = strt To endRow

If Cells(i, col).Value = "" Then Cells(i, col).Value = Cells(i - 1, col).Value

Next i

End Sub
