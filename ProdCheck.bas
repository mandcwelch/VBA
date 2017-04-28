Attribute VB_Name = "ProdCheck"
Sub Productivity()

Dim endRow As Integer
Dim i As Integer
Dim out As Integer
Dim inb As Integer

Application.ScreenUpdating = True

Columns("A:A").Delete

ActiveSheet.name = "Main"

endRow = Cells(1, 1).End(xlDown).Row

For i = 2 To endRow

If Cells(i, 2) = "None" Then Cells(i, 2) = ""

Next i

' Delete blank rows

Range(Cells(1, 2), Cells(endRow + 1, 2)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
endRow = Cells(1, 1).End(xlDown).Row

For i = 2 To endRow

If Cells(i, 2) = "Dial-out" And Cells(i + 1, 2) = "Dial-out" Then Cells(i, 2) = ""
    


Next i

Range(Cells(1, 2), Cells(endRow + 1, 2)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
endRow = Cells(1, 1).End(xlDown).Row
'Seperates outbound calls
out = 2

For i = 2 To endRow

If Cells(i, 2) = "Dial-out" Then
Range(Cells(i, 1), Cells(i, 3)).Copy
Cells(out, 4).Insert
out = out + 1

End If

Next i

'Gathers inbound calls
inb = 2

For i = 2 To endRow

If Cells(i, 2) = "Inbound" Then
Range(Cells(i, 1), Cells(i, 3)).Copy
Cells(inb, 9).Insert
inb = inb + 1

End If


Next i

Columns("A:C").Delete
Columns("G:G").Delete
Columns("B:B").Delete
Columns("c:e").Delete

Range("A1") = "Name"
Range("B1") = "Inbound Call Total"
Range("C1") = "Outbound Call Total"

'sort and transfer data
endRow = Cells(1, 1).End(xlDown).Row
chk = Cells(1, 4).End(xlDown).Row
For i = 2 To endRow

If Cells(i, 1) <> Cells(i, 4) Then
Cells(i, 4) = ""
Cells(i, 5) = ""
Range(Cells(i + 1, 4), Cells(chk, 5)).Cut
Cells(i, 4).Insert
chk = chk - 1
i = i - 1
End If
Next i


End Sub
