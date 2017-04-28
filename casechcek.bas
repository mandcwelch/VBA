Attribute VB_Name = "casechcek"
Sub casecheck()
Dim vnd As String
Dim vndlen As Integer

endRow = Cells(1, 1).End(xlDown).Row

vnd = InputBox("Please enter the vendor to check")

vndlen = Len(vnd)

For i = 2 To endRow

If Left(Cells(i, 9), vndlen) <> vnd Then Cells(i, 9) = ""

Next i

Range(Cells(2, 9), Cells(endRow, 9)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete

endRow = Cells(1, 1).End(xlDown).Row

For i = 2 To endRow

If Left(Cells(i, 15), 6) <> "Vendor" Then Cells(i, 15) = ""

Next i



Range(Cells(2, 15), Cells(endRow, 15)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete

Columns("K:Q").Delete
Columns("H:I").Delete
Columns("E:F").Delete
Columns("C:C").Delete
Columns("A:A").Delete

Columns.AutoFit


Dim endcolumn As Integer
Dim h As Integer
Dim hcell As Range
Dim hcell2 As Range
Range("A1").Select

endcolumn = Range(Selection, Selection.End(xlToRight)).Columns.Count

h = 2
Do
Set hcell = Cells(h, 1)
Set h2cell = Cells(h, endcolumn)

Range(hcell, h2cell).Select
Selection.Interior.color = RGB(213, 232, 255)

h = h + 2

Loop While Not IsEmpty(Cells(h, 1))

With Range("A1:E1")
    .Interior.color = RGB(55, 55, 255)
    .Font.color = RGB(255, 255, 255)
End With

End Sub
