Attribute VB_Name = "CompanySearch"
Sub Company_Search()

Dim i As Integer
Dim endRow As Integer

compname = InputBox("Please enter the name of the company as listed on sheet")

endRow = Cells(1, 1).End(xlDown).Row

For i = 2 To endRow

If Cells(i, 6) <> compname Then Cells(i, 6) = ""

Next i

Range(Cells(2, 6), Cells(endRow, 6)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete

End Sub
