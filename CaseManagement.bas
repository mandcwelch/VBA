Attribute VB_Name = "CaseMAnagement"
Sub case_management()

Dim GroupID As String
Dim lastcell As Integer
Dim i As Integer

lastcell = Cells(Rows.Count, "A").End(xlUp).Row

GroupID = InputBox("Please enter the group ID for the group you are checking on")

For i = 2 To lastcell

If Cells(i, 3).Value <> GroupID Then Cells(i, 3).Value = ""

Next i

Range("C1:C5000").Select
Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete

End Sub
