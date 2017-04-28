Attribute VB_Name = "SJCShuttleCheck"
Sub shuttlecheck()
Dim init As Workbook
Dim Arr As Workbook
Dim Dep As Workbook


endRow = Cells(1, 1).End(xlDown).Row

For i = 2 To endRow

If Cells(i, 6) = "SJC" Or Cells(i, 6) = "SFO" Or Cells(i, 10) = "SJC" Or Cells(i, 10) = "SFO" Then

Cells(i, 14) = 1

End If

Next i



  Range(Cells(2, 14), Cells(endRow, 14)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
  
  Set init = ActiveWorkbook
  Set Arr = Workbooks.Add
  Set Dep = Workbooks.Add
  
  init.Activate
  ActiveSheet.UsedRange.Copy
  
  Arr.Activate
  Cells(1, 1).Insert
  
  init.Activate
  ActiveSheet.UsedRange.Copy

  Dep.Activate
  Cells(1, 1).Insert

Arr.Activate
Columns.AutoFit

arrendrow = Cells(1, 1).End(xlDown).Row

For i = 2 To arrendrow

If Cells(i, 6) <> "SJC" And Cells(i, 6) <> "SFO" Then

Cells(i, 6) = ""

End If

Next i

Range(Cells(2, 6), Cells(arrendrow, 6)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete




Dep.Activate
Columns.AutoFit

dependrow = Cells(1, 1).End(xlDown).Row

For i = 2 To dependrow

If Cells(i, 10) <> "SJC" And Cells(i, 10) <> "SFO" Then

Cells(i, 10) = ""

End If

Next i

Range(Cells(2, 10), Cells(arrendrow, 10)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete


End Sub

Sub vehcheck()

endRow = Cells(1, 1).End(xlDown).Row

For i = 2 To endRow

if cells

End Sub
