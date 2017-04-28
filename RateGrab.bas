Attribute VB_Name = "RateGrab"
Sub RateGrab()

Dim Start As Range
Dim rtsheet As Workbook
Dim crs As Workbook
Dim check As Range

Set rtsheet = ActiveWorkbook

Workbooks.Add

Set crs = ActiveWorkbook
crs.Activate
rtsheet.Activate
Range("A1").Select

'Do While check Is Not Empty

Set Start = Cells.Find("Vendor Name - Ranking")

check = Start.Offset(2, 0)

endRow = check.Offset(0, 2).End(xlDown).Row

'For i = check.Offset(0, 2).Row To endrow

Cells(i, 1) = check.Offset(0, 1)

'cells(i,5) = cells(i,

End Sub


