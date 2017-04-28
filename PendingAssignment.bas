Attribute VB_Name = "PendingAssignment"
Sub Pending_Assignment()

Dim m As Integer
Dim endrange As Integer
Dim endrange2 As Integer
Dim c As Integer
Dim marea As Integer

Application.ScreenUpdating = False

marea = 1

endrange = Range(Cells(1, 1), Cells(1, 1).End(xlDown)).Rows.Count
ActiveSheet.name = "Pendings"
Sheets.Add.name = "Meetings"
Sheets("Pendings").Select
Columns("J:V").Delete
Columns("F:H").Delete

For i = 2 To endrange

If Cells(i, 2) = "driver_assigned" Then Cells(i, 2) = ""
If Cells(i, 2) = "garage_confirmed" Then Cells(i, 2) = ""
If Cells(i, 2) = "driver_onsite" Then Cells(i, 2) = ""

Next i

Range(Cells(1, 2), Cells(endrange, 2)).Select
Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete

endrange2 = Range(Cells(1, 1), Cells(1, 1).End(xlDown)).Rows.Count

For c = 2 To endrange

If Cells(c, 2) = "garage_assigned" Then Cells(c, 2).Interior.color = RGB(255, 0, 0)
If Cells(c, 2) = "mod_pending" Then Cells(c, 2).Interior.color = RGB(155, 155, 0)

Next c

For m = 2 To endrange2

    If Cells(m, 4) <> "" Then
    Cells(m, 1).EntireRow.Cut
    Sheets("Meetings").Select
    Cells(marea, 1).Insert
    marea = marea + 1
    Sheets("Pendings").Select
    End If

Next m

Range(Cells(1, 2), Cells(endrange2, 2)).Select
Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
Cells(1, 1).EntireRow.Copy
    Sheets("Meetings").Select
    Cells(1, 1).Insert
    Sheets("Pendings").Select

Columns("E:E").Delete
Columns("D:D").Delete
Sheets("Meetings").Select
Columns.AutoFit
Sheets("Pendings").Select
Columns.AutoFit

Application.ScreenUpdating = True

End Sub
