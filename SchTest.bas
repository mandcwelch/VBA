Attribute VB_Name = "SchTest"
Sub schedule_formater()
'Version 1 created 10/10/2016 by Michael Welch
'Formats schedule for sending out to agents
Dim i As Integer
Dim TotTab As Range
Dim init As Workbook
Dim main As Workbook
Dim rng As Range

'creates copy of schedule

Set init = ActiveWorkbook
ActiveSheet.Copy

Set main = ActiveWorkbook


'Format Schedule Correctly

Columns("A:A").Insert
main.theme.ThemeColorScheme.Load ("P:\Operations\Group Department\Macros\theme")

If Cells(1, 2) = "" Then Rows("1:1").Delete
If Cells(1, 1) = "" Then Columns("A:A").Delete

On Error Resume Next

    Range(Cells(1, 1), Cells(500, 1)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete

On Error GoTo 0

Columns("A:A").Insert
Columns("A:Z").UnMerge


'Sets Day of Week


Set sun = Range("B1:B500").Find("Sunday")
Set mon = Range("B1:B500").Find("Monday")
Set tue = Range("B1:B500").Find("Tuesday")
Set wed = Range("B1:B500").Find("Wednesday")
Set thu = Range("B1:B500").Find("Thursday")
Set fri = Range("B1:B500").Find("Friday")
Set sat = Range("B1:B500").Find("Saturday")

Set sunrng = Range(sun.Offset(3, -1), mon.Offset(-1, -1))
For Each rCell In sunrng.Cells
rCell.Value = "Sunday"
Next rCell

Set monrng = Range(mon.Offset(3, -1), tue.Offset(-1, -1))
For Each rCell In monrng.Cells
rCell.Value = "Monday"
Next rCell


Set tuerng = Range(tue.Offset(3, -1), wed.Offset(-1, -1))
For Each rCell In tuerng.Cells
rCell.Value = "Tuesday"
Next rCell


Set wedrng = Range(wed.Offset(3, -1), thu.Offset(-1, -1))
For Each rCell In wedrng.Cells
rCell.Value = "Wednesday"
Next rCell


Set thurng = Range(thu.Offset(3, -1), fri.Offset(-1, -1))
For Each rCell In thurng.Cells
rCell.Value = "Thursday"
Next rCell


Set frirng = Range(fri.Offset(3, -1), sat.Offset(-1, -1))
For Each rCell In frirng.Cells
rCell.Value = "Friday"
Next rCell


Set satrng = Range(sat.Offset(3, -1), sat.End(xlDown).Offset(0, -1))
For Each rCell In satrng.Cells
rCell.Value = "Saturday"
Next rCell


' Removes Blanks

On Error Resume Next

Rows("1:1").Delete
Rows("2:2").Delete

    Range(Cells(2, 1), Cells(1000, 1)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    

'Orders rows by day of week

endrow1 = Cells(2, 1).End(xlDown).Row

    ActiveSheet.Sort.SortFields.Add Key:= _
        Range(Cells(2, 2), Cells(endrow1, 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:= _
        Range(Cells(2, 1), Cells(endrow1, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday", DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(Cells(2, 1), Cells(endrow1, 26))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

