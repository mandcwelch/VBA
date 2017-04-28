Attribute VB_Name = "Ops_or_Groups"
Sub Reservation_Overview_Format()
Attribute Reservation_Overview_Format.VB_ProcData.VB_Invoke_Func = "m\n14"

'Created by Michael Welch 2/19/16 to determine whether trips are Ops trips, Meetings, or Events.
'Please contact Michael if you experience any issues with macro

'Define the labels

Dim conf As Integer
Dim cord As Integer
Dim roun As Integer
Dim dup As Integer
Dim lastcell As Integer

Application.ScreenUpdating = False

'Count total reservations numbers for last row

With ActiveSheet

    lastcell = .Cells(.Rows.Count, "A").End(xlUp).Row

End With

'Add Time column

    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
'Move time our of date column
    
    Columns("I:I").Select
    Selection.TextToColumns destination:=Range("I1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
        
'Add Time rounded column
        
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
'Copies time into time rounded column
    
    Columns("J:J").Select
    Selection.Copy
    Range("K1").Select
    ActiveSheet.Paste
    
'Rounds down the time of the rounded column

    For roun = 2 To lastcell

    Cells(roun, 11).Value = WorksheetFunction.RoundDown(Cells(roun, 11), -2)
    
    Next roun
    
'Names the Dat, Time, and Time rounded columns
    
    Range("I1").Value = "Date"
    Range("J1").Value = "Time CDT"
    Range("K1").Value = "Time (rounded)"
    
'Inserts the Designation column
    
Columns("D:D").Select
Selection.Insert
Range("D1").Value = "Designation"

'Checks to see if the trip has an onsite coordinator

For cord = 2 To lastcell
    
    If Cells(cord, 6).Value = "Y" Then
    
     Cells(cord, 4).Value = "Event"
    
    ElseIf Cells(cord, 6).Value = "N" Then
    
        Cells(cord, 4).Value = "Meeting"
    
    End If
    
Next cord

'Checks to see if the trip has a group ID

For conf = 2 To lastcell

    If Cells(conf, 5).Value = "" Then Cells(conf, 4).Value = "Ops"
    
Next conf

'Deletes duplicate rows

For dup = 2 To lastcell

    If Cells(dup, 1).Value = Cells(dup + 1, 1).Value Or Cells(dup, 1).Value = Cells(dup - 1, 1).Value Then
    
    Cells(dup, 1).Select
    ActiveCell.EntireRow.Select
    Selection.ClearContents
    Else
    
    End If
    
Next dup

'Delete empty rows

On Error Resume Next

    Range("A1:A5000").Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete

Range("A2:A5000").Select
 
Application.ScreenUpdating = True

End Sub

