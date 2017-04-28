Attribute VB_Name = "Core_Manifest_Format"
Sub Core_Manifest_Format()

'Adapted by Michael Welch from the original created by Yeong Cheng and Alex Rawlings.
'Adapted 2/5/16 - Please contact Michael if you have any questions or issues.
'Modified to work with Core downloads 3/16/16.
'This macro will automatically format Savoya Manifests straight from Core download.



On Error GoTo initialErrorHandler:

'Prepares the logo for printing.
Dim logoPath As String
logoPath = "P:\Operations\Group Department\Information\Training\Macros\savoya_logo2.jpg"

pExists = True

'Create and name the three sheets.

On Error GoTo namingErrorHandler:
Sheets(1).name = "Offsites"
Sheets.Add.name = "Departures"
Sheets.Add.name = "Arrivals"
On Error GoTo 0

'Adds the group ID for TC copy

GroupID = InputBox("Enter GroupID")

'Allow selection of logo if not connected to P drive

On Error Resume Next
If Dir("P:\", vbDirectory) = vbNullString Then pExists = False
On Error GoTo 0

If pExists = False Then
    MsgBox ("Not connected to P: Drive. Please select the Savoya Logo")
    logoPath = Application.GetOpenFilename
    If logoPath = "" Then
        MsgBox ("Nothing selected. Please try again")
        GoTo finish:
    End If
End If

'Turn off screen updating

Application.ScreenUpdating = False



'*********Format the 3 pages***********


'Select Offsite Sheet

Sheets("Offsites").Select


'Formats the time

TimeFormat ("E2")
TimeFormat ("G2")



'Marks Offsite Trips and sorts by segment

Dim x As Integer
Dim endrange As Integer
Cells(1, 2).Select
endrange = Selection.End(xlDown).Row

For x = 2 To endrange

If Cells(x, 1).Value = "" And Cells(x, 3).Value <> Empty Then Cells(x, 1).Value = "offsite"

Next x

Range("A1").Select

   ActiveSheet.UsedRange.Sort Key1:=Range("A2"), order1:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal
        
Range("A1").Select


'Adds a space between arrivals, departures, and offsites

On Error Resume Next
  Cells.Find(What:="departure").Activate
  
    If ActiveCell = "Departure" Then
    
    ActiveCell.EntireRow.Select
    
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Else
Resume
    End If

    On Error GoTo 0
    On Error Resume Next
Range("A1").Select

      Cells.Find(What:="offsite", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=True, SearchFormat:=False).Activate

   If ActiveCell = "offsite" Then
    
    ActiveCell.EntireRow.Select
    
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Else
Resume
    End If
      On Error GoTo 0

'Removes the top title row

Range("A1").Select

ActiveCell.EntireRow.Delete Shift:=xlUp


' Transfers arrival trips to Arrivals Sheet

If Range("A1") = "Arrival" Then
    arrExists = True
    endrange = Selection.End(xlDown).Row
    totalArr = endrange
    For i = 1 To endrange
        ActiveCell.EntireRow.Cut destination:=Sheets("Arrivals").Cells(i, 1)
        ActiveCell.EntireRow.Delete
    Next i
    ActiveCell.EntireRow.Delete
Else
    MsgBox ("No Arrivals, deleting Arrivals Page")
    Sheets("Arrivals").Delete
    ActiveCell.EntireRow.Delete
End If


' Transfers depature trips to Departures Sheet

Range("A1").Select
If Range("a1") = "" Then ActiveCell.EntireRow.Delete

If Range("A1") = "Departure" Then
    depExists = True
    endrange = Selection.End(xlDown).Row
    totalDep = endrange
    For i = 1 To endrange
        ActiveCell.EntireRow.Cut destination:=Sheets("Departures").Cells(i, 1)
        ActiveCell.EntireRow.Delete
    Next i
    ActiveCell.EntireRow.Delete
    
Else
    MsgBox ("No departure trips, deleting departures page")
    Sheets("Departures").Delete
    
End If

'Checks if there are any offsite trips

If Range("A1") = "offsite" Then

offExists = True
    totalOff = Selection.End(xlDown).Row
Else
    MsgBox ("No offsite trips, deleting offsites page")
    ActiveSheet.Delete
End If
'***********Arrivals**********
If arrExists = False Then GoTo skipArrivals:
    Sheets("Arrivals").Select
    
    Range("A1").Select
    ActiveCell.EntireRow.Insert Shift:=xlDown
    
  
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft

    Range("A1") = "First Name"
    Range("B1") = "Last Name"
    Range("C1") = "Flight Date"
    Range("D1") = "Flight Time"
    Range("E1") = "Pickup Location"
    Range("F1") = "Airline"
    Range("G1") = "Flight Number"
    Range("H1") = "Dropoff Location"
    Range("I1") = "Guests"
    Range("J1") = "Passenger Phone"
    Range("K1") = "Passenger Email"
    Range("L1") = "Confirmation"
    Range("M1") = "Vehicle"
    Range("N1") = "HCP"
    Range("O1") = "VIP"
    Range("P1") = "Shuttle"
    Range("Q1") = "Vendor"
    
       Columns.AutoFit
    
    Columns("L:L").HorizontalAlignment = xlCenter
    Range("L1").HorizontalAlignment = xlLeft
    Columns("M:M").HorizontalAlignment = xlCenter
    Columns("G:G").HorizontalAlignment = xlLeft
    Range("H1").HorizontalAlignment = xlLeft
    
    'Formats the header and adds logo for printing
    
    Range("A1:Q1").Select
    With Selection
        .Font.ColorIndex = 2
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
    End With
    With Selection.Interior
        .ColorIndex = 23
        .Pattern = xlSolid
    End With
   
      Rows("1:1").Insert Shift:=xlDown
    With ActiveSheet.PageSetup
        .LeftHeaderPicture.Filename = logoPath
        .PrintArea = "=" & ActiveSheet.UsedRange.Address
        .PrintTitleRows = "$1:$2"
        .LeftHeader = "&G"
        .RightHeader = "GroupID: " & GroupID & Chr(10) & "Arrival Manifest"
        .CenterFooter = "&D"
        .RightFooter = "&P"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

    
'Sort by Conf, Time, and Date

    ActiveSheet.UsedRange.Sort Key1:=Range("c3"), order1:=xlAscending, key2:=Range( _
        "D3"), order2:=xlAscending, key3:=Range("L3"), Order3:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal

'Adds Highlighting

Range("A3:Q3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=MOD(ROW(),2)=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False

'Adds qoutes for multiple passenger vehicles and deletes all blank columns

Dim p As Integer
Dim rangeend2 As Integer

Cells(3, 12).Select
rangeend2 = Selection.End(xlDown).Row

For p = 3 To rangeend2

If Cells(p, 12).Value = Cells(p - 1, 12).Value Then Cells(p, 13).Value = """"

Next p



'Quotes ("M3")


DeleteBlanks ("Q3") 'Vendor
DeleteBlanks ("P3") 'Shuttle
DeleteBlanks ("O3") 'VIP
DeleteBlanks ("N3") 'HCP
DeleteBlanks ("K3") 'Email
DeleteBlanks ("J3") 'Phone
DeleteBlanks ("I3") 'Guests





skipArrivals:

'**********Departures***********
If depExists = False Then GoTo skipDepartures:
Sheets("Departures").Select

    Range("A1").Select
    ActiveCell.EntireRow.Insert Shift:=xlDown

    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("A1") = "First Name"
    Range("B1") = "Last Name"
    Range("C1") = "Pickup Date"
    Range("D1") = "Pickup Time"
    Range("E1") = "Flight Date"
    Range("F1") = "Flight Time"
    Range("G1") = "Pickup Location"
    Range("H1") = "Airline"
    Range("I1") = "Flight Number"
    Range("J1") = "Dropoff Location"
    Range("K1") = "Guests"
    Range("L1") = "Passenger Number"
    Range("M1") = "Passenger Email"
    Range("N1") = "Confirmation"
    Range("O1") = "Vehicle"
    Range("P1") = "HCP"
    Range("Q1") = "VIP"
    Range("R1") = "Shuttle"
    Range("S1") = "Vendor"
    
    'Formats the header and adds logo for printing
    
        Columns.AutoFit

    Columns("O:O").HorizontalAlignment = xlCenter
    Columns("N:N").HorizontalAlignment = xlCenter
    Range("N1").HorizontalAlignment = xlLeft
    Columns("i:i").HorizontalAlignment = xlLeft
    Range("J1").HorizontalAlignment = xlLeft
    
    Range("A1:S1").Select
    With Selection
        .Font.ColorIndex = 2
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
    End With
    With Selection.Interior
        .ColorIndex = 23
        .Pattern = xlSolid
    End With
    
   Rows("1:1").Insert Shift:=xlDown
    With ActiveSheet.PageSetup
        .LeftHeaderPicture.Filename = logoPath
        .PrintArea = "=" & ActiveSheet.UsedRange.Address
        .PrintTitleRows = "$1:$2"
        .LeftHeader = "&G"
        .RightHeader = "GroupID: " & GroupID & Chr(10) & "Departure Manifest"
        .CenterFooter = "&D"
        .RightFooter = "&P"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

'Sort by Conf, Time, and Date
ActiveSheet.UsedRange.Sort Key1:=Range("C3"), order1:=xlAscending, key2:=Range( _
        "D3"), order2:=xlAscending, key3:=Range("N3"), Order3:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal
        
'Adds Highlighting

Range("A3:S3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=MOD(ROW(),2)=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False

'Adds qoutes for multiple passenger vehicles and deletes all blank columns

Dim o As Integer
Dim rangeend As Integer

Cells(3, 15).Select
rangeend = Selection.End(xlDown).Row

For o = 3 To endrange

If Cells(o, 14).Value = Cells(o - 1, 14).Value Then Cells(o, 15).Value = """"

Next o
'Quotes ("N3")

DeleteBlanks ("S3") 'Vendor
DeleteBlanks ("R3") 'Shuttle
DeleteBlanks ("Q3") 'VIP
DeleteBlanks ("P3") 'HCP
DeleteBlanks ("O3") 'Vehicle
DeleteBlanks ("M3") 'Email
DeleteBlanks ("L3") 'Phone

skipDepartures:

'**********Offsite**********
If offExists = False Then GoTo skipOffsite:
    Sheets("Offsites").Select
    
    Range("A1").Select
    ActiveCell.EntireRow.Insert Shift:=xlDown
    
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft

    

    Range("A1") = "First Name"
    Range("B1") = "Last Name"
    Range("C1") = "Pickup Date"
    Range("D1") = "Pickup Time"
    Range("E1") = "Flight Date"
    Range("F1") = "Flight Time"
    Range("G1") = "Pickup Location"
    Range("H1") = "Airline"
    Range("I1") = "Flight Number"
    Range("J1") = "Stops"
    Range("K1") = "Dropoff Location"
    Range("L1") = "Guests"
    Range("M1") = "Passenger Number"
    Range("N1") = "Passenger Email"
    Range("O1") = "Confirmation"
    Range("P1") = "Vehicle"
    Range("Q1") = "HCP"
    Range("R1") = "VIP"
    Range("S1") = "Shuttle"
    Range("T1") = "Vendor"
    
    'Formats the header and adds logo for printing
        
        Columns.AutoFit
        
    Columns("O:O").HorizontalAlignment = xlCenter
    Columns("N:N").HorizontalAlignment = xlCenter
    Range("N1").HorizontalAlignment = xlLeft
    Columns("i:i").HorizontalAlignment = xlLeft
    Range("J1").HorizontalAlignment = xlLeft
    
    Range("A1:S1").Select
    With Selection
        .Font.ColorIndex = 2
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
    End With
    With Selection.Interior
        .ColorIndex = 23
        .Pattern = xlSolid
    End With
    Rows("1:1").Insert Shift:=xlDown
    With ActiveSheet.PageSetup
        .LeftHeaderPicture.Filename = logoPath
        .PrintArea = "=" & ActiveSheet.UsedRange.Address
        .PrintTitleRows = "$1:$2"
        .LeftHeader = "&G"
        .RightHeader = "GroupID: " & GroupID & Chr(10) & "Offsites Manifest"
        .CenterFooter = "&D"
        .RightFooter = "&P"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
'Sort by Conf, Time, and Date
ActiveSheet.UsedRange.Sort Key1:=Range("C3"), order1:=xlAscending, key2:=Range( _
        "D3"), order2:=xlAscending, key3:=Range("O3"), Order3:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal
   
  'Adds Highlighting

Range("A3:S3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=MOD(ROW(),2)=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
   
'Deletes all blank columns
Columns("M:L").Copy
Columns("O:N").Insert
Columns("M:L").Delete

DeleteBlanks ("T3") 'Vendor
DeleteBlanks ("S3") 'Guests
DeleteBlanks ("R3") 'Passenger Email
DeleteBlanks ("Q3") 'Passenger Email
DeleteBlanks ("N3") 'Flight Number
DeleteBlanks ("M3") 'Airline
DeleteBlanks ("J3") 'Flight Time
DeleteBlanks ("I3")
DeleteBlanks ("H3")
DeleteBlanks ("F3") 'Flight Date
DeleteBlanks ("E3") 'HCP


skipOffsite:

finish:
Sheets(1).Select
Range("A1").Select
Application.ScreenUpdating = True
Exit Sub

namingErrorHandler:
MsgBox ("Error: check that manifest is saved on hard drive as an excel spreadsheet file")
Exit Sub

initialErrorHandler:
MsgBox ("Error occurred before macro could run. Please check code and file formatting")
Exit Sub

End Sub

Sub DeleteColumns(testrange)
cellNeeded = False
Range(testrange).Select
If ActiveCell = "" And Selection.End(xlDown).Row = "65536" Or ActiveCell = "" And Selection.End(xlDown).Row = "1048576" Then
    Range(testrange).EntireColumn.Delete Shift:=xlToLeft
    GoTo Done:
ElseIf Selection.End(xlDown).Row = "65536" Or Selection.End(xlDown).Row = "1048576" Then
    GoTo Done:
End If
If ActiveCell = "no" Or ActiveCell = "No" Or ActiveCell = "Not Provided" Or ActiveCell = "Not provided" Or ActiveCell = "not provided" Or ActiveCell = "xxx-xxx-xxxx" Or ActiveCell = "xxx.xxx.xxxx" Or ActiveCell = "non-VIP" Or ActiveCell = "nonVIP" Or ActiveCell = "non-vip" Or ActiveCell = "nonvip" Then
   revertCell = ActiveCell
   needsRevert = True
   ActiveCell = "0"
End If
testCell = ActiveCell
LastRow = ActiveSheet.UsedRange.Rows.Count
Range(testrange).Select
For i = 3 To LastRow
    ActiveCell.Offset(1, 0).Select
    If ActiveCell = "no" Or ActiveCell = "No" Or ActiveCell = "Not Provided" Or ActiveCell = "Not provided" Or ActiveCell = "not provided" Or ActiveCell = "xxx-xxx-xxxx" Or ActiveCell = "xxx.xxx.xxxx" Or ActiveCell = "non-VIP" Or ActiveCell = "nonVIP" Or ActiveCell = "non-vip" Or ActiveCell = "nonvip" Then
        revertCell = ActiveCell
        needsRevert = True
        ActiveCell = "0"
    End If
    If ActiveCell = testCell Then
    Else
        cellNeeded = True
    End If
Next i
If cellNeeded = False Then
    Range(testrange).EntireColumn.Delete Shift:=xlToLeft
    GoTo Done:
ElseIf needsRevert = True Then
    Range(testrange).Select
    For i = 1 To LastRow - 1
        If ActiveCell = "0" Or ActiveCell = "" Then
            ActiveCell = revertCell
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
End If
Done:
End Sub


Sub DeleteBlanks(testrange)
cellNeeded = False
Range(testrange).Select
If ActiveCell = "" And Selection.End(xlDown).Row = "65536" Or ActiveCell = "" And Selection.End(xlDown).Row = "1048576" Then
    Range(testrange).EntireColumn.Delete Shift:=xlToLeft
    GoTo finish:
ElseIf Selection.End(xlDown).Row = "65536" Or Selection.End(xlDown).Row = "1048576" Then
    GoTo finish:
End If
tryAgain:
Selection.End(xlDown).Select
If Selection.End(xlDown) = "" Then
     LastRow = Selection.Row - 3
     GoTo startSearch:
Else
     GoTo tryAgain:
End If
startSearch:
Range(testrange).Select
For i = 1 To LastRow
     ActiveCell.Offset(1, 0).Select
     If ActiveCell = "" Or ActiveCell = "0" Or ActiveCell = "no" Or ActiveCell = "No" Or ActiveCell = "Not Provided" Or ActiveCell = "Not provided" Or ActiveCell = "not provided" Or ActiveCell = "xxx-xxx-xxxx" Or ActiveCell = "xxx.xxx.xxxx" Or ActiveCell = "non-VIP" Or ActiveCell = "nonVIP" Or ActiveCell = "non-vip" Or ActiveCell = "nonvip" Then
     Else
          cellNeeded = True
     End If
Next i
If cellNeeded = False Then
     Range(testrange).EntireColumn.Delete Shift:=xlToLeft
End If
finish:
End Sub


Sub Quotes(testrange)
Range(testrange).Select
If Selection.End(xlDown).Row = "65536" Or Selection.End(xlDown).Row = "1048576" Then
    GoTo finish:
End If
endrange = Selection.End(xlDown).Row
For i = 3 To endrange
    offsetCount = 0
    testCell = ActiveCell
    For j = endrange - i To 0 Step -1
        offsetCount = offsetCount + 1
        ActiveCell.Offset(1, 0).Select
        If ActiveCell = testCell Then
            ActiveCell.Offset(0, -1) = Chr$(34)
        End If
    Next j
    ActiveCell.Offset(-offsetCount + 1, 0).Select
Next i
finish:
End Sub

Sub TimeFormat(testrange)


'Adds a colon between each number so Excel can read it as a time format.
On Error Resume Next
Dim rCell As Range
    Dim rRng As Range
    Range(testrange).Select
    
    Set rRng = Range(Selection, Selection.End(xlDown))

    For Each rCell In rRng.Cells
       
       rCell = (Left(rCell, Len(rCell) - 2) & ":" & Right(rCell, 2) & ":00")
       
        If Left(rCell, 1) = ":" Then rCell = "00" & rCell

        If Len(rCell) = 1 Then rCell = "00:0" & rCell
       
    Next rCell

'Changes the the time format to H:MM AM/PM.
   
    ActiveCell.EntireColumn.NumberFormat = "h:mm AM/PM"
        
'Removes the formula.
        
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

On Error GoTo 0

End Sub
