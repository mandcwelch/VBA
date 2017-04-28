Attribute VB_Name = "FormatManifest"
Sub FormatManifest()
Attribute FormatManifest.VB_ProcData.VB_Invoke_Func = " \n14"

'FormatManifest Macro v8 originally by Yeong Cheng, updated by Alex Rawlings
'This macro will automatically format Savoya Manifests straight from download.
'Only requirement is that there are no blank rows at the top of the rez. This is only an issue
'if there are neither arrival nor departure rezzes.

'v8 - added compatibility for manifests including pax billing notes column 06.28.11
'v7 - added compatibility for manifests including HCP column 03.09.11
'v6 - added error message to check correct save format, will now automatically resize
'     for offsite only programs -yc 11.18.09
'v5 - remove tracking mechanism as it isn't being used
'v4 - fit to 1 page by width, added page# to footer



On Error GoTo initialErrorHandler:
Dim logoPath As String
logoPath = "P:\Operations\Group Department\Information\Training\Macros\savoya_logo2.jpg"
'Dim groupID As Long
pExists = True

'If pExists = False Then GoTo Finish:
'Export manifest info to P: drive tracking sheet

On Error GoTo namingErrorHandler:
Sheets(1).name = "Offsite"
Sheets.Add.name = "Departures"
Sheets.Add.name = "Arrivals"
On Error GoTo 0


GroupID = InputBox("Enter GroupID")
'If groupID = "0" Then
'    MsgBox ("GroupID not provided. Please try again")
'    GoTo Finish:
'End If

'clientName = InputBox("Enter Client Name, ex: 'Pfizer' or 'Harley Davidson'")
'If clientName = "" Or clientName = "0" Then
'    MsgBox ("Client Name not provided. Please try again")
'    GoTo Finish:
'End If

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


'If Not InStr(clientName, "Pfizer") = 0 Or Not InStr(clientName, "pfizer") = 0 Then
If MsgBox("Show vehicle type for each passenger?", vbYesNo) = vbYes Then skipQuotes = True
'skipQuotes = True
'End If


Application.ScreenUpdating = False

'*********Format the 3 pages***********


Sheets("Offsite").Select
Range("A1").Select

'Check if group is offsite only
If Range("A1") = "" Then
    ActiveCell.EntireRow.Delete
    ActiveCell.EntireRow.Delete
    ActiveCell.EntireRow.Delete
End If

' Arrivals
If Range("M1") = " arr. date" Then
    arrExists = True
    endrange = Selection.End(xlDown).Row
    totalArr = endrange - 1
    For i = 1 To endrange
        ActiveCell.EntireRow.Cut destination:=Sheets("Arrivals").Cells(i, 1)
        ActiveCell.EntireRow.Delete
    Next i
    ActiveCell.EntireRow.Delete
    ActiveCell.EntireRow.Delete
    ActiveCell.EntireRow.Delete
Else
    MsgBox ("No Arrivals, deleting Arrivals Page")
    Sheets("Arrivals").Delete
End If

'Departures
If Range("M1") = " dep. date" Then
    depExists = True
    endrange = Selection.End(xlDown).Row
    totalDep = endrange - 1
    For i = 1 To endrange
        ActiveCell.EntireRow.Cut destination:=Sheets("Departures").Cells(i, 1)
        ActiveCell.EntireRow.Delete
    Next i
    ActiveCell.EntireRow.Delete
    ActiveCell.EntireRow.Delete
    ActiveCell.EntireRow.Delete
Else
    MsgBox ("No departure trips, deleting departures page")
    Sheets("Departures").Delete
End If

'Offsites
If ActiveCell = " rez id" Then
    offExists = True
    totalOff = Selection.End(xlDown).Row - 1
Else
    MsgBox ("No offsite trips, deleting offsite page")
    ActiveSheet.Delete
End If
'***********Arrivals**********
If arrExists = False Then GoTo skipArrivals:
    Sheets("Arrivals").Select
    Columns("D:D").Cut
    Columns("V:V").Select
    Selection.Insert Shift:=xlToRight
    Columns("F:J").Cut
    Columns("V:V").Select
    Selection.Insert Shift:=xlToRight
    Columns("A:A").Cut
    Columns("P:P").Select
    Selection.Insert Shift:=xlToRight

    Range("A1") = "First Name"
    Range("B1") = "Last Name"
    Range("C1") = "VIP"
    Range("D1") = "HCP"
    Range("E1") = "Guests"
    Range("F1") = "Date"
    Range("G1") = "Time"
    Range("H1") = "Airport"
    Range("I1") = "Airline"
    Range("J1") = "Flight"
    Range("K1") = "Origin"
    Range("L1") = "Hotel"
    Range("M1") = "Notes"
    Range("N1") = "Vehicle"
    Range("O1") = "Confirmation"
    Range("P1") = "Passenger Billing Code"
    Range("Q1") = "Passenger Phone"
    Range("R1") = "Passenger Email"
    Range("S1") = "Contact Name"
    Range("T1") = "Contact Phone"
    Range("U1") = "Contact Email"
    Columns.AutoFit
    Columns("A:A").ColumnWidth = 11
    Columns("B:B").ColumnWidth = 13
    Columns("C:C").ColumnWidth = 9
    Columns("D:D").ColumnWidth = 4
    Columns("E:E").ColumnWidth = 4
    Columns("F:F").ColumnWidth = 10
    Columns("G:G").ColumnWidth = 8
    Columns("H:H").ColumnWidth = 12
    Columns("I:I").ColumnWidth = 10
    Columns("J:J").ColumnWidth = 6.5
    Columns("K:K").ColumnWidth = 6.5
    Columns("L:L").ColumnWidth = 15
    Columns("M:M").ColumnWidth = 8
    Columns("O:O").ColumnWidth = 12
    Columns("N:N").HorizontalAlignment = xlCenter
    Range("N1").HorizontalAlignment = xlLeft
    Columns("J:J").HorizontalAlignment = xlCenter
    Range("J1").HorizontalAlignment = xlLeft
    
    Range("A1:U1").Select
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
    Columns("H:H").Insert Shift:=xlToRight
    Columns("G:G").Replace What:="AM ", Replacement:="AM-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("G:G").Replace What:="PM", Replacement:="PM-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("G3:G" & Range("G65536").End(xlUp).Row).TextToColumns destination:=Range("G3"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("H:H").Delete Shift:=xlToLeft
    ActiveSheet.UsedRange.Sort Key1:=Range("F3"), order1:=xlAscending, key2:=Range( _
        "G3"), order2:=xlAscending, key3:=Range("O3"), Order3:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal

If skipQuotes = True Then GoTo skipArrQuotes:
Quotes ("O3")
skipArrQuotes:


skipArrivals:

'**********Departures***********
If depExists = False Then GoTo skipDepartures:
Sheets("Departures").Select
    Columns("D:D").Cut
    Columns("X:X").Select
    Selection.Insert Shift:=xlToRight
    Columns("F:J").Cut
    Columns("X:X").Select
    Selection.Insert Shift:=xlToRight
    Columns("A:A").Cut
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight
    Columns("H:H").Delete
    Columns("M:M").Delete

    Range("A1") = "First Name"
    Range("B1") = "Last Name"
    Range("C1") = "VIP"
    Range("D1") = "HCP"
    Range("E1") = "Guests"
    Range("F1") = "Date"
    Range("G1") = "Hotel Pickup Time"
    Range("H1") = "Flight Departure Time"
    Range("I1") = "Hotel"
    Range("J1") = "Airport"
    Range("K1") = "Airline"
    Range("L1") = "Flight"
    Range("M1") = "Notes"
    Range("N1") = "Vehicle"
    Range("O1") = "Confirmation"
    Range("P1") = "Passenger Billing Code"
    Range("Q1") = "Passenger Phone"
    Range("R1") = "Passenger Email"
    Range("S1") = "Contact Name"
    Range("T1") = "Contact Phone"
    Range("U1") = "Contact Email"


    Columns.AutoFit
    Columns("A:A").ColumnWidth = 11
    Columns("B:B").ColumnWidth = 13
    Columns("C:C").ColumnWidth = 9
    Columns("E:E").ColumnWidth = 3
    Columns("F:F").ColumnWidth = 10
    Columns("G:G").ColumnWidth = 16
    Columns("H:H").ColumnWidth = 18
    Columns("I:I").ColumnWidth = 10
    Columns("J:J").ColumnWidth = 12
    Columns("K:K").ColumnWidth = 10
    Columns("M:M").ColumnWidth = 14
    Columns("N:N").ColumnWidth = 10
    Columns("O:O").ColumnWidth = 12
    Columns("N:N").HorizontalAlignment = xlCenter
    Range("N1").HorizontalAlignment = xlLeft
    Columns("L:L").HorizontalAlignment = xlCenter
    Range("L1").HorizontalAlignment = xlLeft
    
    Range("A1:U1").Select
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
    Columns("H:H").Insert Shift:=xlToRight
    Columns("G:G").Replace What:="AM ", Replacement:="AM-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("G:G").Replace What:="PM", Replacement:="PM-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("G3:G" & Range("G65536").End(xlUp).Row).TextToColumns destination:=Range("G3"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("H:H").Delete Shift:=xlToLeft
   
    Columns("I:I").Insert Shift:=xlToRight
    Columns("H:H").Replace What:="AM ", Replacement:="AM-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("H:H").Replace What:="PM", Replacement:="PM-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("H3:H" & Range("H65536").End(xlUp).Row).TextToColumns destination:=Range("H3"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("I:I").Delete Shift:=xlToLeft
  
    ActiveSheet.UsedRange.Sort Key1:=Range("F3"), order1:=xlAscending, key2:=Range( _
        "G3"), order2:=xlAscending, key3:=Range("O3"), Order3:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal

If skipQuotes = True Then GoTo skipDepQuotes:
Quotes ("O3")
skipDepQuotes:



skipDepartures:

'**********Offsite**********
If offExists = False Then GoTo skipOffsite:
    Sheets("Offsite").Select
    Columns("D:D").Cut
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:J").Cut
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight
    Columns("A:A").Select
    Selection.Cut
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight

    Range("A1") = "First Name"
    Range("B1") = "Last Name"
    Range("C1") = "VIP"
    Range("D1") = "HCP"
    Range("E1") = "Passenger Phone"
    Range("F1") = "Guests"
    Range("G1") = "Trip Type"
    Range("H1") = "Date"
    Range("I1") = "Pickup Time"
    Range("J1") = "Pickup Location"
    Range("K1") = "Pickup Instructions"
    Range("L1") = "Flight"
    Range("M1") = "Drop Location"
    Range("N1") = "Drop Instructions"
    Range("O1") = "Extra Stops"
    Range("P1") = "Vehicle"
    Range("Q1") = "Confirmation"
    Columns.AutoFit
    Columns("A:A").ColumnWidth = 11
    Columns("B:B").ColumnWidth = 13
    Columns("C:C").ColumnWidth = 8
    Columns("D:D").ColumnWidth = 14
    Columns("E:E").ColumnWidth = 14
    Columns("F:F").ColumnWidth = 14
    Columns("G:G").ColumnWidth = 14
    Columns("H:H").ColumnWidth = 12
    Columns("I:I").ColumnWidth = 12
    Columns("J:J").ColumnWidth = 12
    Columns("K:K").ColumnWidth = 14
    Columns("L:L").ColumnWidth = 12
    Columns("M:M").ColumnWidth = 12
    Columns("N:N").ColumnWidth = 14
    Columns("O:O").ColumnWidth = 14
    Columns("P:P").ColumnWidth = 12
    Columns("Q:Q").ColumnWidth = 14
    
    
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
        .RightHeader = "GroupID: " & GroupID & Chr(10) & "Offsite Manifest"
        .CenterFooter = "&D"
        .RightFooter = "&P"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
'Sort by Conf, Time, and Date
    Columns("J:J").Insert Shift:=xlToRight
    Columns("I:I").Replace What:="AM ", Replacement:="AM-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("I:I").Replace What:="PM", Replacement:="PM-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("I3:I" & Range("H65536").End(xlUp).Row).TextToColumns destination:=Range("I3"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("J:J").Delete Shift:=xlToLeft
    ActiveSheet.UsedRange.Sort Key1:=Range("H3"), order1:=xlAscending, key2:=Range( _
        "I3"), order2:=xlAscending, key3:=Range("Q3"), Order3:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal
        
If skipQuotes = True Then GoTo skipOffsiteQuotes:
skipOffsiteQuotes:



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


