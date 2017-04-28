Attribute VB_Name = "NewSystemManifestOriginal"
Sub New_System_Manifest_v2()
Attribute New_System_Manifest_v2.VB_ProcData.VB_Invoke_Func = "q\n14"

'Adapted by Michael Welch from the original created by Yeong Cheng and Alex Rawlings.
'Adapted 2/5/16 - Please contact Michael if you have any questions or issues.
'This macro will automatically format Savoya Manifests straight from download.



On Error GoTo initialErrorHandler:

'Prepares the logo for printing.
Dim logoPath As String
logoPath = "P:\Operations\Group Department\Information\Training\Macros\savoya_logo2.jpg"

pExists = True

'Delete Origin Column

Range("N1").Select

ActiveCell.EntireColumn.Delete Shift:=xlLeft

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

'Keeps or removes vendor assignment information

MSG1 = MsgBox("Is this a Vendor Manifest?", vbYesNo)

    If MSG1 = vbNo Then Columns("T:T").Delete Shift:=xlToLeft

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
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    
    

    Range("A1") = "First Name"
    Range("B1") = "Last Name"
    Range("C1") = "Shuttle"
    Range("D1") = "VIP"
    Range("E1") = "HCP"
    Range("F1") = "Flight Date"
    Range("G1") = "Flight Time"
    Range("H1") = "Pickup Location"
    Range("I1") = "Airline"
    Range("J1") = "Flight Number"
    Range("K1") = "Dropoff"
    Range("L1") = "Vehicle"
    Range("M1") = "Confirmation"
    Range("N1") = "Passenger Number"
    Range("O1") = "Passenger Email"
    Range("P1") = "Guests"
    Range("Q1") = "Vendor"
    
       Columns.AutoFit

    Columns("L:L").HorizontalAlignment = xlCenter
    Range("L1").HorizontalAlignment = xlLeft
    Columns("G:G").HorizontalAlignment = xlRight
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

    ActiveSheet.UsedRange.Sort Key1:=Range("F3"), order1:=xlAscending, key2:=Range( _
        "G3"), order2:=xlAscending, key3:=Range("M3"), Order3:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal

'Adds qoutes for multiple passenger vehicles and deletes all blank columns

Dim p As Integer
Dim rangeend2 As Integer

Cells(3, 13).Select
rangeend2 = Selection.End(xlDown).Row

For p = 3 To rangeend2

If Cells(p, 13).Value = Cells(p - 1, 13).Value Then Cells(p, 12).Value = """"

Next p

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

'Quotes ("M3")


DeleteBlanks ("Q3") 'Vendor
DeleteBlanks ("P3") 'Guests
DeleteBlanks ("O3") 'Passenger Email
DeleteBlanks ("N3") 'Passenger Number
DeleteBlanks ("E3") 'HCP
DeleteBlanks ("D3") 'VIP
DeleteBlanks ("C3") 'Shuttle


skipArrivals:

'**********Departures***********
If depExists = False Then GoTo skipDepartures:
Sheets("Departures").Select

    Range("A1").Select
    ActiveCell.EntireRow.Insert Shift:=xlDown

    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("A1") = "First Name"
    Range("B1") = "Last Name"
    Range("C1") = "Shuttle"
    Range("D1") = "VIP"
    Range("E1") = "HCP"
    Range("F1") = "Pickup Date"
    Range("G1") = "Pickup Time"
    Range("H1") = "Flight Date"
    Range("I1") = "Flight Time"
    Range("J1") = "Pickup Location"
    Range("K1") = "Airline"
    Range("L1") = "Flight Number"
    Range("M1") = "Dropoff"
    Range("N1") = "Vehicle"
    Range("O1") = "Confirmation"
    Range("P1") = "Passenger Number"
    Range("Q1") = "Passenger Email"
    Range("R1") = "Guests"
    Range("S1") = "Vendor"
    
    'Formats the header and adds logo for printing
    
        Columns.AutoFit

    Columns("N:N").HorizontalAlignment = xlCenter
    Range("N1").HorizontalAlignment = xlLeft
    Columns("i:i").HorizontalAlignment = xlRight
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
ActiveSheet.UsedRange.Sort Key1:=Range("F3"), order1:=xlAscending, key2:=Range( _
        "G3"), order2:=xlAscending, key3:=Range("O3"), Order3:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal

 Range("I3").Select
    Range(Selection, Selection.End(xlDown)).Select
    TimeFormat

'Adds qoutes for multiple passenger vehicles and deletes all blank columns

Dim o As Integer
Dim rangeend As Integer

Cells(3, 15).Select
rangeend = Selection.End(xlDown).Row

For o = 3 To endrange

If Cells(o, 15).Value = Cells(o - 1, 15).Value Then Cells(o, 14).Value = """"

Next o

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

'Quotes ("O3")

DeleteBlanks ("S3") 'Vendor
DeleteBlanks ("R3") 'Guests
DeleteBlanks ("Q3") 'Passenger Email
DeleteBlanks ("P3") 'Passenger Email
DeleteBlanks ("E3") 'HCP
DeleteBlanks ("D3") 'VIP
DeleteBlanks ("C3") 'Shuttle





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
    Range("C1") = "Shuttle"
    Range("D1") = "VIP"
    Range("E1") = "HCP"
    Range("F1") = "Pickup Date"
    Range("G1") = "Pickup Time"
    Range("H1") = "Flight Date"
    Range("I1") = "Flight Time"
    Range("J1") = "Pickup Location"
    Range("K1") = "Airline"
    Range("L1") = "Flight Number"
    Range("M1") = "Dropoff"
    Range("N1") = "Vehicle"
    Range("O1") = "Confirmation"
    Range("P1") = "Passenger Number"
    Range("Q1") = "Passenger Email"
    Range("R1") = "Guests"
    Range("S1") = "Vendor"
    
    'Formats the header and adds logo for printing
        
        Columns.AutoFit

    Columns("N:N").HorizontalAlignment = xlCenter
    Range("N1").HorizontalAlignment = xlLeft
    Columns("i:i").HorizontalAlignment = xlRight
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
ActiveSheet.UsedRange.Sort Key1:=Range("F3"), order1:=xlAscending, key2:=Range( _
        "G3"), order2:=xlAscending, key3:=Range("O3"), Order3:=xlAscending, _
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


DeleteBlanks ("S3") 'Vendor
DeleteBlanks ("R3") 'Guests
DeleteBlanks ("Q3") 'Passenger Email
DeleteBlanks ("P3") 'Passenger Email
DeleteBlanks ("L3") 'Flight Number
DeleteBlanks ("K3") 'Airline
DeleteBlanks ("I3") 'Flight Time
DeleteBlanks ("H3") 'Flight Date
DeleteBlanks ("E3") 'HCP
DeleteBlanks ("D3") 'VIP
DeleteBlanks ("C3") 'Shuttle



'Adds Highlighting

Range("A3").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=MOD(ROW(),2)=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False

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

Sub TimeFormat()


'Adds a colon between each number so Excel can read it as a time format.
On Error Resume Next
Dim rCell As Range
    Dim rRng As Range

    Set rRng = Selection

    For Each rCell In rRng.Cells
       
       rCell = (Left(rCell, Len(rCell) - 2) & ":" & Right(rCell, 2) & ":00")
       
    Next rCell

'Changes the the time format to H:MM AM/PM.
   

    Selection.NumberFormat = "h:mm AM/PM"
        
'Removes the formula.
        
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Refreshes the cells and cancels the copy.
       
Selection.TextToColumns destination:=Selection, DataType:=xlDelimited, _
        FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
On Error GoTo 0
End Sub
