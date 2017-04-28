Attribute VB_Name = "New_System_Manifest"
Sub New_System_Manifest()

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
Sheets(1).name = "Offsites"
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


Sheets("Offsites").Select

MSG1 = MsgBox("Is this a Vendor Manifest?", vbYesNo)
If MSG1 = vbNo Then Columns("U:U").Delete Shift:=xlToLeft

Range("A1").Select


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


Range("A1").Select
ActiveCell.EntireRow.Delete Shift:=xlUp
'Check if group is offsite only
'If Range("A1") = "offsite" Then
   
'End If


' Arrivals
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


'Departures
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

'Offsites



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
    Columns("M:M").Select
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
    
       Columns.AutoFit

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
        .RightHeader = "GroupID: " & GroupID & Chr(10) & "Offsite Manifest"
        .CenterFooter = "&D"
        .RightFooter = "&P"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

If skipQuotes = True Then GoTo skipArrQuotes:
Quotes ("O3")
skipArrQuotes:
DeleteBlanks ("U3") 'Contact Email
DeleteBlanks ("T3") 'Contact Phone
DeleteBlanks ("S3") 'Contact Name
DeleteBlanks ("R3") 'Passenger Email
DeleteBlanks ("Q3") 'Passenger PHone
DeleteBlanks ("P3") 'Passenger Billing Code
DeleteBlanks ("M3") 'Notes
DeleteColumns ("E3") 'Guests
DeleteBlanks ("D3") 'HCP

skipArrivals:

'**********Departures***********
If depExists = False Then GoTo skipDepartures:
Sheets("Departures").Select

    Range("A1").Select
    ActiveCell.EntireRow.Insert Shift:=xlDown

    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("M:M").Select
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
    
        Columns.AutoFit

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
        .RightHeader = "GroupID: " & GroupID & Chr(10) & "Offsite Manifest"
        .CenterFooter = "&D"
        .RightFooter = "&P"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

If skipQuotes = True Then GoTo skipDepQuotes:
Quotes ("O3")
skipDepQuotes:
DeleteBlanks ("U3") 'Contact Email
DeleteBlanks ("T3") 'Contact Phone
DeleteBlanks ("S3") 'Contact Name
DeleteBlanks ("R3") 'Passenger Email
DeleteBlanks ("Q3") 'Passenger Phone
DeleteBlanks ("P3") 'Passenger Billing Code
DeleteBlanks ("M3") 'Notes
DeleteColumns ("E3") 'Guests
DeleteBlanks ("D3") 'HCP


skipDepartures:

'**********Offsite**********
If offExists = False Then GoTo skipOffsite:
    Sheets("Offsites").Select
    
    Range("A1").Select
    ActiveCell.EntireRow.Insert Shift:=xlDown
    
       Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("M:M").Select
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
    
        Columns.AutoFit

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
    

        
If skipQuotes = True Then GoTo skipOffsiteQuotes:
skipOffsiteQuotes:
DeleteBlanks ("V3") 'Contact Email
DeleteBlanks ("U3") 'Contact Phone
DeleteBlanks ("T3") 'Contact Name
DeleteBlanks ("S3") 'Passenger Email
DeleteBlanks ("R3") 'Passenger Billing Code
DeleteBlanks ("D3") 'HCP
DeleteBlanks ("N3") 'Extra Stops
DeleteBlanks ("M3") 'Drop Instructions
DeleteBlanks ("K3") 'Flight No
DeleteBlanks ("J3") 'Pickup Instructions
DeleteColumns ("G3") 'Trip Type
DeleteColumns ("F3") 'Guests
DeleteColumns ("E3") 'Passenger Phone
DeleteColumns ("D3") 'HCP
DeleteColumns ("C3") 'VIP


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








