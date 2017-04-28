Attribute VB_Name = "CoreManifestAlternate"
 Sub Core_Manifest_Format_Alternate()

'Adapted by Michael Welch from the original created by Yeong Cheng and Alex Rawlings.
'Adapted 2/5/16 - Please contact Michael if you have any questions or issues.
'This macro will automatically format Savoya Manifests straight from download.



On Error GoTo initialErrorHandler:

'Prepares the logo for printing.
Dim logoPath As String
logoPath = "P:\Operations\Group Department\Information\Training\Macros\savoya_logo2.jpg"

pExists = True

'Delete Origin Column

'Range("N1").Select

'ActiveCell.EntireColumn.Delete Shift:=xlLeft

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

'Changes Binary Code to Yes and No

Range("R:R").Replace "1", "Yes"
Range("R:R").Replace "0", "No"
Range("S:S").Replace "1", "Yes"
Range("S:S").Replace "0", "No"
Range("T:T").Replace "1", "Yes"
Range("T:T").Replace "0", "No"

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


' Transfers arrival trips to Arrivals Sheet

Dim i As Integer
Dim a As Integer
Dim d As Integer
a = 1
d = 1

    endrange = Selection.End(xlDown).Row

    For i = 2 To endrange
        
        If Cells(i, 1).Value = "Arrival" Then
        Cells(i, 1).Activate
        ActiveCell.EntireRow.Cut destination:=Sheets("Arrivals").Cells(a, 1)
        a = a + 1
                    
                    
        ElseIf Cells(i, 1).Value = "Departure" Then
        Cells(i, 1).Activate
        ActiveCell.EntireRow.Cut destination:=Sheets("Departures").Cells(d, 1)
        d = d + 1
        
        End If
    
    Next i
    
  

On Error Resume Next

    Range("A1:A5000").Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete

Range("A2:A5000").Select
Range("A1").Select
ActiveCell.EntireRow.Delete

'***********Arrivals**********


Sheets("Arrivals").Select

If Range("A1") <> "Arrival" Then

    MsgBox ("No arrivals trips, deleting arrivals page")

    ActiveSheet.Delete

    GoTo skipArrivals:

End If


    
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

'Change 1 and 0 to yes and no

    Selection.Replace What:="1", Replacement:="Yes", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="0", Replacement:="No", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False



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

HighLight

'Adds qoutes for multiple passenger vehicles and deletes all blank columns

Dim p As Integer
Dim rangeend2 As Integer

Cells(3, 12).Select
rangeend2 = Selection.End(xlDown).Row

For p = 3 To rangeend2

If Cells(p, 12).Value = Cells(p - 1, 12).Value Then Cells(p, 13).Value = """"

Next p




DeleteBlanks ("Q2") 'Vendor
DeleteYesNo ("P2") 'Shuttle
DeleteYesNo ("O2") 'VIP
DeleteYesNo ("N2") 'HCP
DeleteBlanks ("K2") 'Email
DeleteBlanks ("J2") 'Phone
DeleteBlanks ("I2") 'Guests





skipArrivals:

'**********Departures***********

Sheets("Departures").Select

If Range("A1") <> "Departure" Then

    MsgBox ("No departure trips, deleting departures page")

    ActiveSheet.Delete

GoTo skipDepartures:

End If


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

HighLight

'Adds qoutes for multiple passenger vehicles and deletes all blank columns

Dim o As Integer
Dim rangeend As Integer

Cells(3, 1).Select
rangeend = Selection.End(xlDown).Row

For o = 3 To rangeend

If Cells(o, 14).Value = Cells(o - 1, 14).Value Then Cells(o, 15).Value = """"

Next o

'Replaces duplicate vehicles with quotes

DeleteBlanks ("S2") 'Vendor
DeleteYesNo ("R2") 'Shuttle
DeleteYesNo ("Q2") 'VIP
DeleteYesNo ("P2") 'HCP
DeleteBlanks ("M2") 'Email
DeleteBlanks ("L2") 'Phone
DeleteBlanks ("K2") 'Guests

skipDepartures:

'**********Offsite**********
Sheets("Offsites").Select
If Range("A1") <> "offsite" Then

    MsgBox ("No offsite trips, deleting offsites page")
    ActiveSheet.Delete
    
    GoTo skipOffsite:

End If


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
    
    Range("A1:T1").Select
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

HighLight
   
'Replaces duplicate vehicles with quotes
   
Dim y As Integer
Dim rangeend3 As Integer

Cells(3, 12).Select
rangeend3 = Selection.End(xlDown).Row

For y = 3 To rangeend3

If Cells(y, 15).Value = Cells(y - 1, 15).Value Then Cells(y, 16).Value = """"

Next y
   
   
'Deletes all blank columns

DeleteBlanks ("T2") 'Vendor
DeleteYesNo ("S2") 'Shuttle
DeleteYesNo ("R2") 'VIP
DeleteYesNo ("Q2") 'HCP
DeleteBlanks ("N2") 'Passenger Email
DeleteBlanks ("M2") 'Passenger Number
DeleteBlanks ("L2") 'Guest
DeleteBlanks ("I2") 'Flight Number
DeleteBlanks ("H2") 'Airline
DeleteBlanks ("F2") 'Flight Time
DeleteBlanks ("E2") 'Flight Date

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

Sub DeleteBlanks(testrange)
Dim ColNeeded As Boolean
Dim endCol As Integer
Dim i As Integer


ColNeeded = False
Range(testrange).Select
Range("a2", Selection).Select

endCol = Selection.Columns.Count

Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
endRow = Selection.Rows.Count

    For i = 3 To endRow
    
        If Cells(i, endCol).Value <> 0 Then ColNeeded = True
        
    Next i

Cells(1, endCol).Activate


If ColNeeded = False Then ActiveCell.EntireColumn.Delete

End Sub

Sub DeleteYesNo(testrange)
Dim ColNeeded As Boolean
Dim endCol As Integer
Dim i As Integer


ColNeeded = False
Range(testrange).Select
Range("a2", Selection).Select

endCol = Selection.Columns.Count

Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
endRow = Selection.Rows.Count

    For i = 3 To endRow
    
        If Cells(i, endCol).Value = "Yes" Then ColNeeded = True
        
    Next i

Cells(1, endCol).Activate


If ColNeeded = False Then ActiveCell.EntireColumn.Delete

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

Sub HighLight()
Dim endcolumn As Integer
Dim h As Integer
Dim hcell As Range
Dim hcell2 As Range

Range("A1").Select
'ec = Cells("A1").End(xlRight)
'Selection.End(xlToRight).Select
endcolumn = Range(Selection, Selection.End(xlToRight)).Columns.Count

'.Columns.Count

h = 3
Do
Set hcell = Cells(h, 1)
Set h2cell = Cells(h, endcolumn)

Range(hcell, h2cell).Select
Selection.Interior.color = RGB(213, 232, 255)

h = h + 2

Loop While Not IsEmpty(Cells(h, 1))

End Sub


