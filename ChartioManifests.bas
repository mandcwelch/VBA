Attribute VB_Name = "ChartioManifests"
Sub ChartioManifest()

Dim strsearch As String
Dim lastline As Long, toCopy As Long
Dim searchColumn As String
Dim i As Long, j As Long, K As Long
Dim c As Range
Dim Header() As Variant
Dim ws As Worksheet
Dim wb As Workbook
Dim logoPath As String

ActiveSheet.name = "Manifest"
    Sheets.Add.name = "Offsite"
    Sheets.Add.name = "Departures"
    Sheets.Add.name = "Arrivals"
    
Set ws = Worksheets("Arrivals")
ws.Range("A1:N1") = Array("First Name", "Last Name", "VIP", "Date", "Time", "Airport", _
    "Airline", "Flight", "Origin", "Hotel", "Vehicle", "Confirmation", "Passenger Phone", "Passenger Email")
    Columns.AutoFit

Set ws = Worksheets("Departures")
ws.Range("A1:N1") = Array("First Name", "Last Name", "VIP", "Date", "Hotel Pickup Time", "Flight Departure Time", _
    "Hotel", "Airport", "Airline", "Flight", "Vehicle", "Confirmation", "Passenger Phone", "Passenger Email")
    Columns.AutoFit

Set ws = Worksheets("Offsite")
ws.Range("A1:I1") = Array("First Name", "Last Name", "Date", "Pickup Time", "Pickup Location", "Extra Stops", _
    "Drop Location", "Vehicle", "Confirmation")
    Columns.AutoFit
    
'Copy from Manifest to different tabs
Sheets("Manifest").Select


lastline = Range("B" & Rows.Count).End(xlUp).Row
j = 1
K = 2
For i = 1 To lastline
If Range("B" & i).Value = "ArrivalSegment" Then
   Range("C" & i & ",D" & i & ",E" & i).Copy destination:=Sheets("Arrivals").Range("A" & K)
   Range("G" & i).Copy destination:=Sheets("Arrivals").Range("D" & K)
   Range("J" & i & ",K" & i & ",L" & i & ",M" & i & ",N" & i & ",O" & i & ",P" & i & ",Q" & i & ",R" & i).Copy destination:=Sheets("Arrivals").Range("F" & K)
j = j + 1
K = K + 1
End If
Next

lastline = Range("B" & Rows.Count).End(xlUp).Row
j = 1
K = 2
For i = 1 To lastline
If Range("B" & i).Value = "DepartureSegment" Then
   Range("C" & i & ",D" & i & ",E" & i).Copy destination:=Sheets("Departures").Range("A" & K)
   Range("G" & i & ",H" & i & ",I" & i & ",J" & i).Copy destination:=Sheets("Departures").Range("D" & K)
   Range("N" & i).Copy destination:=Sheets("Departures").Range("H" & K)
   Range("K" & i).Copy destination:=Sheets("Departures").Range("I" & K)
   Range("L" & i).Copy destination:=Sheets("Departures").Range("J" & K)
   Range("O" & i & ",P" & i & ",Q" & i & ",R" & i).Copy destination:=Sheets("Departures").Range("K" & K)
j = j + 1
K = K + 1
End If
Next

lastline = Range("B" & Rows.Count).End(xlUp).Row
j = 1
K = 2
For i = 1 To lastline
If Range("B" & i).Value = "PointToPointSegment" Then
    Range("C" & i & ",D" & i).Copy destination:=Sheets("Arrivals").Range("A" & K)
    Range("G" & i & ",H" & i).Copy destination:=Sheets("Arrivals").Range("C" & K)
    Range("E" & i).Copy destination:=Sheets("Arrivals").Range("J" & K)
    Range("G" & i & ",H" & i & ",I" & i).Copy destination:=Sheets("Departures").Range("N" & K)
j = j + 1
K = K + 1
End If
Next


logoPath = "P:\Operations\Group Department\Information\Training\Macros\savoya_logo2.jpg"
Application.ScreenUpdating = False
GroupID = InputBox("Enter the GroupID for this Manifest")

Sheets("Arrivals").Select
Cells.Select
Cells.EntireColumn.AutoFit
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
    
Sheets("Departures").Select
Cells.Select
Cells.EntireColumn.AutoFit
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

Sheets("Offsite").Select
Cells.Select
Cells.EntireColumn.AutoFit
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

End Sub


