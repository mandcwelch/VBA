Attribute VB_Name = "GW"
Sub GW_Schedule()
Dim endRow As Integer
Dim i As Integer

i = 2

Nextrun:

If Cells(i, 20) = "" Then GoTo Endrun

If Left(Cells(i, 20), 2) <> "GW" Then

Cells(i, 20).EntireRow.Delete

GoTo Nextrun

End If

i = i + 1

GoTo Nextrun
Endrun:

Columns("Q:S").Delete
Columns("K:K").Delete
Columns("B:F").Delete

Columns("K:K").Cut
Columns("A:A").Insert

ActiveSheet.name = "New York"
Sheets.Add.name = "Dallas"
Sheets.Add.name = "Boston"
Sheets.Add.name = "D.C."
Sheets.Add.name = "San Diego"
Sheets.Add.name = "San Francisco"
Sheets.Add.name = "Austin"
Sheets.Add.name = "Houston"
Sheets.Add.name = "Los Angeles"
Sheets.Add.name = "London"
Sheets.Add.name = "Denver"
Sheets.Add.name = "Seattle"


Dim DAL As Integer
Dim DC As Integer
Dim SNA As Integer
Dim SAN As Integer
Dim AUS As Integer
Dim HOU As Integer
Dim LOS As Integer
Dim LON As Integer
Dim DEN As Integer
Dim SEA As Integer
Dim BOS As Integer

DAL = 2
DC = 2
SNA = 2
SAN = 2
AUS = 2
HOU = 2
LOS = 2
LON = 2
DEN = 2
SEA = 2
BOS = 2

Sheets("New York").Select
endRow = Cells(1, 1).End(xlDown).Row

'Dallas Sheet

For i = 2 To endRow

If Cells(i, 1) = "GW Chris Countryman DAL" Or Cells(i, 1) = "GW Edd Holt, DAL/LAS" _
Or Cells(i, 1) = "GW Matthew Dusek, DAL/LAS" Or Cells(i, 1) = "GW Joel Hoover, DAL/LAS" _
Or Cells(i, 1) = "GW Cory Countryman DAL/LAS" Or Cells(i, 1) = "GW Tyler Young DAL" Then

Cells(i, 1).EntireRow.Cut
Sheets("Dallas").Cells(DAL, 1).Insert
DAL = DAL + 1

End If

'DC Sheet

If Cells(i, 1) = "GW John McCarthy, DC" Then

Cells(i, 1).EntireRow.Cut
Sheets("D.C.").Cells(DC, 1).Insert
DC = DC + 1

End If

'SNA Sheet

'If Cells(i, 1) = "GW John McCarthy, DC" Then

'Cells(i, 1).EntireRow.Cut
'Sheets("D.C.").Cells(DC, 1).Insert
'DC = DC + 1

'End If

'SAN Sheet

If Cells(i, 1) = "GW Javier Roca" Then

Cells(i, 1).EntireRow.Cut
Sheets("San Francisco").Cells(SAN, 1).Insert
SAN = SAN + 1

End If

'Austin Sheet

If Cells(i, 1) = "GW Ben Frazier AUS/SAT" Then

Cells(i, 1).EntireRow.Cut
Sheets("Austin").Cells(AUS, 1).Insert
AUS = AUS + 1

End If

'Boston Sheet

If Cells(i, 1) = "GW (Mike) Michael Domnarski, BOS" Then

Cells(i, 1).EntireRow.Cut
Sheets("Boston").Cells(BOS, 1).Insert
BOS = BOS + 1

End If

'Houston Sheet

If Cells(i, 1) = "GW Ryan Thomason HOU" Then

Cells(i, 1).EntireRow.Cut
Sheets("Houston").Cells(HOU, 1).Insert
HOU = HOU + 1

End If

'Los Angeles

If Cells(i, 1) = "GW Steve Pullin LAX" Or Cells(i, 1) = "GW Chris Ring SAN/SNA" Then

Cells(i, 1).EntireRow.Cut
Sheets("Los Angeles").Cells(LOS, 1).Insert
LOS = LOS + 1

End If

'London Sheet

If Cells(i, 1) = "GW Vladimir Perevezentsev LHR" Or Cells(i, 1) = "GW Guy Batchelor LHR" Then

Cells(i, 1).EntireRow.Cut
Sheets("London").Cells(LON, 1).Insert
LON = LON + 1

End If

Next i

 Dim wrksht As Worksheet
    For Each wrksht In Worksheets
         
        wrksht.Select
        Cells(1, 1) = "Agent"
Cells(1, 2) = "Reservation"
Cells(1, 3) = "Date"
Cells(1, 4) = "Time"
Cells(1, 5) = "Time (CDT)"
Cells(1, 6) = "Passenger"
Cells(1, 7) = "PAX Count"
Cells(1, 8) = "Pickup Location"
Cells(1, 9) = "Drop Off Location"
Cells(1, 9) = "Arrival Airline"
Cells(1, 10) = "Departure Airline"
Cells(1, 11) = "Vehicle"
Cells(1, 12) = "Type"
Cells(1, 12) = "Notes"
        Columns.AutoFit
         
    Next wrksht

Range((Cells(1, 1)), (Cells(endRow, 1))).SpecialCells(xlCellTypeBlanks).EntireRow.Delete

End Sub
