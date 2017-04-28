Attribute VB_Name = "Productivity"
Sub Prod_Report()
Dim endRow As Integer

Columns("A:A").Delete

ActiveSheet.name = "Main"
Range("A1") = "Call Type"
Range("B1") = "Agent"
Range("C1") = "Call Time"
Range("D1") = "Ring Time"
Range("E1") = "Call Total"

endRow = Cells(1, 1).End(xlDown).Row

For i = 2 To endRow

Cells(i, 5) = 1

Next i

' Delete blank rows

Range(Cells(1, 2), Cells(endRow + 1, 2)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
endRow = Cells(1, 1).End(xlDown).Row
' creates table

Sheets.Add.name = " Productivity for " & Format(Date, "mmmm, d yyy")

Range("A1") = "Call Type"
Range("B1") = "Agent"
Range("C1") = "Call Time"
Range("D1") = "Ring Time"
Range("E1") = "Call Total"

Sheets.Add.name = "Outbound"

Range("A1") = "Call Type"
Range("B1") = "Agent"
Range("C1") = "Call Time"
Range("D1") = "Ring Time"
Range("E1") = "Call Total"

Sheets.Add.name = "Inbound"

Range("A1") = "Call Type"
Range("B1") = "Agent"
Range("C1") = "Call Time"
Range("D1") = "Ring Time"
Range("E1") = "Call Total"

Sheets("Main").Activate

' Converts time to minutes

For i = 2 To endRow

Cells(i, 3).Value = Cells(i, 3).Value * 60
Cells(i, 4).Value = Cells(i, 4).Value * 60

Next i

'Seperates outbound calls
out = 2

For i = 2 To endRow

If Cells(i, 1) = "Dialout" Then
Cells(i, 1).EntireRow.Copy
Sheets("Outbound").Cells(out, 1).Insert
out = out + 1

End If

Next i

'Seperates inbound calls
inb = 2

For i = 2 To endRow

If Cells(i, 1) = "Inbound" Then
Cells(i, 1).EntireRow.Copy
Sheets("Inbound").Cells(inb, 1).Insert
inb = inb + 1

End If

Next i

End Sub
