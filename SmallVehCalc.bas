Attribute VB_Name = "SmallVehCalc"
Sub smallveh()

ActiveSheet.name = "Arrivals"
Sheets.Add.name = "Departures"

Range("A1") = "Small Vehicle Calculator - Departures"
Range("A2") = "Passenger Count"
Range("B2").Interior.color = RGB(250, 250, 0)

Range("A5") = "Sedan"
Range("A6") = "SUV"
Range("A7") = "Van"

Range("B4") = "Pax/Car"
Range("B5") = 1.5
Range("B6") = 3.5
Range("B7") = 6

Range("C4") = "Vehicle Breakdown"
Range("C5") = "20%"
Range("C6") = "40%"
Range("C7") = "40%"

Range("D4") = "Transfers"
Range("D5") = "=roundup((B2/B5)*C5,0)"
Range("D6") = "=roundup((B2/B6)*C6,0)"
Range("D7") = "=roundup((B2/B7)*C7,0)"

Columns.AutoFit

Sheets("Arrivals").Activate

Range("A1") = "Small Vehicle Calculator - Arrivals"

Range("A2") = "Passenger Count"
Range("B2").Interior.color = RGB(250, 250, 0)

Range("A5") = "Sedan"
Range("A6") = "SUV"
Range("A7") = "Van"

Range("B4") = "Pax/Car"
Range("B5") = 1.5
Range("B6") = 3.5
Range("B7") = 6

Range("C4") = "Vehicle Breakdown"
Range("C5") = "65%"
Range("C6") = "20%"
Range("C7") = "15%"

Range("D4") = "Transfers"
Range("D5") = "=roundup((B2/B5)*C5,0)"
Range("D6") = "=roundup((B2/B6)*C6,0)"
Range("D7") = "=roundup((B2/B7)*C7,0)"

Columns.AutoFit

End Sub
