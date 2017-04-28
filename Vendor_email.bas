Attribute VB_Name = "Vendor_email"
Sub vendor_table()

vendor = Range("A1")

Range("A3") = "Report for " & vendor

Range("A4") = "Service Deviation"
Range("A5") = "Trips Fully Managed on core.savoya.net"
Range("A6") = "Driver Assignerd 6+ hours before trip"
Range("A7") = "Driver App Used"
Range("A8") = "Trips Billed Within 24 Hours"
Range("A9") = "Trips Auto-Closed/Auto-Billed"

Range("B3") = "Number"
Range("C3") = "Total Trips"
Range("D3") = "Percentage"
Range("E3") = "Savoya Goal"
Range("F3") = "Last Month"
Range("G3") = "2017 Total"
Range("H3") = "2016 Total"


Range("E4") = ".5%"
Range("E5") = "95%"
Range("E6") = "90%"
Range("E7") = "90%"
Range("E8") = "100%"
Range("E9") = "0%"

Range("B4") = Range("H1")
Range("B5") = Range("K1")
Range("B6") = Range("L1")
Range("B7") = Range("J1")
Range("B8") = Range("F1")
Range("B9") = Range("G1")

Rows("1:2").Delete

Range(Cells(1, 1), Cells(7, 8)).Borders.Weight = 3
Columns.AutoFit

End Sub
