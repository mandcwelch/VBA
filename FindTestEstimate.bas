Attribute VB_Name = "FindTestEstimate"
Sub TestEstimate()

Dim fCell As Range
Dim rCell As Range
Dim endRow As Range
Dim totrange As Range

sedan = 0
suv = 0
van = 0
Coach = 0
mini = 0



Set fCell = Range(Cells(2, 1), Cells(2, 15)).Find("Vehicle")

Set endRow = fCell.End(xlDown)

Set totrange = Range(fCell, endRow)

For Each rCell In totrange.Cells

If rCell = "Sedan" Then sedan = sedan + 1
If rCell = "SUV" Then suv = suv + 1
If rCell = "Van" Then van = van + 1
If rCell = "Coach Bus" Then Coach = Coach + 1
If rCell = "Mini" Then mini = mini + 1

Next rCell

Sheets("Estimate Template").Activate



'endrow.Offset(2, -1) = "Sedans"
'endrow.Offset(3, -1) = "SUVs"
'endrow.Offset(4, -1) = "Vans"
'endrow.Offset(5, -1) = "Coaches"
'endrow.Offset(6, -1) = "Minis"

'endrow.Offset(2, 0) = Sedan
'endrow.Offset(3, 0) = SUV
'endrow.Offset(4, 0) = Van
'endrow.Offset(5, 0) = Coach
'endrow.Offset(6, 0) = Mini

Range("D26") = sedan
Range("D27") = suv
Range("D28") = van



End Sub
