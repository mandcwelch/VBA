Attribute VB_Name = "GroupCont"
Sub groupdate()
'seperates by group and date

Dim first, firstdate, firstend As Range
Dim second, seconddate, secondend As Range
Dim third, thirddate, thirdend As Range
Dim forth, fourthdate, forthend As Range
Dim fifth, fifthdate, fifthend As Range
Dim sixth, sixthdate, sixthend As Range
Dim seventh, seventhdate, seventhend As Range
Dim eighth, eighthdate, eighthend As Range
Dim ninth, ninthdate, ninthend As Range
Dim tenth, tenthdate, tenthend As Range
'assign -date values

Set firstdate = Cells(2, 1)
firstdate.Select

Selection.End(xlDown).Offset(2, 0).Select
Set seconddate = Selection

Selection.End(xlDown).Offset(2, 0).Select
Set thirddate = Selection

Selection.End(xlDown).Offset(2, 0).Select
Set fourthdate = Selection

Selection.End(xlDown).Offset(2, 0).Select
Set fifthdate = Selection

Selection.End(xlDown).Offset(2, 0).Select
Set sixthdate = Selection

Selection.End(xlDown).Offset(2, 0).Select
Set seventhdate = Selection

Selection.End(xlDown).Offset(2, 0).Select
Set eighthdate = Selection

Selection.End(xlDown).Offset(2, 0).Select
Set ninthdate = Selection

Selection.End(xlDown).Offset(2, 0).Select
Set tenthdate = Selection

'determines group existence
firsttrue = False
secondtrue = False
thirdtrue = False
fourthtrue = False
fifthtrue = False
sixthtrue = False
seventhtrue = False
eighthtrue = False
ninthtrue = False
tenthtrue = False

If firstdate.Offset(0, 34) <> "" Then firsttrue = True
If seconddate.Offset(0, 34) <> "" Then secondtrue = True
If thirddate.Offset(0, 34) <> "" Then thirdtrue = True
If fourthdate.Offset(0, 34) <> "" Then fourthtrue = True
If fifthdate.Offset(0, 34) <> "" Then fifthtrue = True
If sixthdate.Offset(0, 34) <> "" Then sixthtrue = True
If seventhdate.Offset(0, 34) <> "" Then seventhtrue = True
If eighthdate.Offset(0, 34) <> "" Then eighthtrue = True
If ninthdate.Offset(0, 34) <> "" Then ninthtrue = True
If tenthdate.Offset(0, 34) <> "" Then tenthtrue = True


firstdate.Select
Set firstend = Selection.End(xlDown).Offset(0, 38)
Set first = Range(firstdate, firstend)

seconddate.Select
Set secondend = Selection.End(xlDown).Offset(0, 38)
Set second = Range(seconddate, secondend)

thirddate.Select
Set thirdend = Selection.End(xlDown).Offset(0, 38)
Set third = Range(thirddate, thirdend)

fourthdate.Select
Set fourthend = Selection.End(xlDown).Offset(0, 38)
Set fourth = Range(fourthdate, fourthend)

End Sub

