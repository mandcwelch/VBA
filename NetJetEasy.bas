Attribute VB_Name = "NetJetEasy"
Sub NetJetNetPull()

Dim i As Integer
Dim dt As Date
Application.ScreenUpdating = False
Columns("H:H").Insert
Columns("H:H").Value = Columns("G:G").Value
Columns("H:H").NumberFormat = "hhmm"
Columns("G:G").NumberFormat = "m/dd/yyyy"
Cells(1, 8) = "Pickup Time"

'Removes unneeded columns



Columns("U:V").Delete
Columns("J:S").Delete
Columns("I:I").Delete
Columns("C:E").Delete

Columns.AutoFit

ActiveSheet.UsedRange.Sort Key1:=Range("E2"), order1:=xlAscending, key2:=Range( _
        "D3"), order2:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal

'Creates additional sheets

'ActiveSheet.name = "NetJets " & Format(Date, "mmmm dd") & " to " & Format(Date + 1, "mmmm dd")
'Sheets.Add.name = "NetJets " & Format(Date + 5, "mmmm dd")
'Sheets.Add.name = "NetJets " & Format(Date + 4, "mmmm dd")
'Sheets.Add.name = "NetJets " & Format(Date + 3, "mmmm dd")
'Sheets.Add.name = "NetJets " & Format(Date + 2, "mmmm dd")

'Sheets("NetJets " & Format(Date, "mmmm dd") & " to " & Format(Date + 1, "mmmm dd")).Move _
'before:=Sheets("NetJets " & Format(Date + 2, "mmmm dd"))

'Removes unneeded trips

For i = 2 To 5000

If Cells(i, 3) <> "Marquis Jet" And _
 Cells(i, 3) <> "EJM (Executive Jet Management)" And _
 Cells(i, 3) <> "NetJets" Then

 Cells(i, 3) = ""
End If

Next i

    Range("C1:C5000").Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete


'Colors same day/next days


Dim endRow As Integer

endRow = Cells(1, 1).End(xlDown).Row

Columns("D:D").NumberFormat = "general"

For i = 2 To endRow

Cells(i, 4) = Left(Cells(i, 4), 5)


Next i

For i = 2 To 5000
Cells(i, 7) = Cells(i, 4)

Next i

Columns("G:G").NumberFormat = "m/dd/yyyy"
Columns("D:D").NumberFormat = "m/dd/yyyy"

For i = 2 To 5000
If Cells(i, 7) = Date Then Cells(i, 4).Font.color = RGB(255, 0, 0)
If Cells(i, 7) = Date Then Cells(i, 4).Interior.color = RGB(255, 255, 0)
If Cells(i, 7) = Date + 1 Then Cells(i, 4).Font.color = RGB(255, 0, 0)
If Cells(i, 2) = "garage_assigned" Then Cells(i, 2).Interior.color = RGB(255, 255, 0)
If Cells(i, 2) = "garage_assigned" Then Cells(i, 2).Interior.color = RGB(255, 0, 0)
If Cells(i, 2) = "mod_pending" Then Cells(i, 2).Interior.color = RGB(255, 153, 50)

Next i

Columns("G:G").Delete


'moves later trips
'Cells(1, 1).Select
'Selection.EntireRow.Copy

'Sheets("NetJets " & Format(Date + 5, "mmmm dd")).Activate
'Range("A1") = "Rez. ID"
'Range("B1") = "Status"
'Range("C1") = "Company Name"
'Range("D1") = "Dallas time"
'Sheets("NetJets " & Format(Date + 4, "mmmm dd")).Activate
'Range("A1") = "Rez. ID"
'Range("B1") = "Status"
'Range("C1") = "Company Name"
'range("D1") = "Dallas time"
'Sheets("NetJets " & Format(Date + 3, "mmmm dd")).Activate
'Range("A1") = "Rez. ID"
'Range("B1") = "Status"
'Range("C1") = "Company Name"
'Range("D1") = "Dallas time"
'Sheets("NetJets " & Format(Date + 2, "mmmm dd")).Activate
'Range("A1") = "Rez. ID"
'Range("B1") = "Status"
'Range("C1") = "Company Name"
'Range("D1") = "Dallas time"
'Sheets("NetJets " & Format(Date, "mmmm dd") & " to " & Format(Date + 1, "mmmm dd")).Activate
       ' a = 2
        'b = 2
        'c = 2
        'd = 2
        
    'For i = 2 To 5000
        
        'If Cells(i, 4) = Format(Date + 2, "mm/dd/yyyy hhmm") Then
        'Cells(i, 1).Activate
        'ActiveCell.EntireRow.Cut Destination:=Sheets("NetJets " & Format(Date + 2, "mmmm dd")).Cells(a, 1)
        'a = a + 1
                    
        'ElseIf Cells(i, 4) = Format(Date + 3, "mm/dd/yyyy hhmm") Then
        'Cells(i, 1).Activate
        'ActiveCell.EntireRow.Cut Destination:=Sheets("NetJets " & Format(Date + 3, "mmmm dd")).Cells(b, 1)
        'b = b + 1
        
        'ElseIf Cells(i, 4) = Format(Date + 4, "mm/dd/yyyy hhmm") Then
        'Cells(i, 1).Activate
        'ActiveCell.EntireRow.Cut Destination:=Sheets("NetJets " & Format(Date + 4, "mmmm dd")).Cells(c, 1)
        'c = c + 1
        
        'ElseIf Cells(i, 4) = Format(Date + 5, "mm/dd/yyyy hhmm") Then
      '  Cells(i, 1).Activate
       ' ActiveCell.EntireRow.Cut Destination:=Sheets("NetJets " & Format(Date + 5, "mmmm dd")).Cells(d, 1)
     '   d = d + 1
        
        
    '    End If
    
   ' Next i





'For i = 1 To 5000

'If Cells(i, 4) = Date + 1 Then Cells(i, 4).Font.Color = RGB(206, 216, 66)

'Next i

endRow = Cells(1, 1).End(xlDown).Row

For i = 2 To endRow

If Mid(Cells(i, 1), 8, 1) = "]" Then
Cells(i, 1) = Mid(Cells(i, 1), 2, 6)
Else: Cells(i, 1) = Mid(Cells(i, 1), 2, 8)
End If


Next i


Columns("A:F").AutoFit


Application.ScreenUpdating = True
End Sub







