Attribute VB_Name = "NetJets"
Sub NetJets()

Dim i As Integer
Dim dt As Date
'Removes unneeded columns

Columns("J:V").Delete
Columns("G:H").Delete
Columns("C:E").Delete

Columns.AutoFit

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

For i = 2 To 5000
Cells(i, 5) = Left(Cells(i, 4), 10)

Next i

For i = 2 To 5000
If Cells(i, 5) = Date Then Cells(i, 4).Font.color = RGB(256, 0, 0)
If Cells(i, 5) = Date + 1 Then Cells(i, 4).Font.color = RGB(206, 216, 66)
If Cells(i, 2) = "garage_assigned" Then Cells(i, 2).Interior.color = RGB(255, 0, 0)
If Cells(i, 2) = "mod_pending" Then Cells(i, 2).Interior.color = RGB(255, 255, 0)

Next i

Columns("E:E").Delete


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
        a = 2
        b = 2
        c = 2
        d = 2
        
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

End Sub



