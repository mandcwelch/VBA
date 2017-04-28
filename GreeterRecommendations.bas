Attribute VB_Name = "GreeterRecommendations"
Sub Greeter_Recomendation()

Dim i As Integer
Dim r As Integer
Dim endRow As Integer

    ActiveSheet.UsedRange.Sort Key1:=Range("c3"), order1:=xlAscending, key2:=Range( _
        "D3"), order2:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal

ActiveSheet.name = "Manifest"



'Creates rounded time column

'Columns("G:G").Insert
'Columns("G:G").Insert


'TimeFormats2

endRow = Cells(1, 1).End(xlDown).Row

'   For i = 2 To endrow

'Cells(i, 8).Value = "=(ROUNDDOWN(((G" & i & ")* 1440)/60,0) * 60)/1440"
    
    
 '   Next i
    
  '      Columns("H:H").Copy
   ' Columns("H:H").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    '    :=False, Transpose:=False

'Columns("G:G").Delete
'Columns("G:G").NumberFormat = "hhmm"
'Cells(1, 7) = "Time Rounded"

airportfalse

For i = 1 To endRow

If Cells(i, 8) = "FLL" Then FLL = True

If Cells(i, 8) = "MIA" Then MIA = True

If Cells(i, 8) = "PHX" Then PHX = True

Next i

If FLL = True Then FLLSheet

If MIA = True Then MIASheet

If PHX = True Then PHXSheet


End Sub

Sub airportfalse()

FLL = False
MIA = False
PHX = False
JFK = False

End Sub


Sub FLLSheet()

r = 1

Sheets.Add.name = "FLL"

Sheets("Manifest").Activate
endRow = Cells(1, 1).End(xlDown).Row
For i = 2 To endRow

If Cells(i, 8) = "FLL" Then

Range(Cells(i, 1), Cells(i, 17)).Copy
Sheets("FLL").Cells(r, 1).Insert
r = r + 1
End If

Next i

End Sub

Sub MIASheet()

r = 1

Sheets.Add.name = "MIA"

Sheets("Manifest").Activate
endRow = Cells(1, 1).End(xlDown).Row
For i = 2 To endRow

If Cells(i, 8) = "MIA" Then

Range(Cells(i, 1), Cells(i, 17)).Copy
Sheets("MIA").Cells(r, 1).Insert
r = r + 1
End If

Next i
Sheets("MIA").Activate
For i = 1 To endRow

If Cells(i, 9) = "AA" Then Cells(i, 18) = "Terminal 1"
If Cells(i, 9) = "DL" Then Cells(i, 18) = "Terminal 3"
If Cells(i, 9) = "UA" Then Cells(i, 18) = "Terminal 3"

Next i

    ActiveSheet.UsedRange.Sort Key1:=Range("R1"), order1:=xlAscending, key2:=Range( _
        "C1"), order2:=xlAscending, key3:=Range("D1"), Order3:=xlAscending, _
        Header:=False, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal
        
'Terminal One Set Up
        
t = Cells(1, 1).End(xlDown).Row

Cells(t + 2, 1) = "Teminal 1"
Cells(t + 2, 2) = "12 AM"
Cells(t + 3, 2) = "1 AM"
Cells(t + 4, 2) = "2 AM"
Cells(t + 5, 2) = "3 AM"
Cells(t + 6, 2) = "4 AM"
Cells(t + 7, 2) = "5 AM"
Cells(t + 8, 2) = "6 AM"
Cells(t + 9, 2) = "7 AM"
Cells(t + 10, 2) = "8 AM"
Cells(t + 11, 2) = "9 AM"
Cells(t + 12, 2) = "10 AM"
Cells(t + 13, 2) = "11 AM"
Cells(t + 14, 2) = "12 PM"
Cells(t + 15, 2) = "1 PM"
Cells(t + 16, 2) = "2 PM"
Cells(t + 17, 2) = "3 PM"
Cells(t + 18, 2) = "4 PM"
Cells(t + 19, 2) = "5 PM"
Cells(t + 20, 2) = "6 PM"
Cells(t + 21, 2) = "7 PM"
Cells(t + 22, 2) = "8 PM"
Cells(t + 23, 2) = "9 PM"
Cells(t + 24, 2) = "10 PM"
Cells(t + 25, 2) = "11 PM"



For i = 1 To t

If Left(Cells(i, 6), 2) = "00" And Cells(i, 18) = "Terminal 1" Then time0 = time0 + 1
If Left(Cells(i, 6), 2) = "01" And Cells(i, 18) = "Terminal 1" Then time1 = time1 + 1
If Left(Cells(i, 6), 2) = "02" And Cells(i, 18) = "Terminal 1" Then time2 = time2 + 1
If Left(Cells(i, 6), 2) = "03" And Cells(i, 18) = "Terminal 1" Then time3 = time3 + 1
If Left(Cells(i, 6), 2) = "04" And Cells(i, 18) = "Terminal 1" Then time4 = time4 + 1
If Left(Cells(i, 6), 2) = "05" And Cells(i, 18) = "Terminal 1" Then time5 = time5 + 1
If Left(Cells(i, 6), 2) = "06" And Cells(i, 18) = "Terminal 1" Then time6 = time6 + 1
If Left(Cells(i, 6), 2) = "07" And Cells(i, 18) = "Terminal 1" Then time7 = time7 + 1
If Left(Cells(i, 6), 2) = "08" And Cells(i, 18) = "Terminal 1" Then time8 = time8 + 1
If Left(Cells(i, 6), 2) = "09" And Cells(i, 18) = "Terminal 1" Then time9 = time9 + 1
If Left(Cells(i, 6), 2) = "10" And Cells(i, 18) = "Terminal 1" Then time10 = time10 + 1
If Left(Cells(i, 6), 2) = "11" And Cells(i, 18) = "Terminal 1" Then time11 = time11 + 1
If Left(Cells(i, 6), 2) = "12" And Cells(i, 18) = "Terminal 1" Then time12 = time12 + 1
If Left(Cells(i, 6), 2) = "13" And Cells(i, 18) = "Terminal 1" Then time13 = time13 + 1
If Left(Cells(i, 6), 2) = "14" And Cells(i, 18) = "Terminal 1" Then time14 = time14 + 1
If Left(Cells(i, 6), 2) = "15" And Cells(i, 18) = "Terminal 1" Then time15 = time15 + 1
If Left(Cells(i, 6), 2) = "16" And Cells(i, 18) = "Terminal 1" Then time16 = time16 + 1
If Left(Cells(i, 6), 2) = "17" And Cells(i, 18) = "Terminal 1" Then time17 = time17 + 1
If Left(Cells(i, 6), 2) = "18" And Cells(i, 18) = "Terminal 1" Then time18 = time18 + 1
If Left(Cells(i, 6), 2) = "19" And Cells(i, 18) = "Terminal 1" Then time19 = time19 + 1
If Left(Cells(i, 6), 2) = "20" And Cells(i, 18) = "Terminal 1" Then time20 = time20 + 1
If Left(Cells(i, 6), 2) = "21" And Cells(i, 18) = "Terminal 1" Then time21 = time21 + 1
If Left(Cells(i, 6), 2) = "22" And Cells(i, 18) = "Terminal 1" Then time22 = time22 + 1
If Left(Cells(i, 6), 2) = "23" And Cells(i, 18) = "Terminal 1" Then time23 = time23 + 1

Next i


Cells(t + 2, 3) = time0
Cells(t + 3, 3) = time1
Cells(t + 4, 3) = time2
Cells(t + 5, 3) = time3
Cells(t + 6, 3) = time4
Cells(t + 7, 3) = time5
Cells(t + 8, 3) = time6
Cells(t + 9, 3) = time7
Cells(t + 10, 3) = time8
Cells(t + 11, 3) = time9
Cells(t + 12, 3) = time10
Cells(t + 13, 3) = time11
Cells(t + 14, 3) = time12
Cells(t + 15, 3) = time13
Cells(t + 16, 3) = time14
Cells(t + 17, 3) = time15
Cells(t + 18, 3) = time16
Cells(t + 19, 3) = time17
Cells(t + 20, 3) = time18
Cells(t + 21, 3) = time19
Cells(t + 22, 3) = time20
Cells(t + 23, 3) = time21
Cells(t + 24, 3) = time22
Cells(t + 25, 3) = time23

'Terminal 3 setup

Cells(t + 28, 1) = "Teminal 1"
Cells(t + 28, 2) = "12 AM"
Cells(t + 29, 2) = "1 AM"
Cells(t + 30, 2) = "2 AM"
Cells(t + 31, 2) = "3 AM"
Cells(t + 32, 2) = "4 AM"
Cells(t + 33, 2) = "5 AM"
Cells(t + 34, 2) = "6 AM"
Cells(t + 35, 2) = "7 AM"
Cells(t + 36, 2) = "8 AM"
Cells(t + 37, 2) = "9 AM"
Cells(t + 38, 2) = "10 AM"
Cells(t + 39, 2) = "11 AM"
Cells(t + 40, 2) = "12 PM"
Cells(t + 41, 2) = "1 PM"
Cells(t + 42, 2) = "2 PM"
Cells(t + 43, 2) = "3 PM"
Cells(t + 44, 2) = "4 PM"
Cells(t + 45, 2) = "5 PM"
Cells(t + 46, 2) = "6 PM"
Cells(t + 47, 2) = "7 PM"
Cells(t + 48, 2) = "8 PM"
Cells(t + 49, 2) = "9 PM"
Cells(t + 50, 2) = "10 PM"
Cells(t + 51, 2) = "11 PM"

For i = 1 To t

If Left(Cells(i, 6), 2) = "00" And Cells(i, 18) = "Terminal 3" Then time30 = time30 + 1
If Left(Cells(i, 6), 2) = "01" And Cells(i, 18) = "Terminal 3" Then time31 = time31 + 1
If Left(Cells(i, 6), 2) = "02" And Cells(i, 18) = "Terminal 3" Then time32 = time32 + 1
If Left(Cells(i, 6), 2) = "03" And Cells(i, 18) = "Terminal 3" Then time33 = time33 + 1
If Left(Cells(i, 6), 2) = "04" And Cells(i, 18) = "Terminal 3" Then time34 = time34 + 1
If Left(Cells(i, 6), 2) = "05" And Cells(i, 18) = "Terminal 3" Then time35 = time35 + 1
If Left(Cells(i, 6), 2) = "06" And Cells(i, 18) = "Terminal 3" Then time36 = time36 + 1
If Left(Cells(i, 6), 2) = "07" And Cells(i, 18) = "Terminal 3" Then time37 = time37 + 1
If Left(Cells(i, 6), 2) = "08" And Cells(i, 18) = "Terminal 3" Then time38 = time38 + 1
If Left(Cells(i, 6), 2) = "09" And Cells(i, 18) = "Terminal 3" Then time39 = time39 + 1
If Left(Cells(i, 6), 2) = "10" And Cells(i, 18) = "Terminal 3" Then time310 = time310 + 1
If Left(Cells(i, 6), 2) = "11" And Cells(i, 18) = "Terminal 3" Then time311 = time311 + 1
If Left(Cells(i, 6), 2) = "12" And Cells(i, 18) = "Terminal 3" Then time312 = time312 + 1
If Left(Cells(i, 6), 2) = "13" And Cells(i, 18) = "Terminal 3" Then time313 = time313 + 1
If Left(Cells(i, 6), 2) = "14" And Cells(i, 18) = "Terminal 3" Then time314 = time314 + 1
If Left(Cells(i, 6), 2) = "15" And Cells(i, 18) = "Terminal 3" Then time315 = time315 + 1
If Left(Cells(i, 6), 2) = "16" And Cells(i, 18) = "Terminal 3" Then time316 = time316 + 1
If Left(Cells(i, 6), 2) = "17" And Cells(i, 18) = "Terminal 3" Then time317 = time317 + 1
If Left(Cells(i, 6), 2) = "18" And Cells(i, 18) = "Terminal 3" Then time318 = time318 + 1
If Left(Cells(i, 6), 2) = "19" And Cells(i, 18) = "Terminal 3" Then time319 = time319 + 1
If Left(Cells(i, 6), 2) = "20" And Cells(i, 18) = "Terminal 3" Then time320 = time320 + 1
If Left(Cells(i, 6), 2) = "21" And Cells(i, 18) = "Terminal 3" Then time321 = time321 + 1
If Left(Cells(i, 6), 2) = "22" And Cells(i, 18) = "Terminal 3" Then time322 = time322 + 1
If Left(Cells(i, 6), 2) = "23" And Cells(i, 18) = "Terminal 3" Then time323 = time323 + 1

Next i




Cells(t + 28, 3) = time30
Cells(t + 29, 3) = time31
Cells(t + 30, 3) = time32
Cells(t + 31, 3) = time33
Cells(t + 32, 3) = time34
Cells(t + 33, 3) = time35
Cells(t + 34, 3) = time36
Cells(t + 35, 3) = time37
Cells(t + 36, 3) = time38
Cells(t + 37, 3) = time39
Cells(t + 38, 3) = time310
Cells(t + 39, 3) = time311
Cells(t + 40, 3) = time312
Cells(t + 41, 3) = time313
Cells(t + 42, 3) = time314
Cells(t + 43, 3) = time315
Cells(t + 44, 3) = time316
Cells(t + 45, 3) = time317
Cells(t + 46, 3) = time318
Cells(t + 47, 3) = time319
Cells(t + 48, 3) = time320
Cells(t + 49, 3) = time321
Cells(t + 50, 3) = time322
Cells(t + 51, 3) = time323


    'Range(Cells(t + 3, 3), Cells(t + 25, 3)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete

End Sub

Sub PHXSheet()

r = 1

Sheets.Add.name = "PHX"

Sheets("Manifest").Activate
endRow = Cells(1, 1).End(xlDown).Row
For i = 2 To endRow

If Cells(i, 8) = "PHX" Then

Range(Cells(i, 1), Cells(i, 17)).Copy
Sheets("PHX").Cells(r, 1).Insert
r = r + 1
End If

Next i
Sheets("PHX").Activate
For i = 1 To endRow

If Cells(i, 9) = "AA" Then Cells(i, 18) = "Terminal 1"
If Cells(i, 9) = "DL" Then Cells(i, 18) = "Terminal 3"
If Cells(i, 9) = "UA" Then Cells(i, 18) = "Terminal 3"

Next i

    ActiveSheet.UsedRange.Sort Key1:=Range("R1"), order1:=xlAscending, key2:=Range( _
        "C1"), order2:=xlAscending, key3:=Range("D1"), Order3:=xlAscending, _
        Header:=False, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal
        
'Terminal One Set Up
        
t = Cells(1, 1).End(xlDown).Row

Cells(t + 2, 1) = "Teminal 1"
Cells(t + 2, 2) = "12 AM"
Cells(t + 3, 2) = "1 AM"
Cells(t + 4, 2) = "2 AM"
Cells(t + 5, 2) = "3 AM"
Cells(t + 6, 2) = "4 AM"
Cells(t + 7, 2) = "5 AM"
Cells(t + 8, 2) = "6 AM"
Cells(t + 9, 2) = "7 AM"
Cells(t + 10, 2) = "8 AM"
Cells(t + 11, 2) = "9 AM"
Cells(t + 12, 2) = "10 AM"
Cells(t + 13, 2) = "11 AM"
Cells(t + 14, 2) = "12 PM"
Cells(t + 15, 2) = "1 PM"
Cells(t + 16, 2) = "2 PM"
Cells(t + 17, 2) = "3 PM"
Cells(t + 18, 2) = "4 PM"
Cells(t + 19, 2) = "5 PM"
Cells(t + 20, 2) = "6 PM"
Cells(t + 21, 2) = "7 PM"
Cells(t + 22, 2) = "8 PM"
Cells(t + 23, 2) = "9 PM"
Cells(t + 24, 2) = "10 PM"
Cells(t + 25, 2) = "11 PM"



For i = 1 To t

If Left(Cells(i, 6), 2) = "00" And Cells(i, 18) = "Terminal 1" Then time0 = time0 + 1
If Left(Cells(i, 6), 2) = "01" And Cells(i, 18) = "Terminal 1" Then time1 = time1 + 1
If Left(Cells(i, 6), 2) = "02" And Cells(i, 18) = "Terminal 1" Then time2 = time2 + 1
If Left(Cells(i, 6), 2) = "03" And Cells(i, 18) = "Terminal 1" Then time3 = time3 + 1
If Left(Cells(i, 6), 2) = "04" And Cells(i, 18) = "Terminal 1" Then time4 = time4 + 1
If Left(Cells(i, 6), 2) = "05" And Cells(i, 18) = "Terminal 1" Then time5 = time5 + 1
If Left(Cells(i, 6), 2) = "06" And Cells(i, 18) = "Terminal 1" Then time6 = time6 + 1
If Left(Cells(i, 6), 2) = "07" And Cells(i, 18) = "Terminal 1" Then time7 = time7 + 1
If Left(Cells(i, 6), 2) = "08" And Cells(i, 18) = "Terminal 1" Then time8 = time8 + 1
If Left(Cells(i, 6), 2) = "09" And Cells(i, 18) = "Terminal 1" Then time9 = time9 + 1
If Left(Cells(i, 6), 2) = "10" And Cells(i, 18) = "Terminal 1" Then time10 = time10 + 1
If Left(Cells(i, 6), 2) = "11" And Cells(i, 18) = "Terminal 1" Then time11 = time11 + 1
If Left(Cells(i, 6), 2) = "12" And Cells(i, 18) = "Terminal 1" Then time12 = time12 + 1
If Left(Cells(i, 6), 2) = "13" And Cells(i, 18) = "Terminal 1" Then time13 = time13 + 1
If Left(Cells(i, 6), 2) = "14" And Cells(i, 18) = "Terminal 1" Then time14 = time14 + 1
If Left(Cells(i, 6), 2) = "15" And Cells(i, 18) = "Terminal 1" Then time15 = time15 + 1
If Left(Cells(i, 6), 2) = "16" And Cells(i, 18) = "Terminal 1" Then time16 = time16 + 1
If Left(Cells(i, 6), 2) = "17" And Cells(i, 18) = "Terminal 1" Then time17 = time17 + 1
If Left(Cells(i, 6), 2) = "18" And Cells(i, 18) = "Terminal 1" Then time18 = time18 + 1
If Left(Cells(i, 6), 2) = "19" And Cells(i, 18) = "Terminal 1" Then time19 = time19 + 1
If Left(Cells(i, 6), 2) = "20" And Cells(i, 18) = "Terminal 1" Then time20 = time20 + 1
If Left(Cells(i, 6), 2) = "21" And Cells(i, 18) = "Terminal 1" Then time21 = time21 + 1
If Left(Cells(i, 6), 2) = "22" And Cells(i, 18) = "Terminal 1" Then time22 = time22 + 1
If Left(Cells(i, 6), 2) = "23" And Cells(i, 18) = "Terminal 1" Then time23 = time23 + 1

Next i


Cells(t + 2, 3) = time0
Cells(t + 3, 3) = time1
Cells(t + 4, 3) = time2
Cells(t + 5, 3) = time3
Cells(t + 6, 3) = time4
Cells(t + 7, 3) = time5
Cells(t + 8, 3) = time6
Cells(t + 9, 3) = time7
Cells(t + 10, 3) = time8
Cells(t + 11, 3) = time9
Cells(t + 12, 3) = time10
Cells(t + 13, 3) = time11
Cells(t + 14, 3) = time12
Cells(t + 15, 3) = time13
Cells(t + 16, 3) = time14
Cells(t + 17, 3) = time15
Cells(t + 18, 3) = time16
Cells(t + 19, 3) = time17
Cells(t + 20, 3) = time18
Cells(t + 21, 3) = time19
Cells(t + 22, 3) = time20
Cells(t + 23, 3) = time21
Cells(t + 24, 3) = time22
Cells(t + 25, 3) = time23

'Terminal 2 setup

Cells(t + 28, 1) = "Teminal 1"
Cells(t + 28, 2) = "12 AM"
Cells(t + 29, 2) = "1 AM"
Cells(t + 30, 2) = "2 AM"
Cells(t + 31, 2) = "3 AM"
Cells(t + 32, 2) = "4 AM"
Cells(t + 33, 2) = "5 AM"
Cells(t + 34, 2) = "6 AM"
Cells(t + 35, 2) = "7 AM"
Cells(t + 36, 2) = "8 AM"
Cells(t + 37, 2) = "9 AM"
Cells(t + 38, 2) = "10 AM"
Cells(t + 39, 2) = "11 AM"
Cells(t + 40, 2) = "12 PM"
Cells(t + 41, 2) = "1 PM"
Cells(t + 42, 2) = "2 PM"
Cells(t + 43, 2) = "3 PM"
Cells(t + 44, 2) = "4 PM"
Cells(t + 45, 2) = "5 PM"
Cells(t + 46, 2) = "6 PM"
Cells(t + 47, 2) = "7 PM"
Cells(t + 48, 2) = "8 PM"
Cells(t + 49, 2) = "9 PM"
Cells(t + 50, 2) = "10 PM"
Cells(t + 51, 2) = "11 PM"

For i = 1 To t

If Left(Cells(i, 6), 2) = "00" And Cells(i, 18) = "Terminal 2" Then time20 = time20 + 1
If Left(Cells(i, 6), 2) = "01" And Cells(i, 18) = "Terminal 2" Then time21 = time21 + 1
If Left(Cells(i, 6), 2) = "02" And Cells(i, 18) = "Terminal 2" Then time22 = time22 + 1
If Left(Cells(i, 6), 2) = "03" And Cells(i, 18) = "Terminal 2" Then time23 = time23 + 1
If Left(Cells(i, 6), 2) = "04" And Cells(i, 18) = "Terminal 2" Then time24 = time24 + 1
If Left(Cells(i, 6), 2) = "05" And Cells(i, 18) = "Terminal 2" Then time25 = time25 + 1
If Left(Cells(i, 6), 2) = "06" And Cells(i, 18) = "Terminal 2" Then time26 = time26 + 1
If Left(Cells(i, 6), 2) = "07" And Cells(i, 18) = "Terminal 2" Then time27 = time27 + 1
If Left(Cells(i, 6), 2) = "08" And Cells(i, 18) = "Terminal 2" Then time28 = time28 + 1
If Left(Cells(i, 6), 2) = "09" And Cells(i, 18) = "Terminal 2" Then time29 = time29 + 1
If Left(Cells(i, 6), 2) = "10" And Cells(i, 18) = "Terminal 2" Then time210 = time210 + 1
If Left(Cells(i, 6), 2) = "11" And Cells(i, 18) = "Terminal 2" Then time211 = time211 + 1
If Left(Cells(i, 6), 2) = "12" And Cells(i, 18) = "Terminal 2" Then time212 = time212 + 1
If Left(Cells(i, 6), 2) = "13" And Cells(i, 18) = "Terminal 2" Then time213 = time213 + 1
If Left(Cells(i, 6), 2) = "14" And Cells(i, 18) = "Terminal 2" Then time214 = time214 + 1
If Left(Cells(i, 6), 2) = "15" And Cells(i, 18) = "Terminal 2" Then time215 = time215 + 1
If Left(Cells(i, 6), 2) = "16" And Cells(i, 18) = "Terminal 2" Then time216 = time216 + 1
If Left(Cells(i, 6), 2) = "17" And Cells(i, 18) = "Terminal 2" Then time217 = time217 + 1
If Left(Cells(i, 6), 2) = "18" And Cells(i, 18) = "Terminal 2" Then time218 = time218 + 1
If Left(Cells(i, 6), 2) = "19" And Cells(i, 18) = "Terminal 2" Then time219 = time219 + 1
If Left(Cells(i, 6), 2) = "20" And Cells(i, 18) = "Terminal 2" Then time220 = time220 + 1
If Left(Cells(i, 6), 2) = "21" And Cells(i, 18) = "Terminal 2" Then time221 = time221 + 1
If Left(Cells(i, 6), 2) = "22" And Cells(i, 18) = "Terminal 2" Then time222 = time222 + 1
If Left(Cells(i, 6), 2) = "23" And Cells(i, 18) = "Terminal 2" Then time223 = time223 + 1

Next i




Cells(t + 28, 3) = time20
Cells(t + 29, 3) = time21
Cells(t + 30, 3) = time22
Cells(t + 31, 3) = time23
Cells(t + 32, 3) = time24
Cells(t + 33, 3) = time25
Cells(t + 34, 3) = time26
Cells(t + 35, 3) = time27
Cells(t + 36, 3) = time28
Cells(t + 37, 3) = time29
Cells(t + 38, 3) = time210
Cells(t + 39, 3) = time211
Cells(t + 40, 3) = time212
Cells(t + 41, 3) = time213
Cells(t + 42, 3) = time214
Cells(t + 43, 3) = time215
Cells(t + 44, 3) = time216
Cells(t + 45, 3) = time217
Cells(t + 46, 3) = time218
Cells(t + 47, 3) = time219
Cells(t + 48, 3) = time320
Cells(t + 49, 3) = time321
Cells(t + 50, 3) = time322
Cells(t + 51, 3) = time323

'Terminal 3 setup

Cells(t + 54, 1) = "Teminal 1"
Cells(t + 55, 2) = "12 AM"
Cells(t + 56, 2) = "1 AM"
Cells(t + 57, 2) = "2 AM"
Cells(t + 58, 2) = "3 AM"
Cells(t + 59, 2) = "4 AM"
Cells(t + 60, 2) = "5 AM"
Cells(t + 61, 2) = "6 AM"
Cells(t + 62, 2) = "7 AM"
Cells(t + 63, 2) = "8 AM"
Cells(t + 64, 2) = "9 AM"
Cells(t + 65, 2) = "10 AM"
Cells(t + 66, 2) = "11 AM"
Cells(t + 67, 2) = "12 PM"
Cells(t + 68, 2) = "1 PM"
Cells(t + 69, 2) = "2 PM"
Cells(t + 70, 2) = "3 PM"
Cells(t + 71, 2) = "4 PM"
Cells(t + 72, 2) = "5 PM"
Cells(t + 73, 2) = "6 PM"
Cells(t + 74, 2) = "7 PM"
Cells(t + 75, 2) = "8 PM"
Cells(t + 76, 2) = "9 PM"
Cells(t + 77, 2) = "10 PM"
Cells(t + 78, 2) = "11 PM"

For i = 1 To t

If Left(Cells(i, 6), 2) = "00" And Cells(i, 18) = "Terminal 3" Then time30 = time30 + 1
If Left(Cells(i, 6), 2) = "01" And Cells(i, 18) = "Terminal 3" Then time31 = time31 + 1
If Left(Cells(i, 6), 2) = "02" And Cells(i, 18) = "Terminal 3" Then time32 = time32 + 1
If Left(Cells(i, 6), 2) = "03" And Cells(i, 18) = "Terminal 3" Then time33 = time33 + 1
If Left(Cells(i, 6), 2) = "04" And Cells(i, 18) = "Terminal 3" Then time34 = time34 + 1
If Left(Cells(i, 6), 2) = "05" And Cells(i, 18) = "Terminal 3" Then time35 = time35 + 1
If Left(Cells(i, 6), 2) = "06" And Cells(i, 18) = "Terminal 3" Then time36 = time36 + 1
If Left(Cells(i, 6), 2) = "07" And Cells(i, 18) = "Terminal 3" Then time37 = time37 + 1
If Left(Cells(i, 6), 2) = "08" And Cells(i, 18) = "Terminal 3" Then time38 = time38 + 1
If Left(Cells(i, 6), 2) = "09" And Cells(i, 18) = "Terminal 3" Then time39 = time39 + 1
If Left(Cells(i, 6), 2) = "10" And Cells(i, 18) = "Terminal 3" Then time310 = time310 + 1
If Left(Cells(i, 6), 2) = "11" And Cells(i, 18) = "Terminal 3" Then time311 = time311 + 1
If Left(Cells(i, 6), 2) = "12" And Cells(i, 18) = "Terminal 3" Then time312 = time312 + 1
If Left(Cells(i, 6), 2) = "13" And Cells(i, 18) = "Terminal 3" Then time313 = time313 + 1
If Left(Cells(i, 6), 2) = "14" And Cells(i, 18) = "Terminal 3" Then time314 = time314 + 1
If Left(Cells(i, 6), 2) = "15" And Cells(i, 18) = "Terminal 3" Then time315 = time315 + 1
If Left(Cells(i, 6), 2) = "16" And Cells(i, 18) = "Terminal 3" Then time316 = time316 + 1
If Left(Cells(i, 6), 2) = "17" And Cells(i, 18) = "Terminal 3" Then time317 = time317 + 1
If Left(Cells(i, 6), 2) = "18" And Cells(i, 18) = "Terminal 3" Then time318 = time318 + 1
If Left(Cells(i, 6), 2) = "19" And Cells(i, 18) = "Terminal 3" Then time319 = time319 + 1
If Left(Cells(i, 6), 2) = "20" And Cells(i, 18) = "Terminal 3" Then time320 = time320 + 1
If Left(Cells(i, 6), 2) = "21" And Cells(i, 18) = "Terminal 3" Then time321 = time321 + 1
If Left(Cells(i, 6), 2) = "22" And Cells(i, 18) = "Terminal 3" Then time322 = time322 + 1
If Left(Cells(i, 6), 2) = "23" And Cells(i, 18) = "Terminal 3" Then time323 = time323 + 1

Next i


Cells(t + 28, 3) = time30
Cells(t + 29, 3) = time31
Cells(t + 30, 3) = time32
Cells(t + 31, 3) = time33
Cells(t + 32, 3) = time34
Cells(t + 33, 3) = time35
Cells(t + 34, 3) = time36
Cells(t + 35, 3) = time37
Cells(t + 36, 3) = time38
Cells(t + 37, 3) = time39
Cells(t + 38, 3) = time310
Cells(t + 39, 3) = time311
Cells(t + 40, 3) = time312
Cells(t + 41, 3) = time313
Cells(t + 42, 3) = time314
Cells(t + 43, 3) = time315
Cells(t + 44, 3) = time316
Cells(t + 45, 3) = time317
Cells(t + 46, 3) = time318
Cells(t + 47, 3) = time319
Cells(t + 48, 3) = time320
Cells(t + 49, 3) = time321
Cells(t + 50, 3) = time322
Cells(t + 51, 3) = time323


End Sub
Sub JFKSheet()

r = 1

Sheets.Add.name = "JFK"

Sheets("Manifest").Activate
endRow = Cells(1, 1).End(xlDown).Row
For i = 2 To endRow

If Cells(i, 8) = "JFK" Then

Range(Cells(i, 1), Cells(i, 17)).Copy
Sheets("JFK").Cells(r, 1).Insert
r = r + 1
End If

Next i
Sheets("JFK").Activate
For i = 1 To endRow

If Cells(i, 9) = "AA" Then Cells(i, 18) = "Terminal 1"
If Cells(i, 9) = "DL" Then Cells(i, 18) = "Terminal 3"
If Cells(i, 9) = "UA" Then Cells(i, 18) = "Terminal 3"

Next i

    ActiveSheet.UsedRange.Sort Key1:=Range("R1"), order1:=xlAscending, key2:=Range( _
        "C1"), order2:=xlAscending, key3:=Range("D1"), Order3:=xlAscending, _
        Header:=False, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal
        
'Terminal One Set Up
        
t = Cells(1, 1).End(xlDown).Row

Cells(t + 2, 1) = "Teminal 1"
Cells(t + 2, 2) = "12 AM"
Cells(t + 3, 2) = "1 AM"
Cells(t + 4, 2) = "2 AM"
Cells(t + 5, 2) = "3 AM"
Cells(t + 6, 2) = "4 AM"
Cells(t + 7, 2) = "5 AM"
Cells(t + 8, 2) = "6 AM"
Cells(t + 9, 2) = "7 AM"
Cells(t + 10, 2) = "8 AM"
Cells(t + 11, 2) = "9 AM"
Cells(t + 12, 2) = "10 AM"
Cells(t + 13, 2) = "11 AM"
Cells(t + 14, 2) = "12 PM"
Cells(t + 15, 2) = "1 PM"
Cells(t + 16, 2) = "2 PM"
Cells(t + 17, 2) = "3 PM"
Cells(t + 18, 2) = "4 PM"
Cells(t + 19, 2) = "5 PM"
Cells(t + 20, 2) = "6 PM"
Cells(t + 21, 2) = "7 PM"
Cells(t + 22, 2) = "8 PM"
Cells(t + 23, 2) = "9 PM"
Cells(t + 24, 2) = "10 PM"
Cells(t + 25, 2) = "11 PM"



For i = 1 To t

If Left(Cells(i, 6), 2) = "00" And Cells(i, 18) = "Terminal 1" Then time0 = time0 + 1
If Left(Cells(i, 6), 2) = "01" And Cells(i, 18) = "Terminal 1" Then time1 = time1 + 1
If Left(Cells(i, 6), 2) = "02" And Cells(i, 18) = "Terminal 1" Then time2 = time2 + 1
If Left(Cells(i, 6), 2) = "03" And Cells(i, 18) = "Terminal 1" Then time3 = time3 + 1
If Left(Cells(i, 6), 2) = "04" And Cells(i, 18) = "Terminal 1" Then time4 = time4 + 1
If Left(Cells(i, 6), 2) = "05" And Cells(i, 18) = "Terminal 1" Then time5 = time5 + 1
If Left(Cells(i, 6), 2) = "06" And Cells(i, 18) = "Terminal 1" Then time6 = time6 + 1
If Left(Cells(i, 6), 2) = "07" And Cells(i, 18) = "Terminal 1" Then time7 = time7 + 1
If Left(Cells(i, 6), 2) = "08" And Cells(i, 18) = "Terminal 1" Then time8 = time8 + 1
If Left(Cells(i, 6), 2) = "09" And Cells(i, 18) = "Terminal 1" Then time9 = time9 + 1
If Left(Cells(i, 6), 2) = "10" And Cells(i, 18) = "Terminal 1" Then time10 = time10 + 1
If Left(Cells(i, 6), 2) = "11" And Cells(i, 18) = "Terminal 1" Then time11 = time11 + 1
If Left(Cells(i, 6), 2) = "12" And Cells(i, 18) = "Terminal 1" Then time12 = time12 + 1
If Left(Cells(i, 6), 2) = "13" And Cells(i, 18) = "Terminal 1" Then time13 = time13 + 1
If Left(Cells(i, 6), 2) = "14" And Cells(i, 18) = "Terminal 1" Then time14 = time14 + 1
If Left(Cells(i, 6), 2) = "15" And Cells(i, 18) = "Terminal 1" Then time15 = time15 + 1
If Left(Cells(i, 6), 2) = "16" And Cells(i, 18) = "Terminal 1" Then time16 = time16 + 1
If Left(Cells(i, 6), 2) = "17" And Cells(i, 18) = "Terminal 1" Then time17 = time17 + 1
If Left(Cells(i, 6), 2) = "18" And Cells(i, 18) = "Terminal 1" Then time18 = time18 + 1
If Left(Cells(i, 6), 2) = "19" And Cells(i, 18) = "Terminal 1" Then time19 = time19 + 1
If Left(Cells(i, 6), 2) = "20" And Cells(i, 18) = "Terminal 1" Then time20 = time20 + 1
If Left(Cells(i, 6), 2) = "21" And Cells(i, 18) = "Terminal 1" Then time21 = time21 + 1
If Left(Cells(i, 6), 2) = "22" And Cells(i, 18) = "Terminal 1" Then time22 = time22 + 1
If Left(Cells(i, 6), 2) = "23" And Cells(i, 18) = "Terminal 1" Then time23 = time23 + 1

Next i


Cells(t + 2, 3) = time0
Cells(t + 3, 3) = time1
Cells(t + 4, 3) = time2
Cells(t + 5, 3) = time3
Cells(t + 6, 3) = time4
Cells(t + 7, 3) = time5
Cells(t + 8, 3) = time6
Cells(t + 9, 3) = time7
Cells(t + 10, 3) = time8
Cells(t + 11, 3) = time9
Cells(t + 12, 3) = time10
Cells(t + 13, 3) = time11
Cells(t + 14, 3) = time12
Cells(t + 15, 3) = time13
Cells(t + 16, 3) = time14
Cells(t + 17, 3) = time15
Cells(t + 18, 3) = time16
Cells(t + 19, 3) = time17
Cells(t + 20, 3) = time18
Cells(t + 21, 3) = time19
Cells(t + 22, 3) = time20
Cells(t + 23, 3) = time21
Cells(t + 24, 3) = time22
Cells(t + 25, 3) = time23

'Terminal 2 setup

Cells(t + 28, 1) = "Teminal 1"
Cells(t + 28, 2) = "12 AM"
Cells(t + 29, 2) = "1 AM"
Cells(t + 30, 2) = "2 AM"
Cells(t + 31, 2) = "3 AM"
Cells(t + 32, 2) = "4 AM"
Cells(t + 33, 2) = "5 AM"
Cells(t + 34, 2) = "6 AM"
Cells(t + 35, 2) = "7 AM"
Cells(t + 36, 2) = "8 AM"
Cells(t + 37, 2) = "9 AM"
Cells(t + 38, 2) = "10 AM"
Cells(t + 39, 2) = "11 AM"
Cells(t + 40, 2) = "12 PM"
Cells(t + 41, 2) = "1 PM"
Cells(t + 42, 2) = "2 PM"
Cells(t + 43, 2) = "3 PM"
Cells(t + 44, 2) = "4 PM"
Cells(t + 45, 2) = "5 PM"
Cells(t + 46, 2) = "6 PM"
Cells(t + 47, 2) = "7 PM"
Cells(t + 48, 2) = "8 PM"
Cells(t + 49, 2) = "9 PM"
Cells(t + 50, 2) = "10 PM"
Cells(t + 51, 2) = "11 PM"

For i = 1 To t

If Left(Cells(i, 6), 2) = "00" And Cells(i, 18) = "Terminal 2" Then time20 = time20 + 1
If Left(Cells(i, 6), 2) = "01" And Cells(i, 18) = "Terminal 2" Then time21 = time21 + 1
If Left(Cells(i, 6), 2) = "02" And Cells(i, 18) = "Terminal 2" Then time22 = time22 + 1
If Left(Cells(i, 6), 2) = "03" And Cells(i, 18) = "Terminal 2" Then time23 = time23 + 1
If Left(Cells(i, 6), 2) = "04" And Cells(i, 18) = "Terminal 2" Then time24 = time24 + 1
If Left(Cells(i, 6), 2) = "05" And Cells(i, 18) = "Terminal 2" Then time25 = time25 + 1
If Left(Cells(i, 6), 2) = "06" And Cells(i, 18) = "Terminal 2" Then time26 = time26 + 1
If Left(Cells(i, 6), 2) = "07" And Cells(i, 18) = "Terminal 2" Then time27 = time27 + 1
If Left(Cells(i, 6), 2) = "08" And Cells(i, 18) = "Terminal 2" Then time28 = time28 + 1
If Left(Cells(i, 6), 2) = "09" And Cells(i, 18) = "Terminal 2" Then time29 = time29 + 1
If Left(Cells(i, 6), 2) = "10" And Cells(i, 18) = "Terminal 2" Then time210 = time210 + 1
If Left(Cells(i, 6), 2) = "11" And Cells(i, 18) = "Terminal 2" Then time211 = time211 + 1
If Left(Cells(i, 6), 2) = "12" And Cells(i, 18) = "Terminal 2" Then time212 = time212 + 1
If Left(Cells(i, 6), 2) = "13" And Cells(i, 18) = "Terminal 2" Then time213 = time213 + 1
If Left(Cells(i, 6), 2) = "14" And Cells(i, 18) = "Terminal 2" Then time214 = time214 + 1
If Left(Cells(i, 6), 2) = "15" And Cells(i, 18) = "Terminal 2" Then time215 = time215 + 1
If Left(Cells(i, 6), 2) = "16" And Cells(i, 18) = "Terminal 2" Then time216 = time216 + 1
If Left(Cells(i, 6), 2) = "17" And Cells(i, 18) = "Terminal 2" Then time217 = time217 + 1
If Left(Cells(i, 6), 2) = "18" And Cells(i, 18) = "Terminal 2" Then time218 = time218 + 1
If Left(Cells(i, 6), 2) = "19" And Cells(i, 18) = "Terminal 2" Then time219 = time219 + 1
If Left(Cells(i, 6), 2) = "20" And Cells(i, 18) = "Terminal 2" Then time220 = time220 + 1
If Left(Cells(i, 6), 2) = "21" And Cells(i, 18) = "Terminal 2" Then time221 = time221 + 1
If Left(Cells(i, 6), 2) = "22" And Cells(i, 18) = "Terminal 2" Then time222 = time222 + 1
If Left(Cells(i, 6), 2) = "23" And Cells(i, 18) = "Terminal 2" Then time223 = time223 + 1

Next i




Cells(t + 28, 3) = time20
Cells(t + 29, 3) = time21
Cells(t + 30, 3) = time22
Cells(t + 31, 3) = time23
Cells(t + 32, 3) = time24
Cells(t + 33, 3) = time25
Cells(t + 34, 3) = time26
Cells(t + 35, 3) = time27
Cells(t + 36, 3) = time28
Cells(t + 37, 3) = time29
Cells(t + 38, 3) = time210
Cells(t + 39, 3) = time211
Cells(t + 40, 3) = time212
Cells(t + 41, 3) = time213
Cells(t + 42, 3) = time214
Cells(t + 43, 3) = time215
Cells(t + 44, 3) = time216
Cells(t + 45, 3) = time217
Cells(t + 46, 3) = time218
Cells(t + 47, 3) = time219
Cells(t + 48, 3) = time320
Cells(t + 49, 3) = time321
Cells(t + 50, 3) = time322
Cells(t + 51, 3) = time323

'Terminal 3 setup

Cells(t + 54, 1) = "Teminal 1"
Cells(t + 55, 2) = "12 AM"
Cells(t + 56, 2) = "1 AM"
Cells(t + 57, 2) = "2 AM"
Cells(t + 58, 2) = "3 AM"
Cells(t + 59, 2) = "4 AM"
Cells(t + 60, 2) = "5 AM"
Cells(t + 61, 2) = "6 AM"
Cells(t + 62, 2) = "7 AM"
Cells(t + 63, 2) = "8 AM"
Cells(t + 64, 2) = "9 AM"
Cells(t + 65, 2) = "10 AM"
Cells(t + 66, 2) = "11 AM"
Cells(t + 67, 2) = "12 PM"
Cells(t + 68, 2) = "1 PM"
Cells(t + 69, 2) = "2 PM"
Cells(t + 70, 2) = "3 PM"
Cells(t + 71, 2) = "4 PM"
Cells(t + 72, 2) = "5 PM"
Cells(t + 73, 2) = "6 PM"
Cells(t + 74, 2) = "7 PM"
Cells(t + 75, 2) = "8 PM"
Cells(t + 76, 2) = "9 PM"
Cells(t + 77, 2) = "10 PM"
Cells(t + 78, 2) = "11 PM"

For i = 1 To t

If Left(Cells(i, 6), 2) = "00" And Cells(i, 18) = "Terminal 3" Then time30 = time30 + 1
If Left(Cells(i, 6), 2) = "01" And Cells(i, 18) = "Terminal 3" Then time31 = time31 + 1
If Left(Cells(i, 6), 2) = "02" And Cells(i, 18) = "Terminal 3" Then time32 = time32 + 1
If Left(Cells(i, 6), 2) = "03" And Cells(i, 18) = "Terminal 3" Then time33 = time33 + 1
If Left(Cells(i, 6), 2) = "04" And Cells(i, 18) = "Terminal 3" Then time34 = time34 + 1
If Left(Cells(i, 6), 2) = "05" And Cells(i, 18) = "Terminal 3" Then time35 = time35 + 1
If Left(Cells(i, 6), 2) = "06" And Cells(i, 18) = "Terminal 3" Then time36 = time36 + 1
If Left(Cells(i, 6), 2) = "07" And Cells(i, 18) = "Terminal 3" Then time37 = time37 + 1
If Left(Cells(i, 6), 2) = "08" And Cells(i, 18) = "Terminal 3" Then time38 = time38 + 1
If Left(Cells(i, 6), 2) = "09" And Cells(i, 18) = "Terminal 3" Then time39 = time39 + 1
If Left(Cells(i, 6), 2) = "10" And Cells(i, 18) = "Terminal 3" Then time310 = time310 + 1
If Left(Cells(i, 6), 2) = "11" And Cells(i, 18) = "Terminal 3" Then time311 = time311 + 1
If Left(Cells(i, 6), 2) = "12" And Cells(i, 18) = "Terminal 3" Then time312 = time312 + 1
If Left(Cells(i, 6), 2) = "13" And Cells(i, 18) = "Terminal 3" Then time313 = time313 + 1
If Left(Cells(i, 6), 2) = "14" And Cells(i, 18) = "Terminal 3" Then time314 = time314 + 1
If Left(Cells(i, 6), 2) = "15" And Cells(i, 18) = "Terminal 3" Then time315 = time315 + 1
If Left(Cells(i, 6), 2) = "16" And Cells(i, 18) = "Terminal 3" Then time316 = time316 + 1
If Left(Cells(i, 6), 2) = "17" And Cells(i, 18) = "Terminal 3" Then time317 = time317 + 1
If Left(Cells(i, 6), 2) = "18" And Cells(i, 18) = "Terminal 3" Then time318 = time318 + 1
If Left(Cells(i, 6), 2) = "19" And Cells(i, 18) = "Terminal 3" Then time319 = time319 + 1
If Left(Cells(i, 6), 2) = "20" And Cells(i, 18) = "Terminal 3" Then time320 = time320 + 1
If Left(Cells(i, 6), 2) = "21" And Cells(i, 18) = "Terminal 3" Then time321 = time321 + 1
If Left(Cells(i, 6), 2) = "22" And Cells(i, 18) = "Terminal 3" Then time322 = time322 + 1
If Left(Cells(i, 6), 2) = "23" And Cells(i, 18) = "Terminal 3" Then time323 = time323 + 1

Next i


Cells(t + 28, 3) = time30
Cells(t + 29, 3) = time31
Cells(t + 30, 3) = time32
Cells(t + 31, 3) = time33
Cells(t + 32, 3) = time34
Cells(t + 33, 3) = time35
Cells(t + 34, 3) = time36
Cells(t + 35, 3) = time37
Cells(t + 36, 3) = time38
Cells(t + 37, 3) = time39
Cells(t + 38, 3) = time310
Cells(t + 39, 3) = time311
Cells(t + 40, 3) = time312
Cells(t + 41, 3) = time313
Cells(t + 42, 3) = time314
Cells(t + 43, 3) = time315
Cells(t + 44, 3) = time316
Cells(t + 45, 3) = time317
Cells(t + 46, 3) = time318
Cells(t + 47, 3) = time319
Cells(t + 48, 3) = time320
Cells(t + 49, 3) = time321
Cells(t + 50, 3) = time322
Cells(t + 51, 3) = time323


End Sub

Sub TimeFormats2()
'Created by Michael Welch

'Adds a colon between each number so Excel can read it as a time format.
On Error Resume Next
Dim rCell As Range
    Dim rRng As Range
endRow = Cells(1, 1).End(xlDown).Row
    Set rRng = Range("G2", Range(Cells(2, 7), Cells(endRow, 7)))
  
    For Each rCell In rRng.Cells
       
       rCell = (Left(rCell.Offset(, -1), Len(rCell.Offset(, -1)) - 2) & ":" & Right(rCell.Offset(, -1), 2) & ":00")
       
       If rCell.Left(1) = ":" Then rCell = "00" & rCell
       
       
    Next rCell

'Changes the the time format to HHMM.
   

    Selection.NumberFormat = "h:mm am/pm"
        
'Removes the formula.
        
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Refreshes the cells and cancels the copy.
       
Selection.TextToColumns destination:=Selection, DataType:=xlDelimited, _
        FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        
    Columns("G:G").NumberFormat = "hh:mm am/pm"
End Sub


