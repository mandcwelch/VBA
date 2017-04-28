Attribute VB_Name = "VendorRateCheck"
Sub vendorratechecker()

Dim rate1 As String

endRow = Cells(2, 9).End(xlDown).Row

For i = 2 To endRow

rate1 = Cells(i, 9)

Cells.Find(rate1).Activate

If ActiveCell.Column = 9 Then

Cells.FindNext(ActiveCell).Activate

End If

If ActiveCell.Column <> 9 Then

ActiveCell.Offset(0, 5) = Cells(i, 10)

End If

Next i

End Sub

Sub ratecompare()

For i = 1 To 36

If Cells(i, 6) = 0 Then

Cells(i, 6).Interior.color = RGB(255, 20, 20)

GoTo norate

End If

Cells(i, 5).Interior.color = RGB(20, 20, 255)

If Cells(i, 5) + 1 < Cells(i, 6) Then Cells(i, 5).Interior.color = RGB(100, 200, 20)

If Cells(i, 5) > Cells(i, 6) + 1 Then Cells(i, 5).Interior.color = RGB(200, 100, 20)

If Cells(i, 5) + 10 < Cells(i, 6) Then Cells(i, 5).Interior.color = RGB(20, 255, 20)

If Cells(i, 5) > Cells(i, 6) + 10 Then Cells(i, 5).Interior.color = RGB(255, 20, 20)

norate:

Next i

End Sub

Sub NoRateGrab()

For i = 1 To 130

If Cells(i, 6).Interior.color = RGB(255, 20, 20) Then

Cells(i, 7) = Cells(i, 1)

Cells(i, 8) = Cells(i, 5)

End If

Next i

End Sub

Sub chicagorategrab()

Dim check As String

endRow = Cells(2, 1).End(xlDown).Row
endRow2 = Cells(2, 5).End(xlDown).Row

For i = 2 To endRow

check = Left(Cells(i, 1), 5)

For x = 2 To endRow2

check2 = Left(Cells(x, 5), 5)

If check2 = check Then

Range(Cells(x, 3), Cells(x, 4)) = Range(Cells(i, 1), Cells(i, 2))
Cells(x, 3) = Cells(i, 1)
Cells(x, 4) = Cells(i, 2)
GoTo nxtone

End If

Next x

nxtone:

Next i


End Sub
