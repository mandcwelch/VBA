Attribute VB_Name = "IATA"
Sub IATAdup()
Dim i As Integer
Dim x As Integer
Dim endrange As Integer

Cells(2, 1).Select
endrange = Range(Selection, Selection.End(xlDown)).Rows.Count


For i = 2 To endrange
If Left(Cells(i, 8), 3) = Mid(Cells(i, 8), 5, 3) Then
x = Len(Cells(i, 8)) - 4
Cells(i, 8) = Right(Cells(i, 8), x)
End If
Next i


For i2 = 2 To endrange
If Left(Cells(i2, 12), 3) = Mid(Cells(i2, 12), 5, 3) Then
x = Len(Cells(i2, 12)) - 4
Cells(i2, 12) = Right(Cells(i2, 12), x)
End If
Next i2

End Sub
