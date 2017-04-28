Attribute VB_Name = "testing"
Sub test()

Rows("2:2").Insert
Range("A2") = "City"

endRow = Cells(3, 1).End(xlDown).Row
endCol = Cells(3, 1).End(xlToRight).Column



For col = 3 To (endCol * 2) Step 2

Columns(col).Insert

Cells(2, col) = "Rate"
Cells(2, col - 1) = "Time(Minutes)"

Next col

endCol = Cells(2, 1).End(xlToRight).Column

Cells(1, endCol + 1) = "Hourly Rate:"
Cells(1, endCol + 2) = InputBox("What is the hourly rate?")
Cells(1, endCol + 2).NumberFormat = ("$00.00")

For col = 3 To endCol Step 2
Columns(col).NumberFormat = ("$00.00")

For i = 3 To endRow

If ((((Cells(i, col - 1) / 60) * Cells(1, endCol + 2)) * 2) + (Cells(1, endCol + 2) * 0.25)) < Cells(1, endCol + 2) Then
Cells(i, col) = Cells(1, endCol + 2)
Else: Cells(i, col) = Round((((Cells(i, col - 1) / 60) * Cells(1, endCol + 2)) * 2) + (Cells(1, endCol + 2) * 0.25))
End If

Next i

Next col


End Sub

Sub nexttest()

Cells(12, 2) = "=ROUNDUP((IF(((((B3/60)*$N$2)*2)+($N$2*0.25))<$N$2,$N$2,((((B3/60)*$N$2)*2)+($N$2*0.25))))*2,0)/2"

Cells(13, 2).Value = Cells(12, 2).Value

End Sub
