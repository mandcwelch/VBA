Attribute VB_Name = "TailClick"
Sub TailClick()

Dim main As Range
Dim endRow As Integer

endRow = Cells(1, 1).End(xlDown).Row
For i = 1 To endRow

If Cells(i, 1) = "1" Then Cells(i, 1) = ""

Next i

For i = 1 To endRow

If Cells(i, 1) = "Confirmation Number" Then

    web = Cells(i, 2)
    Shell ("Chrome.exe -url https://core.savoya.net/#/savoya/reservations/" & web)
    MsgBox ("Reservation: " & web & ".")

End If

Next i


End Sub

