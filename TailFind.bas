Attribute VB_Name = "TailFind"
Sub TailFind()
ActiveSheet.name = "Tail Report"
For i = 1 To 1000

If Cells(i, 1) = "" And Cells(i + 1, 1) <> "" Then Cells(i, 1) = "1"

Next i


endRow = Cells(1, 1).End(xlDown).Row
For i = 1 To endRow

If Cells(i, 1) = "1" Then Cells(i, 1) = ""

Next i

For i = 1 To endRow
If Cells(i, 1).Interior.color = RGB(250, 250, 0) Then GoTo Done
If Cells(i, 1) = "Confirmation Number" Then
    
    Range(Cells(i, 1), Cells(i, 2)).Interior.color = RGB(250, 250, 0)
    Cells(i, 1).Select
Application.Goto Worksheets("Tail Report").Cells(i, 1), True
    web = Cells(i, 2)
    Shell ("C:\Program Files (x86)\Google\Chrome\Application\Chrome.exe -url https://core.savoya.net/#/savoya/reservations/" & web)
        If MsgBox("Reservation: " & web & ".", vbOKCancel) = vbCancel Then End
        
        
End If

Done:

Next i


End Sub
