Attribute VB_Name = "TransferLocID"
Sub TransferID()

endRow = Cells(1, 1).End(xlDown).Row

For i = 2 To endRow


If Len(Cells(i, 2)) = 2 Then Cells(i, 5) = ", " & Cells(i, 2)
If Cells(i, 2) = "New Jersey" Then Cells(i, 5) = ", NJ"
If Cells(i, 2) = "New York" Then Cells(i, 5) = ", NY"

Cells(i, 7) = Cells(i, 1).Value & Cells(i, 5).Value

Cells(i, 8) = Cells(i, 3)

Next i



End Sub

Sub idshift()

Dim id As Integer

On Error Resume Next

endRow = Cells(1, 1).End(xlDown).Row

For i = 1 To endRow

fnd = Cells(i, 1).Value
Cells.Range("F:G").Find(fnd, , , xlWhole).Offset(0, 1).Select
shft = ActiveCell


Cells(i, 2) = shft

Next i

End Sub


Sub completetransfer()

If Cells(1, 1) = "" Then
MsgBox ("Please paste all transit info into Column A")
End
End If

' preps the cost columns

endRow = Cells(1, 1).End(xlDown).Row

For i = 2 To endRow

If Cells(i, 1) = "To/From:" Then Cells(i, 1).EntireRow.Delete

Cells(i, 10) = Cells(i, 1)

Next i

If Cells(1, 12) = "" Then
MsgBox ("Please paste Location CSV into column L")
End
End If

' preps the tranfer locations

endRow = Cells(1, 12).End(xlDown).Row

For i = 2 To endRow


If Len(Cells(i, 13)) = 2 Then Cells(i, 12) = Cells(i, 12) & ", " & Cells(i, 13)
If Cells(i, 13) = "New Jersey" Then Cells(i, 12) = Cells(i, 12) & ", NJ"
If Cells(i, 13) = "New York" Then Cells(i, 12) = Cells(i, 12) & ", NY"
If Cells(i, 13) = "Pennsylvania" Then Cells(i, 12) = Cells(i, 12) & ", PA"
If Cells(i, 13) = "Delaware" Then Cells(i, 12) = Cells(i, 12) & ", DE"
Cells(i, 13) = Cells(i, 14)

Next i


Columns("N:N").Delete

' Adds location ids to places

Dim id As Integer

On Error Resume Next

endRow = Cells(2, 10).End(xlDown).Row

For i = 2 To endRow

fnd = Cells(i, 10).Value
sft = Cells.Range("L:M").Find(fnd, , , xlWhole).Offset(0, 1)



Cells(i, 11) = sft
sft = ""
If Cells(i, 11) = "" Then Cells(i, 11).Interior.color = RGB(255, 55, 55)
Next i

Columns("L:M").Delete

For i = 2 To endRow

If Cells(i, 11) = "" Then
MsgBox ("There are some missing location ids.  Please locate and run TranferFinish after you have filled them in.")
End
End If

Next i

transferfinish


End Sub

Sub transferfinish()

Dim vehnum As Integer

endRow = Cells(1, 1).End(xlDown).Row



air = Cells(1, 1).End(xlToRight).Column + 2

Columns("J:K").Cut
Columns("A:B").Insert

For i = 2 To endRow

Cells(i, 1) = Left(Cells(i, 1), Len(Cells(i, 1)) - 4)

Next i

Cells(1, 12) = "vendor"
Cells(1, 13) = "vendor_ID"
Cells(1, 14) = "transfer_cost"
Cells(1, 15) = "IATA"
Cells(1, 16) = "stop_2"
Cells(1, 17) = "Location ID"
Cells(1, 18) = "vehicle"
Cells(1, 19) = "cost_cxl"

vendor = InputBox("Who is the Vendor?")
venid = InputBox("What is the Vendor ID?")

vehnum = InputBox("How many vehicle types are there?")

nxt = 2

For veh = 1 To vehnum

vehtype = InputBox("What it the vehicle type?")
canc = InputBox("What is the cancelation?")

For col = 4 To air


For i = 2 To endRow

Cells(nxt, 12) = vendor
Cells(nxt, 13) = venid
Cells(nxt, 14) = Cells(i, col)
Cells(nxt, 15) = Cells(1, col)
Cells(nxt, 16) = Cells(i, 1)
Cells(nxt, 17) = Cells(i, 2)
Cells(nxt, 18) = vehtype
Cells(nxt, 19) = canc

nxt = nxt + 1

Next i

Next col

Next veh

Columns("A:K").Delete

End Sub
