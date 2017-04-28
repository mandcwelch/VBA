Attribute VB_Name = "DriverAppUsechart"
Option Explicit

Sub Driver_OS_chart()
Dim i As Integer
Dim endRow As Integer

If Cells(1, 6) <> "" Then Columns("E:F").Delete

endRow = Cells(1, 1).End(xlDown).Row

For i = 2 To endRow

    If Cells(i, 5) = "No Vendor Onsite Text" Then

        Range(Cells(i, 1), Cells(i, 5)).Interior.color = RGB(255, 255, 153)
        GoTo noText
    End If
        
     If MsgBox("Is the following a proper onsite note?" & vbCrLf & vbCrLf & Cells(i, 5), vbYesNo) = vbNo Then
          
          Range(Cells(i, 1), Cells(i, 5)).Interior.color = RGB(255, 204, 255)
          
          Else: Range(Cells(i, 1), Cells(i, 5)).Interior.color = RGB(204, 255, 204)
          
    End If
    
noText:

Next i

Cells(endRow + 2, 1) = "Good Onsite Text"
Cells(endRow + 2, 1).Interior.color = RGB(204, 255, 204)

Cells(endRow + 3, 1) = "Onsite Text Needs Assistance"
Cells(endRow + 4, 1).Interior.color = RGB(255, 204, 255)

Cells(endRow + 4, 1) = "App Not Used"
Cells(endRow + 4, 1).Interior.color = RGB(255, 255, 153)

With Range(Cells(1, 1), Cells(1, 5))
        .Font.ColorIndex = 2
        .Font.Bold = True
        .Interior.ColorIndex = 23
        .Interior.Pattern = xlSolid
    End With
    
With Range(Cells(2, 1), Cells(endRow, 5))

        .Borders.Weight = 2
        
    End With
    
Columns.AutoFit
    
Columns("E:E").ColumnWidth = 25

For i = 1 To endRow
    Cells(i, 6) = " "
Next i

End Sub
