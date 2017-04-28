Attribute VB_Name = "Transfersrate"
Sub transferrate()

If Range("A1") = "" Then
MsgBox ("Please paste the cities, with state if possible, in the leftmost column and the airport addresses in the top row, with the Service Area Name in the first cell.")
End
End If

On Error Resume Next

ScreenUpdating = False

endCol = Cells(1, 1).End(xlToRight).Column

'check each column

For i = 2 To endCol

origin = Cells(1, i)

endRow = Cells(1, 1).End(xlDown).Row


    For x = 2 To endRow

    destination = Cells(x, 1)

    xmlstrdur = "URL;http://maps.googleapis.com/maps/api/directions/xml?origin=" & origin & ",&destination=" & destination & "&alternatives=false"

    With ActiveSheet.QueryTables.Add(Connection:=xmlstrdur, destination:=Range("M1"))


    .name = "1"
    .FieldNames = True
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = True
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .WebSelectionType = xlSpecifiedTables
    .WebFormatting = xlWebFormattingNone
    .WebTables = "2"
    .WebPreFormattedTextToColumns = True
    .WebConsecutiveDelimitersAsOne = True
    .WebSingleBlockTextImport = False
    .WebDisableDateRecognition = False
    .WebDisableRedirections = False
    .Refresh BackgroundQuery:=False
    
    
    End With

    If Range("M2") = "/error_message" Then
    Cells(x, i) = "Error"
    Else:
    Cells(x, i) = Range("AD8")
    Cells(x, i) = Round(Cells(x, i) / 60)
    End If
    Columns("M:BK").Delete

Application.Wait (Now + TimeValue("00:00:01"))

Next x

Next i

'rates

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

ScreenUpdating = True

End Sub
