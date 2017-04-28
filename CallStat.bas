Attribute VB_Name = "CallStat"
Sub CallStats()
Dim endRow As Integer

Application.ScreenUpdating = False

Columns("A:A").Delete

ActiveSheet.name = "Main"
'Sheets.Add.name = "Call Stats"
'Sheets("Main").Activate

endRow = Cells(1, 1).End(xlDown).Row

For i = 2 To endRow

Cells(i, 3) = 1

Next i

' Delete blank rows

Range(Cells(1, 2), Cells(endRow + 1, 2)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
endRow = Cells(1, 1).End(xlDown).Row

'Seperates outbound calls
out = 2

For i = 2 To endRow

If Cells(i, 1) = "Dialout" Then
Range(Cells(i, 1), Cells(i, 3)).Cut
Cells(out, 4).Insert
out = out + 1

End If

Next i

'Gathers inbound calls
inb = 2

For i = 2 To endRow

If Cells(i, 1) = "Inbound" Then
Range(Cells(i, 1), Cells(i, 3)).Copy
Cells(inb, 9).Insert
inb = inb + 1

End If

Next i

Columns("A:C").Delete

Range("A1") = "Call Type"
Range("B1") = "Agent"
Range("C1") = "Call Total"

Range("F1") = "Call Type"
Range("G1") = "Agent"
Range("H1") = "Call Total"

Tablequery

Application.ScreenUpdating = True

End Sub

Sub Tablequery()

end1 = Cells(1, 1).End(xlDown).Row

end2 = Cells(1, 6).End(xlDown).Row

    
'Outbound table

    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(end1, 3)), , xlYes).name = _
        "Outbound"
    Range("Outbound[#All]").Select
    ActiveWorkbook.Queries.Add name:="Outbound", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""Outbound""]}[Content]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Call Type"", type text}, {""Agent"", type text}, {""Call Total"", Int64.Type}})," & Chr(13) & "" & Chr(10) & "    #""Grouped Rows"" = Table.Group(#""Changed Type"", {""Agent""}, {{""Call Count"", each List.Sum([Call Total]), type number}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Grouped Rows"""
    
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Outbound" _
        , destination:=Sheets("Main").Range("J2")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Outbound]")
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
        .PreserveColumnInfo = False
        .ListObject.DisplayName = "Table1_2"
        .Refresh BackgroundQuery:=False
    End With
    
    'Inbound table
    
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 6), Cells(end2, 8)), , xlYes).name = _
        "Inbound"
    Range("Inbound[#All]").Select
    ActiveWorkbook.Queries.Add name:="Inbound", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""Inbound""]}[Content]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Call Type"", type text}, {""Agent"", type text}, {""Call Total"", Int64.Type}})," & Chr(13) & "" & Chr(10) & "    #""Grouped Rows"" = Table.Group(#""Changed Type"", {""Agent""}, {{""Call Count"", each List.Sum([Call Total]), type number}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Grouped Rows"""
    
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Inbound" _
        , destination:=Sheets("Main").Range("M2")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Inbound]")
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
        .PreserveColumnInfo = False
        .ListObject.DisplayName = "Table3_4"
        .Refresh BackgroundQuery:=False
        End With
    

    
Columns("A:I").Delete

tend1 = Cells(1, 1).End(xlDown).Row

tend2 = Cells(1, 4).End(xlDown).Row
    
    
    
    ActiveWorkbook.Worksheets("Main").ListObjects("Table1_2").Sort.SortFields.Add _
        Key:=Range(Cells(2, 2), Cells(tend1, 2)), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortTextAsNumbers
        
    With ActiveWorkbook.Worksheets("Main").ListObjects("Table1_2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
     ActiveWorkbook.Worksheets("Main").ListObjects("Table3_4").Sort.SortFields.Add _
        Key:=Range(Cells(2, 5), Cells(tend2, 5)), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortTextAsNumbers
    
    With ActiveWorkbook.Worksheets("Main").ListObjects("Table3_4").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Range("A1") = "Outbound Call Totals"
Range("D1") = "Inbound Call Totals"
Range("A1:B1").Interior.color = RGB(200, 200, 255)
Range("D1:E1").Interior.color = RGB(200, 200, 255)


End Sub

