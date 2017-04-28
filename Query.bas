Attribute VB_Name = "Query"
Sub Query()
Attribute Query.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro24 Macro
'

'
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$C$1058"), , xlYes).name = _
        "Outbound"
    Range("Table1[#All]").Select
    ActiveWorkbook.Queries.Add name:="Table1", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""Table1""]}[Content]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Call Type"", type text}, {""Agent"", type text}, {""Call Total"", Int64.Type}})," & Chr(13) & "" & Chr(10) & "    #""Grouped Rows"" = Table.Group(#""Changed Type"", {""Agent""}, {{""Call Count"", each List.Sum([Call Total]), type number}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Grouped Rows"""
    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Table1" _
        , destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table1]")
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
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Application.CommandBars("Workbook Queries").Visible = False
    ActiveWorkbook.Worksheets("Sheet3").ListObjects("Table1_2").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Sheet3").ListObjects("Table1_2").Sort.SortFields. _
        Add Key:=Range("Table1_2[Call Count]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet3").ListObjects("Table1_2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
