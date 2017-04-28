Attribute VB_Name = "SorTests"
Sub Macro17()
Attribute Macro17.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro17 Macro
'

'
    ActiveWorkbook.Worksheets("Main").ListObjects("Inbound").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Main").ListObjects("Inbound").Sort.SortFields.Add _
        Key:=Range("Inbound[[#All],[Call Total]]"), SortOn:=xlSortOnValues, Order _
        :=xlDescending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Main").ListObjects("Inbound").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("G3").Select
End Sub
Sub sorttest()
'
' Sorttest Macro
'

'
    Range("E4:F4").Select
End Sub
Sub sorttester()
Attribute sorttester.VB_ProcData.VB_Invoke_Func = " \n14"
'
' sorttester Macro
'

'
    ActiveWorkbook.Worksheets("Main").ListObjects("Inbound").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Main").ListObjects("Inbound").Sort.SortFields.Add _
        Key:=Range("Inbound[[#All],[Call Total]]"), SortOn:=xlSortOnValues, Order _
        :=xlDescending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Main").ListObjects("Inbound").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
