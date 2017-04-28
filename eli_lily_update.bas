Attribute VB_Name = "eli_lily_update"
Sub Eli_Lilly_update()

Dim i As Integer
Dim endrange As Integer
Application.ScreenUpdating = False
'orders by Group ID and date

ActiveSheet.UsedRange.Sort Key1:=Range("AI2"), order1:=xlAscending, key2:=Range( _
        "R2"), order2:=xlAscending, key3:=Range("B2"), Order3:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, _
        DataOption3:=xlSortNormal

'First Format

Range("AL1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Base Rate"
    Columns("AH:AH").Select
    Selection.Cut
    Columns("AK:AK").Select
    Selection.Insert Shift:=xlToRight
    Selection.Cut
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight
    Range("AK1") = "Total Charge"
    Columns("AH:AI").Select
    Selection.Cut
    Columns("AJ:AJ").Select
    Selection.Insert Shift:=xlToRight
    Selection.Cut
    Columns("AK:AK").Select
    Selection.Insert Shift:=xlToRight
    Columns("O:O").Select
    Selection.Cut
    Columns("AI:AI").Select
    Selection.Insert Shift:=xlToRight
    Columns("N:N").Select
    Selection.Cut
    Columns("AH:AH").Select
    Selection.Insert Shift:=xlToRight
    Columns("M:M").Select
    Selection.Cut
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight
    Columns("L:L").Select
    Selection.Cut
    Columns("AF:AF").Select
    Selection.Insert Shift:=xlToRight
    Columns("D:D").Select
    Columns("I:I").Select
    Selection.Cut
    Columns("AE:AE").Select
    Selection.Insert Shift:=xlToRight
    Columns("H:H").Select
    Selection.Cut
    Columns("AD:AD").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:G").Select
    Selection.Copy
    Columns("AC:AC").Select
    Selection.Insert Shift:=xlToRight
    Columns("F:F").Select
    Application.CutCopyMode = False
    Selection.Cut
    Columns("AC:AC").Select
    Selection.Insert Shift:=xlToRight
    Columns("E:E").Select
    Selection.Cut
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight
    Columns("D:D").Select
    Selection.Cut
    Columns("AA:AA").Select
    Selection.Insert Shift:=xlToRight
    Columns("X:X").Select
    Selection.Cut
    Columns("Z:Z").Select
    Selection.Insert Shift:=xlToRight
    Columns("S:W").Select
    Selection.Cut
    Columns("Y:Y").Select
    Selection.Insert Shift:=xlToRight
    Columns("E:E").Select
    Columns("H:H").Select
    Selection.Cut
    Columns("S:S").Select
    Selection.Insert Shift:=xlToRight
    Selection.Cut
    Columns("T:T").Select
    Selection.Insert Shift:=xlToRight
    Columns("P:Q").Select
    Selection.Cut
    Columns("T:T").Select
    Selection.Insert Shift:=xlToRight
    Columns("E:E").Select
    Selection.Cut
    Columns("P:P").Select
    Selection.Insert Shift:=xlToRight
    Columns("L:L").Select
    Selection.Cut
    Columns("O:O").Select
    Selection.Insert Shift:=xlToRight
    Columns("K:K").Select
    Selection.Cut
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight
    Columns("J:J").Select
    Selection.Cut
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight
    Columns("I:I").Select
    Selection.Cut
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight
    Columns("A:A").Select
    Selection.Cut
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Cut
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:G").Select
    Selection.Cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Columns("A:A").Select
    Selection.Cut
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight
    Selection.Cut
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Cut
    Columns("AF:AF").Select
    Selection.Insert Shift:=xlToRight
    Columns("C:C").Select
    Selection.Cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Columns("B:B").Select
    Selection.Cut
    Columns("AC:AC").Select
    Selection.Insert Shift:=xlToRight
    Columns("AB:AB").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Reservation Date"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Reservation Number"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Passenger Name"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "TC Name"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Metro"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Email Address"
    Range("H7").Select
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Vehicle Type"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Pax Count"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Parking"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "Tolls"
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "Taxes"
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "Airport Fees"
    Range("V1").Select
    ActiveCell.FormulaR1C1 = "Misc. Fees"
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "Stops"
    Range("X3").Select

'highlights the trips

HighLights

'Counts rows

Cells(1, 1).Select
endrange = Range(Selection, Selection.End(xlDown)).Rows.Count

'Seperates transients

For i = 2 To endrange

If Cells(i, 15) = "" And Cells(i - 1, 15) <> "" Then

Cells(i, 35).EntireRow.Insert
endrange = endrange + 1

End If

Next i
'Seperates Groups

For i = 2 To endrange

If Cells(i, 15) = Cells(i - 1, 15) And Cells(i, 15) <> Cells(i + 1, 15) Then

Cells(i + 1, 15).EntireRow.Insert
endrange = endrange + 1

End If

Next i

'Adds second space

For i = 2 To endrange + 20

If Cells(i, 1) <> "" And Cells(i - 1, 1) = "" Then
Cells(i, 1).EntireRow.Insert
i = i + 2
endrange = endrange + 1
End If

Next i

    Range("A1:A5000").Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Select
    Selection.Interior.ColorIndex = white

Columns.AutoFit
Rows(1).Font.Bold = True
Rows(1).Font.Underline = True
Rows(1).Font.color = RGB(116, 125, 154)
Rows(1).Font.Size = 12

Cells(1, 1).Select
Selection.EntireRow.Insert
Cells(1, 1) = "Monthy Usage Report for Eli Lilly & Company - " & Format(Date, "mmmm yyyy")
Cells(1, 1).EntireRow.Select
Selection.Insert
Cells(1, 1).EntireRow.Select
Selection.Insert

'Final Formating

    Rows("4:4").Select
    Selection.RowHeight = 30
    Columns("E:E").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("F:F").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("G:G").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("D:D").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("C:C").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("A:A").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("H:H").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("I:I").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("K:K").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("L:L").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("M:M").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("N:N").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("O:O").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("P:P").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("S:S").Select
    Rows("4:4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("F:F").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("I:I").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("I4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1:AM6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = RGB(255, 255, 255)
    End With
    Range("A3").Select
    'ActiveCell.FormulaR1C1 = "Relax, you booked with Savoya..."
    Range("A3").Select
    With Selection.Font
        .name = "Calibri"
        .Size = 20
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("A5").Select
    With Selection.Font
        .name = "Calibri"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.Font.Bold = True
    With Selection.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.499984740745262
    End With
    Range("C4").Select
        Range("A6").Select
    Columns("A:A").ColumnWidth = 17.57
    Range("T34").Select
    Columns("C:C").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("C6").Select
    Columns("C:C").ColumnWidth = 19.43
    Columns("C:C").ColumnWidth = 20.57
    Columns("D:D").Select
    ActiveWindow.SmallScroll Down:=-27
    Range("D1").Select
    Columns("D:D").ColumnWidth = 31.57
    Columns("D:D").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("D6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("E:E").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("E6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveWindow.SmallScroll Down:=-12
    Range("F6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("G6").Select
    Columns("G:G").ColumnWidth = 29.14
    Columns("G:G").ColumnWidth = 26.14
    Columns("G:G").ColumnWidth = 23.14
    Columns("G:G").ColumnWidth = 21.57
    Columns("H:H").ColumnWidth = 20.71
    Columns("H:H").ColumnWidth = 23.43
    Columns("K:K").ColumnWidth = 12.57
    Columns("K:K").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("P:P").Select
    Columns("Y:Y").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.ColumnWidth = 31.14
    Selection.ColumnWidth = 22
    Columns("X:X").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("Z:Z").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("AA:AA").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("AA6").Select
    Columns("AA:AA").ColumnWidth = 22.86
    Columns("AB:AB").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.ColumnWidth = 26.14

    Range("AC6").Select
    Columns("AC:AC").ColumnWidth = 21.29
    Columns("AD:AD").ColumnWidth = 18.43
    Columns("AE:AE").ColumnWidth = 38.14
    Columns("AE:AE").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.ColumnWidth = 27.86
    Columns("AF:AF").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.ColumnWidth = 35
    Columns("AG:AG").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.ColumnWidth = 27
    Range("AG6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("AF6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("AE6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("AH:AH").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("AH6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
With ActiveSheet.Pictures.Insert("P:\Operations\Group Department\Macros\AddIn\Logo\Savoya.png")
    With .ShapeRange
        .LockAspectRatio = msoTrue
        .Width = 40
        .Height = 60
    End With
    .Left = ActiveSheet.Cells(1, 1).Left
    .Top = ActiveSheet.Cells(1, 1).Top
    .placement = 1
    .PrintObject = True
    
Columns("AN:XFD").EntireColumn.Hidden = True
End With
    
Application.ScreenUpdating = True
End Sub
Sub HighLights()
Dim endcolumn As Integer
Dim h As Integer
Dim hcell As Range
Dim hcell2 As Range

h = 2

Do
Set hcell = Cells(h, 1)
Set h2cell = Cells(h, 39)

Range(hcell, h2cell).Select
Selection.Interior.color = RGB(213, 232, 255)

h = h + 2

Loop While Not IsEmpty(Cells(h, 1))

End Sub




