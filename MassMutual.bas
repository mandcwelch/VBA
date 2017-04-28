Attribute VB_Name = "MassMutual"
Sub MassMutual()

Dim i As Integer

Cells(1, 1).Select
endrange = Range(Selection, Selection.End(xlDown)).Rows.Count

For i = 2 To endrange

If Cells(i, 6) <> "Mass Mutual Executive Travel" Then Cells(i, 6) = ""

Next i

On Error Resume Next

    Range("F1:F5000").Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete

    Columns("V:V").Delete
    Columns("R:T").Delete
    Columns("I:I").Delete
    Columns("B:E").Delete
    Columns("B:B").Delete
    Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells.EntireColumn.AutoFit
    
    
    Rows("1:1").RowHeight = 60
    Rows("2:2").RowHeight = 30
    
With ActiveSheet.Pictures.Insert("P:\Operations\Group Department\Information\Training\Macros\savoya_logo2.jpg")
    With .ShapeRange
        .LockAspectRatio = msoTrue
        .Width = 40
        .Height = 60
    End With
    .Left = ActiveSheet.Cells(1, 1).Left
    .Top = ActiveSheet.Cells(1, 1).Top
    .placement = 1
    .PrintObject = True
    
End With

Range("A3:L3").Select
    With Selection
        .Font.ColorIndex = 2
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
    End With
    With Selection.Interior
        .ColorIndex = 23
        .Pattern = xlSolid
    End With
    
Range("A2:L2").Select
    With Selection
        .Font.ColorIndex = 2
        .Font.Bold = True
        .Font.Size = 16
        .VerticalAlignment = xlCenter
    End With
    With Selection.Interior
        .ColorIndex = 23
        .Pattern = xlSolid
    End With
    
    Cells(2, 1) = "Mass Mutual Exeutive Travel - " & Cells(5, 2).Value
    
    
    h = 4
Do
Set hcell = Cells(h, 1)
Set h2cell = Cells(h, 12)

Range(hcell, h2cell).Select
Selection.Interior.color = RGB(213, 232, 255)

h = h + 2

Loop While Not IsEmpty(Cells(h, 1))

End Sub
