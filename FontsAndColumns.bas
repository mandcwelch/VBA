Attribute VB_Name = "FontsAndColumns"
Sub test()

Dim ws As Worksheet

Set ws = ActiveSheet

ws.Cells(1, 1).Value = "sample"

ws.Cells(1, 1).Font.Bold = True

ws.Cells(1, 1).Copy

ws.Cells(2, 1).PasteSpecial Paste:=xlValues


End Sub


Sub ProcedureA()

ActiveSheet.Cells(1, 1).Value = "sample"

Call ProcedureB
ActiveSheet.Cells(1, 1).Font.Bold = True

End Sub

Sub ProcedureB()

ActiveSheet.Cells(1, 1).Font.Size = 24

ActiveSheet.Cells(1, 1).Font.name = "Arial"

End Sub

Sub DeleteColumns()
Dim ColNeeded As Boolean
Dim rCell As Range
Dim endCol As Integer
Dim endRow As Integer

ColNeeded = False

endCol = InputBox("Enter Column to Check")

Range("A1").Select
endrange = Selection.End(xlDown)
endRow = endrange.Rows.Count
    
    Set rRng = Range(endCol, 3)(endCol, endRow)

    For Each rCell In rRng.Cells

        If rCell.Value <> 0 Then ColNeeded = True

Next rCell

End Sub
