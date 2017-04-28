Attribute VB_Name = "TimeTest"
Sub TimeFormatTest()


'Adds a colon between each number so Excel can read it as a time format.
On Error Resume Next
Dim rCell As Range
    Dim rRng As Range
    
    
    Set rRng = Selection

    For Each rCell In rRng.Cells
       
       rCell = (Left(rCell, Len(rCell) - 2) & ":" & Right(rCell, 2) & ":00")
       
       If Left(rCell, 1) = ":" Then rCell = "00" & rCell
       
       If Len(rCell) = 1 Then rCell = "00:0" & rCell
       
    Next rCell

'Changes the the time format to H:MM AM/PM.
   

    ActiveCell.EntireColumn.NumberFormat = "h:mm AM/PM"
        
'Removes the formula.
        
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

On Error GoTo 0

End Sub

