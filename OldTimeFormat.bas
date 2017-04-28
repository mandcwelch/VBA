Attribute VB_Name = "OldTimeFormat"
Sub OldSystemTimeFormat()
Attribute OldSystemTimeFormat.VB_Description = "Formats Time Properly for Old System"
Attribute OldSystemTimeFormat.VB_ProcData.VB_Invoke_Func = "t\n14"
'
' Created by: Michael Welch 10.26.16
' Use to change selected cells into prefered format for finishde manifest

Selection.NumberFormat = "General"
'
Selection.Replace What:=":", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
 Selection.Replace What:="hrs", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'Adds a colon between each number so Excel can read it as a time format.

Dim rCell As Range
    Dim rRng As Range

    Set rRng = Selection

    For Each rCell In rRng.Cells
       
       rCell = (Left(rCell, Len(rCell) - 2) & ":" & Right(rCell, 2) & ":00")
       
    Next rCell

'Changes the the time format to HHMM.
   

    Selection.NumberFormat = "h:mm am/pm"
        
'Removes the formula.
        
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Refreshes the cells and cancels the copy.
       
Selection.TextToColumns destination:=Selection, DataType:=xlDelimited

        
        
End Sub
