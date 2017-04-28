Attribute VB_Name = "TimeFormat"
  
Sub Time_Format()
Attribute Time_Format.VB_ProcData.VB_Invoke_Func = " \n14"
'Created by Michael Welch
'Time Format Macros to change ####Hrs to HHMM

' V.1 Changes ####Hrs to HHMM
' V.2 Can now change ###Hrs and ### to HHMM
' V.3 Now stops the copy and refreshes the cells

'Removes "Hrs" from the selected cells.

 Selection.Replace What:="hrs", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'Adds a colon between each number so Excel can read it as a time format.
On Error Resume Next
Dim rCell As Range
    Dim rRng As Range

    Set rRng = Selection

    For Each rCell In rRng.Cells
       
       rCell = (Left(rCell, Len(rCell) - 2) & ":" & Right(rCell, 2) & ":00")
       
       If rCell.Left(1) = ":" Then rCell = "00" & rCell
       
       
       
    Next rCell

'Changes the the time format to HHMM.
   

    Selection.NumberFormat = "h:mm am/pm"
        
'Removes the formula.
        
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Refreshes the cells and cancels the copy.
       
Selection.TextToColumns destination:=Selection, DataType:=xlDelimited, _
        FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        
        
    
End Sub
