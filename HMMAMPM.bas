Attribute VB_Name = "HMMAMPM"
Sub TimeFormat()
'Created by Michael Welch
'Time Format Macros to change ####Hrs to HHMM

' V.1 Changes ####Hrs to HHMM
' V.2 Can now change ###Hrs and ### to HHMM
' V.3 Now stops the copy and refreshes the cells


'Adds a colon between each number so Excel can read it as a time format.

Dim rCell As Range
    Dim rRng As Range

    Set rRng = Selection

    For Each rCell In rRng.Cells
       
       rCell = (Left(rCell, Len(rCell) - 2) & ":" & Right(rCell, 2) & ":00")
       
    Next rCell

'Changes the the time format to HH:MM AM/PM.
   

    Selection.NumberFormat = "h:mm AM/PM"
        
'Removes the formula.
        
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Refreshes the cells and cancels the copy.
       
Selection.TextToColumns destination:=Selection, DataType:=xlDelimited, _
        FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        
        
    
End Sub

