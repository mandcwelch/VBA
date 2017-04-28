Attribute VB_Name = "PickUpTime2"
Sub PickupTime2()

'Created by Michael Welch - 4/20/16
'Replaces older macro to round time and make departures on new system.



Dim dtime As Integer
Dim i As Integer

TimeFormats

dtime = InputBox("Please enter the dwell time in minutes.")
endRow = Cells(2, 1).End(xlDown).Row



For i = 2 To endRow

Cells(i, 4) = "=(ROUNDDOWN(((1+F" & i & ")* 1440-" & dtime & ")/15,0) * 15)/1440"

Next i
    Columns("D:D").Copy
    Columns("D:D").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

   'To change the pickup date according to the arrival time.

For i = 2 To endRow

If Cells(i, 4) < 1 Then

Cells(i, 3) = Cells(i, 5) - 1

Else: Cells(i, 3) = Cells(i, 5)

End If

Next i

Columns("D:D").NumberFormat = "hhmm"
Columns("F:F").NumberFormat = "hhmm"


End Sub

Sub TimeFormats()
'Created by Michael Welch

'Adds a colon between each number so Excel can read it as a time format.
On Error Resume Next
Dim rCell As Range
    Dim rRng As Range

    Set rRng = Range("F2", Range("F2").End(xlDown))

    answer = MsgBox("Is the time in text format?" & vbNewLine & "(no colons in the function bar)", vbYesNo)

    If answer = vbNo Then GoTo rtime
    
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
        
rtime:
    
End Sub
