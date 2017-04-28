Attribute VB_Name = "PickupTime"
Sub pickuptime()

'Created by John Stanley. For any questions, email jstanley@savoya.com

'v.1 - Pickup time only
'v.2 - Combined other macros and this one adjusts date of pickup according to the time of arrival.
'V.3 - Changed Dwell time input to accept in minutes.


Dim dtime As Variant
Dim x As Integer
Dim numrows As Integer
Dim rCell As Range
Dim iHours As Integer
Dim iMins As Integer


'Dwell time for the arrival trip
dtime = Application.InputBox(Prompt:="What is the Dwell time in mins? (E.g. 120)", Type:=1)
Application.ScreenUpdating = False
TimeFormats
' Row Count
numrows = Range("F2", Range("F2").End(xlDown)).Rows.Count

' Select the cell in the spreadsheet under the pickup column.
Range("D2").Select
' Establish Loop.
    For x = 1 To numrows
        'ActiveCell.NumberFormat = vbCrLf & "hh:mm AM/PM"
        ActiveCell.FormulaR1C1 = "=(1+RC[2]-" & dtime & "/1440)-MOD((1+RC[2]-" & dtime & "/1440),15/24/60)"
       'Inserting an if statement to test date condition
        ActiveCell.NumberFormat = "hhmm"
        'this is to remove the formulas
        ActiveCell.Copy
        ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

   'To change the pickup date according to the arrival time.
       ActiveCell.Offset(0, -1).FormulaR1C1 = "=IF((1-RC[1])>0,RC[2]-1,RC[2])"
       ActiveCell.Offset(1, 0).Select

Next
 
Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 
'to match the format of pickup time
Range("F2").Select
For x = 1 To numrows
    ActiveCell.NumberFormat = "hhmm"
    ActiveCell.Offset(1, 0).Select
Next

Range("C2").Select
For x = 1 To numrows
    ActiveCell.NumberFormat = "m/dd/yyyy"
    ActiveCell.Offset(1, 0).Select
Next
    
    
End Sub

  
Sub TimeFormats()
'Created by Michael Welch

'Adds a colon between each number so Excel can read it as a time format.
On Error Resume Next
Dim rCell As Range
    Dim rRng As Range

    Set rRng = Range("F2", Range("F2").End(xlDown))

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

