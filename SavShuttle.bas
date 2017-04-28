Attribute VB_Name = "SavShuttle"
Sub SavShuttle()
'Application written by John Stanley. For any errors encountered or questions unanswered, please email jstanley@savoya.com

'v.1 Pickup time only
'v.2 - Combined other macros and this one adjusts date of pickup according to the time of arrival.
'V.3 - Changed Dwell time input to accept in minutes.
'V.4 - Combined a Shuttle form with v.3 SavManifest
'V.5 - Added a Special Airports module,added 30 more shuttle inputs and a test for Excel time format, July 2015


Dim res As String
With ActiveSheet
res = .Evaluate("cell(""Format""," & .Range("F2").Address & ")")
End With


If res = "G" Then
    Call SS_2
    Call SS_1
Else
    Call SS_1
    
End If
End Sub

Private Sub SS_1()

Dim dtime As Variant
Dim x As Integer
Dim numrows As Integer
Dim rCell As Range
Dim iHours As Integer
Dim iMins As Integer
Dim LastRow As Long
Dim Filename As String
Dim myResponse As String
Dim LSearchRow As Integer
Dim LCopyToRow As Integer
Dim LSearchValue As String
Dim K As Variant

'Dwell time for the arrival trip
dtime = Application.InputBox(Prompt:="What is the Dwell time in mins? (E.g. 120)", Type:=1)
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
 
Columns("F:F").Select
    Selection.NumberFormat = "hhmm"
        
        
        
'Special Airports Module
   
  K = Application.InputBox(Prompt:="Enter the number of SPECIAL airports in this itinerary", Type:=1)
  Application.ScreenUpdating = False
'Row Count
  numrows = Range("A2", Range("A2").End(xlDown)).Rows.Count
   
  For x = 1 To K
    LSearchValue = InputBox("Please enter a special airport", "Enter value")
    dtime = Application.InputBox(Prompt:="What is the Dwell time in mins? (E.g. 120)", Type:=1)

'Start search in row 2
  LSearchRow = 2
  While Len(Range("A" & CStr(LSearchRow)).Value) > 0
      'If value in column J = LSearchValue
      If Range("J" & CStr(LSearchRow)).Value = LSearchValue Then
         'Select row in Sheet1 to copy
         Rows(CStr(LSearchRow) & ":" & CStr(LSearchRow)).Select
         Selection.Copy
         Range("D" & (ActiveCell.Row)).Formula = "=(1+RC[2]-" & dtime & "/1440)-MOD((1+RC[2]-" & dtime & "/1440),15/24/60)"
         Range("D" & (ActiveCell.Row)).NumberFormat = "hhmm"
       'Inserting an if statement to test date condition
        'this is to remove the formulas
        Range("D" & (ActiveCell.Row)).Copy
        Range("D" & (ActiveCell.Row)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'To change the pickup date according to the arrival time.
        'ActiveCell.Offset(0, -1).FormulaR1C1 = "=IF((1-RC[1])>0,RC[2]-1,RC[2])"
        ActiveCell.Offset(1, 0).Select
End If
'If value in column J = LSearchValue
      If Range("G" & CStr(LSearchRow)).Value = LSearchValue Then
         'Select row in Sheet1 to copy
          Rows(CStr(LSearchRow) & ":" & CStr(LSearchRow)).Select
          Selection.Copy
          Range("D" & (ActiveCell.Row)).Formula = "=(1+RC[2]-" & dtime & "/1440)-MOD((1+RC[2]-" & dtime & "/1440),15/24/60)"
          Range("D" & (ActiveCell.Row)).NumberFormat = "hhmm"
          'Inserting an if statement to test date condition
            'this is to remove the formulas
            Range("D" & (ActiveCell.Row)).Copy
            Range("D" & (ActiveCell.Row)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
'To change the pickup date according to the arrival time.
'ActiveCell.Offset(0, -1).FormulaR1C1 = "=IF((1-RC[1])>0,RC[2]-1,RC[2])"
          ActiveCell.Offset(1, 0).Select
  
End If
LSearchRow = LSearchRow + 1
Wend
Next
               
'Shuttles Module
myResponse = MsgBox("Do you have Shuttles in this manifest?", vbYesNo)
    If myResponse = vbNo Then Exit Sub

Application.ScreenUpdating = True
ActiveSheet.name = "Departures"

'adds sheet for Input
Worksheets.Add(After:=Worksheets(Worksheets.Count)).name = "Input"
  Range("A1").FormulaR1C1 = "Date"
  Range("B1").FormulaR1C1 = "Flt Arr. Time 1"
  Range("C1").FormulaR1C1 = "Flt Arr. Time 2"
  Range("D1").FormulaR1C1 = "P/U Time"
  Range("E1").FormulaR1C1 = "Veh.Type"
  Range("F1").FormulaR1C1 = "Airport"
  Range("G1").FormulaR1C1 = "Pick-up Location"

'Invoke the userform
  Shuttle.Show
  Sheets("Departures").Select
'Hides the first two rows temporarily
  Rows("1:1").Select
  Selection.EntireRow.Hidden = True

'Sorts the whole worksheet
  Columns("A:A").Select
    Range("A2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove


    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=RC[5]&RC[6]&RC[7]&RC[10]&RC[8]&RC[9]"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("A2:A" & lngLastRow).FillDown


    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A2:AX" & LastRow).Sort Key1:=Range("A2:A" & LastRow), _
    order1:=xlAscending, Header:=xlNo

'Deletes the column after sorting procedure
    Columns("A:A").Select
    Range("A2").Activate
    Selection.Delete Shift:=xlToLeft


'The process begins(True or false) depending on the time proximity
    Columns("G:G").Select
    Range("G2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-3]=R[1]C[-3],RC[-3]=R[-1]C[-3]),RC[-3]=R[1]C[-3],RC[-3])"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("G2:G" & lngLastRow).FillDown

'Shuttle input
  'First shuttle
    Columns("H:H").Select
    Range("H2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-3]=DATEVALUE(TEXT(Input!R2C1,""mm/dd/yy"")),RC[-2]>=Input!R2C2,RC[-2]<=Input!R2C3,RC[1]=Input!R2C6,RC[4]=Input!R2C7),Input!R2C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("H2:H" & lngLastRow).FillDown
  
  'Second shuttle
    Columns("I:I").Select
    Range("I2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("I2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-4]=DATEVALUE(TEXT(Input!R3C1,""mm/dd/yy"")),RC[-3]>=Input!R3C2,RC[-3]<=Input!R3C3,RC[1]=Input!R3C6,RC[4]=Input!R3C7),Input!R3C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("I2:I" & lngLastRow).FillDown

  'Third shuttle
    Columns("J:J").Select
    Range("J2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-5]=DATEVALUE(TEXT(Input!R4C1,""mm/dd/yy"")),RC[-4]>=Input!R4C2,RC[-4]<=Input!R4C3,RC[1]=Input!R4C6,RC[4]=Input!R4C7),Input!R4C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("J2:J" & lngLastRow).FillDown

  'Fourth shuttle
    Columns("K:K").Select
    Range("K2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("K2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-6]=DATEVALUE(TEXT(Input!R5C1,""mm/dd/yy"")),RC[-5]>=Input!R5C2,RC[-5]<=Input!R5C3,RC[1]=Input!R5C6,RC[4]=Input!R5C7),Input!R5C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("K2:K" & lngLastRow).FillDown

  'Fifth shuttle
    Columns("L:L").Select
    Range("L2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("L2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-7]=DATEVALUE(TEXT(Input!R6C1,""mm/dd/yy"")),RC[-6]>=Input!R6C2,RC[-6]<=Input!R6C3,RC[1]=Input!R6C6,RC[4]=Input!R6C7),Input!R6C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("L2:L" & lngLastRow).FillDown

    'Sixth shuttle
    Columns("M:M").Select
    Range("M2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("M2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-8]=DATEVALUE(TEXT(Input!R7C1,""mm/dd/yy"")),RC[-7]>=Input!R7C2,RC[-7]<=Input!R7C3,RC[1]=Input!R7C6,RC[4]=Input!R7C7),Input!R7C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("M2:M" & lngLastRow).FillDown

    'Seventh shuttle
    Columns("N:N").Select
    Range("N2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("N2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-9]=DATEVALUE(TEXT(Input!R8C1,""mm/dd/yy"")),RC[-8]>=Input!R8C2,RC[-8]<=Input!R8C3,RC[1]=Input!R8C6,RC[4]=Input!R8C7),Input!R8C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("N2:N" & lngLastRow).FillDown

   'Eighth shuttle
    Columns("O:O").Select
    Range("O2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("O2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-10]=DATEVALUE(TEXT(Input!R9C1,""mm/dd/yy"")),RC[-9]>=Input!R9C2,RC[-9]<=Input!R9C3,RC[1]=Input!R9C6,RC[4]=Input!R9C7),Input!R9C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("O2:O" & lngLastRow).FillDown

    'Ninth shuttle
    Columns("P:P").Select
    Range("P2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("P2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-11]=DATEVALUE(TEXT(Input!R10C1,""mm/dd/yy"")),RC[-10]>=Input!R10C2,RC[-10]<=Input!R10C3,RC[1]=Input!R10C6,RC[4]=Input!R10C7),Input!R10C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("P2:P" & lngLastRow).FillDown

   'Tenth shuttle
    Columns("Q:Q").Select
    Range("Q2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-12]=DATEVALUE(TEXT(Input!R11C1,""mm/dd/yy"")),RC[-11]>=Input!R11C2,RC[-11]<=Input!R11C3,RC[1]=Input!R11C6,RC[4]=Input!R11C7),Input!R11C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("Q2:Q" & lngLastRow).FillDown

    '11th shuttle
    Columns("R:R").Select
    Range("R2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("R2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-13]=DATEVALUE(TEXT(Input!R12C1,""mm/dd/yy"")),RC[-12]>=Input!R12C2,RC[-12]<=Input!R12C3,RC[1]=Input!R12C6,RC[4]=Input!R12C7),Input!R12C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("R2:R" & lngLastRow).FillDown

   '12th shuttle
    Columns("S:S").Select
    Range("S2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("S2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-14]=DATEVALUE(TEXT(Input!R13C1,""mm/dd/yy"")),RC[-13]>=Input!R13C2,RC[-13]<=Input!R13C3,RC[1]=Input!R13C6,RC[4]=Input!R13C7),Input!R13C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("S2:S" & lngLastRow).FillDown

    '13th shuttle
    Columns("T:T").Select
    Range("T2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("T2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-15]=DATEVALUE(TEXT(Input!R14C1,""mm/dd/yy"")),RC[-14]>=Input!R14C2,RC[-14]<=Input!R14C3,RC[1]=Input!R14C6,RC[4]=Input!R14C7),Input!R14C4,""""),"""")"
   lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("T2:T" & lngLastRow).FillDown

    '14th shuttle
    Columns("U:U").Select
    Range("U2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("U2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-16]=DATEVALUE(TEXT(Input!R15C1,""mm/dd/yy"")),RC[-15]>=Input!R15C2,RC[-15]<=Input!R15C3,RC[1]=Input!R15C6,RC[4]=Input!R15C7),Input!R15C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("U2:U" & lngLastRow).FillDown

    '15th shuttle
    Columns("V:V").Select
    Range("V2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("V2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-17]=DATEVALUE(TEXT(Input!R16C1,""mm/dd/yy"")),RC[-16]>=Input!R16C2,RC[-16]<=Input!R16C3,RC[1]=Input!R16C6,RC[4]=Input!R16C7),Input!R16C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("V2:V" & lngLastRow).FillDown

    '16th shuttle
    Columns("W:W").Select
    Range("W2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("W2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-18]=DATEVALUE(TEXT(Input!R17C1,""mm/dd/yy"")),RC[-17]>=Input!R17C2,RC[-17]<=Input!R17C3,RC[1]=Input!R17C6,RC[4]=Input!R17C7),Input!R17C4,""""),"""")"
   lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("W2:W" & lngLastRow).FillDown

    '17th shuttle
    Columns("X:X").Select
    Range("X2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("X2").Select
    ActiveCell.FormulaR1C1 = _
       "=IFERROR(IF(AND(RC[-19]=DATEVALUE(TEXT(Input!R18C1,""mm/dd/yy"")),RC[-18]>=Input!R18C2,RC[-18]<=Input!R18C3,RC[1]=Input!R18C6,RC[4]=Input!R18C7),Input!R18C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("X2:X" & lngLastRow).FillDown

   '18th shuttle
    Columns("Y:Y").Select
    Range("Y2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("Y2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-20]=DATEVALUE(TEXT(Input!R19C1,""mm/dd/yy"")),RC[-19]>=Input!R19C2,RC[-19]<=Input!R19C3,RC[1]=Input!R19C6,RC[4]=Input!R19C7),Input!R19C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("Y2:Y" & lngLastRow).FillDown

    '19th shuttle
    Columns("Z:Z").Select
    Range("Z2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("Z2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-21]=DATEVALUE(TEXT(Input!R20C1,""mm/dd/yy"")),RC[-20]>=Input!R20C2,RC[-20]<=Input!R20C3,RC[1]=Input!R20C6,RC[4]=Input!R20C7),Input!R20C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("Z2:Z" & lngLastRow).FillDown

   '20th shuttle
    Columns("AA:AA").Select
    Range("AA2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AA2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-22]=DATEVALUE(TEXT(Input!R21C1,""mm/dd/yy"")),RC[-21]>=Input!R21C2,RC[-21]<=Input!R21C3,RC[1]=Input!R21C6,RC[4]=Input!R21C7),Input!R21C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("AA2:AA" & lngLastRow).FillDown

    '21th shuttle
    Columns("AB:AB").Select
    Range("AB2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AB2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-23]=DATEVALUE(TEXT(Input!R22C1,""mm/dd/yy"")),RC[-22]>=Input!R22C2,RC[-22]<=Input!R22C3,RC[1]=Input!R22C6,RC[4]=Input!R22C7),Input!R22C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("AB2:AB" & lngLastRow).FillDown

   '22nd shuttle
    Columns("AC:AC").Select
    Range("AC2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AC2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-24]=DATEVALUE(TEXT(Input!R23C1,""mm/dd/yy"")),RC[-23]>=Input!R23C2,RC[-23]<=Input!R23C3,RC[1]=Input!R23C6,RC[4]=Input!R23C7),Input!R23C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("AC2:AC" & lngLastRow).FillDown

    '23rd shuttle
    Columns("AD:AD").Select
    Range("AD2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AD2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-25]=DATEVALUE(TEXT(Input!R24C1,""mm/dd/yy"")),RC[-24]>=Input!R24C2,RC[-24]<=Input!R24C3,RC[1]=Input!R24C6,RC[4]=Input!R24C7),Input!R24C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("AD2:AD" & lngLastRow).FillDown

    '24th shuttle
    Columns("AE:AE").Select
    Range("AE2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AE2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-26]=DATEVALUE(TEXT(Input!R25C1,""mm/dd/yy"")),RC[-25]>=Input!R25C2,RC[-25]<=Input!R25C3,RC[1]=Input!R25C6,RC[4]=Input!R25C7),Input!R25C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("AE2:AE" & lngLastRow).FillDown

    '25th shuttle
    Columns("AF:AF").Select
    Range("AF2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AF2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-27]=DATEVALUE(TEXT(Input!R26C1,""mm/dd/yy"")),RC[-26]>=Input!R26C2,RC[-26]<=Input!R26C3,RC[1]=Input!R26C6,RC[4]=Input!R26C7),Input!R26C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("AF2:AF" & lngLastRow).FillDown

    '26th shuttle
    Columns("AG:AG").Select
    Range("AG2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AG2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-28]=DATEVALUE(TEXT(Input!R27C1,""mm/dd/yy"")),RC[-27]>=Input!R27C2,RC[-27]<=Input!R27C3,RC[1]=Input!R27C6,RC[4]=Input!R27C7),Input!R27C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("AG2:AG" & lngLastRow).FillDown

    '27th shuttle
    Columns("AH:AH").Select
    Range("AH2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AH2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-29]=DATEVALUE(TEXT(Input!R28C1,""mm/dd/yy"")),RC[-28]>=Input!R28C2,RC[-28]<=Input!R28C3,RC[1]=Input!R28C6,RC[4]=Input!R28C7),Input!R28C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("AH2:AH" & lngLastRow).FillDown

   '28th shuttle
    Columns("AI:AI").Select
    Range("AI2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AI2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-30]=DATEVALUE(TEXT(Input!R29C1,""mm/dd/yy"")),RC[-29]>=Input!R29C2,RC[-29]<=Input!R29C3,RC[1]=Input!R29C6,RC[4]=Input!R29C7),Input!R29C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("AI2:AI" & lngLastRow).FillDown

    '29th shuttle
    Columns("AJ:AJ").Select
    Range("AJ2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AJ2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-31]=DATEVALUE(TEXT(Input!R30C1,""mm/dd/yy"")),RC[-30]>=Input!R30C2,RC[-30]<=Input!R30C3,RC[1]=Input!R30C6,RC[4]=Input!R30C7),Input!R30C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("AJ2:AJ" & lngLastRow).FillDown

   '30th shuttle
    Columns("AK:AK").Select
    Range("AK2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AK2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-32]=DATEVALUE(TEXT(Input!R31C1,""mm/dd/yy"")),RC[-31]>=Input!R31C2,RC[-31]<=Input!R31C3,RC[1]=Input!R31C6,RC[4]=Input!R31C7),Input!R31C4,""""),"""")"
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
      Range("AK2:AK" & lngLastRow).FillDown


'Pickup times established based on the shuttles

Columns("AL:AL").Select
    Range("AL2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AL2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC4,RC[-30])"
    Range("AL3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AL2:AL" & lngLastRow).FillDown


Columns("AM:AM").Select
    Range("AM2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AM2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("AM3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AM2:AM" & lngLastRow).FillDown

Columns("AN:AN").Select
    Range("AN2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AN2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("AN3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AN2:AN" & lngLastRow).FillDown

Columns("AO:AO").Select
    Range("AO2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AO2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("AO3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AO2:AO" & lngLastRow).FillDown

Columns("AP:AP").Select
    Range("AP2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AP2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("AP3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AP2:AP" & lngLastRow).FillDown

Columns("AQ:AQ").Select
    Range("AQ2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AQ2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("AQ3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AQ2:AQ" & lngLastRow).FillDown

Columns("AR:AR").Select
    Range("AR2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AR2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("AR3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AR2:AR" & lngLastRow).FillDown

Columns("AS:AS").Select
    Range("AS2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AS2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("AS3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AS2:AS" & lngLastRow).FillDown

Columns("AT:AT").Select
    Range("AT2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AT2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("AT3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AT2:AT" & lngLastRow).FillDown

Columns("AU:AU").Select
    Range("AU2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AU2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("AU3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AU2:AU" & lngLastRow).FillDown

Columns("AV:AV").Select
    Range("AV2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AV2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("AV3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AV2:AV" & lngLastRow).FillDown

Columns("AW:AW").Select
    Range("AW2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AW2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("AW3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AW2:AW" & lngLastRow).FillDown

Columns("AX:AX").Select
    Range("AX2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AX2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("AX3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AX2:AX" & lngLastRow).FillDown

Columns("AY:AY").Select
    Range("AY2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AY2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("AY3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AY2:AY" & lngLastRow).FillDown

Columns("AZ:AZ").Select
    Range("AZ2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AZ2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("AZ3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("AZ2:AZ" & lngLastRow).FillDown

Columns("BA:BA").Select
    Range("BA2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BA2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BA3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BA2:BA" & lngLastRow).FillDown

Columns("BB:BB").Select
    Range("BB2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BB2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BB3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BB2:BB" & lngLastRow).FillDown

Columns("BC:BC").Select
    Range("BC2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BC2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BC3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BC2:BC" & lngLastRow).FillDown

Columns("BD:BD").Select
    Range("BD2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BD2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BD3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BD2:BD" & lngLastRow).FillDown

Columns("BE:BE").Select
    Range("BE2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BE2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BE3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BE2:BE" & lngLastRow).FillDown

Columns("BF:BF").Select
    Range("BF2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BF2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BF3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BF2:BF" & lngLastRow).FillDown

Columns("BG:BG").Select
    Range("BG2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BG2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BG3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BG2:BG" & lngLastRow).FillDown

Columns("BH:BH").Select
    Range("BH2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BH2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BH3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BH2:BH" & lngLastRow).FillDown

Columns("BI:BI").Select
    Range("BI2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BI2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BI3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BI2:BI" & lngLastRow).FillDown

Columns("BJ:BJ").Select
    Range("BJ2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BJ2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BJ3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BJ2:BJ" & lngLastRow).FillDown

Columns("BK:BK").Select
    Range("BK2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BK2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BK3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BK2:BK" & lngLastRow).FillDown

Columns("BL:BL").Select
    Range("BL2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BL2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BL3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BL2:BL" & lngLastRow).FillDown

Columns("BM:BM").Select
    Range("BM2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BM2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BM3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BM2:BM" & lngLastRow).FillDown

Columns("BN:BN").Select
    Range("BN2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BN2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BN3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BN2:BN" & lngLastRow).FillDown

Columns("BO:BO").Select
    Range("BO2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BO2").Select
    ActiveCell.FormulaR1C1 = "=IF(OR(RC[-30]="""",RC[-30]=""#VALUE!""),RC[-1],RC[-30])"
    Range("BO3").Select
    lngLastRow = Range("B" & Rows.Count).End(xlUp).Row
    Range("BO2:BO" & lngLastRow).FillDown

'Remove all formulas
    Columns("H:BO").Select
    Range("H2").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Copy final row to the new pickup column
    Range("BO2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
                
'Delete the rows after the process
        Columns("G:BO").Select
    Range("G2").Activate
    Selection.Delete Shift:=xlToLeft

'Get back 1st row
    Rows("1:1").EntireRow.AutoFit
MsgBox ("You will be prompted to save your input settings as a PDF file")

'Save as PDF
    Sheets("Input").Select

'Inserts Current Date and Time and formats Input
    Range("B11").FormulaR1C1 = "File Saved:"

    Range("B12").Select
    ActiveCell.FormulaR1C1 = "=NOW()"
    Range("B13").Select
    Cells.Select
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
    Selection.Font.Bold = True
    Cells.EntireColumn.AutoFit


    If ActiveWindow.SelectedSheets.Count > 1 Then
        MsgBox "There is more than one sheet selected," & vbNewLine & _
               "and every selected sheet will be published."
    End If

    'Call the function with the correct arguments.
    'You can also use Sheets("Sheet3") instead of ActiveSheet in the code(the sheet does not need to be active then).
    Filename = RDB_Create_PDF(Sheets("Input"), "", True, True)

    If Filename <> "" Then
        'Uncomment the following statement if you want to send the PDF by e-mail.
        'RDB_Mail_PDF_Outlook FileName, "jstanley@savoya.com", "This is the subject", _
           "See the attached PDF file with the last figures" _
          & vbNewLine & vbNewLine & "Regards John Stanley", False
    Else
        MsgBox "It is not possible to create the PDF; possible reasons:" & vbNewLine & _
               "Add-in is not installed" & vbNewLine & _
               "You canceled the GetSaveAsFilename dialog" & vbNewLine & _
               "The path to save the file is not correct" & vbNewLine & _
               "PDF file exists and you canceled overwriting it."
    End If
'End If
'End If


End Sub

'Save as PDF (adapted from Microsoft Code)

Function RDB_Create_PDF(Myvar As Object, FixedFilePathName As String, _
                 OverwriteIfFileExist As Boolean, OpenPDFAfterPublish As Boolean) As String
    Dim FileFormatstr As String
    Dim Fname As Variant

    'Test to see if the Microsoft Create/Send add-in is installed.
    If Dir(Environ("commonprogramfiles") & "\Microsoft Shared\OFFICE" _
         & Format(Val(Application.Version), "00") & "\EXP_PDF.DLL") <> "" Then

        If FixedFilePathName = "" Then
            'Open the GetSaveAsFilename dialog to enter a file name for the PDF file.
            FileFormatstr = "PDF Files (*.pdf), *.pdf"
            Fname = Application.GetSaveAsFilename("", filefilter:=FileFormatstr, _
                  Title:="Create PDF")

            'If you cancel this dialog, exit the function.
            If Fname = False Then Exit Function
        Else
            Fname = FixedFilePathName
        End If

        'If OverwriteIfFileExist = False then test to see if the PDF
        'already exists in the folder and exit the function if it does.
        If OverwriteIfFileExist = False Then
            If Dir(Fname) <> "" Then Exit Function
        End If

        'Now export the PDF file.
        On Error Resume Next
        Myvar.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=Fname, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=OpenPDFAfterPublish
        On Error GoTo 0

        'If the export is successful, return the file name.
        If Dir(Fname) <> "" Then RDB_Create_PDF = Fname
    End If
End Function

Private Sub SS_2()

'Declaring the variables for this manifest
    Dim rCell As Range
    Dim iHours As Integer
    Dim iMins As Integer
    Dim x As Integer
    Dim numrows As Integer

Application.ScreenUpdating = True

numrows = Range("F2", Range("F2").End(xlDown)).Rows.Count
Range("F2").Select

' Establish "For" loop to loop through the flight arrival times.
    For x = 1 To numrows
    For Each rCell In Selection
        If IsNumeric(rCell.Value) And Len(rCell.Value) > 0 Then
            iHours = rCell.Value \ 100
            iMins = rCell.Value Mod 100
            rCell.Value = (iHours + iMins / 60) / 24
            rCell.NumberFormat = "hhmm"
            ActiveCell.Offset(1, 0).Select
        End If
    Next
    Next
End Sub










