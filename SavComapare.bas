Attribute VB_Name = "SavComapare"
Sub SavCompare()

'v.1 Application design by John Stanley, jstanley@savoya.com, April 15,2015
'v.2 Added colors to changes and keep the same formatting, April 19,2015

'This macro will collect the two manifests into a new excel sheet, compares both and then highlights the changes in 3 colors.
'Declaring all the variables

Dim customerBook As Workbook
Dim filter As String
Dim caption As String
Dim customerFilename As String
Dim customerWorkbook As Workbook
Dim targetWorkbook As Workbook
Dim w1 As Worksheet, w2 As Worksheet
Dim c As Range, LR As Long, FR As Long

Application.ScreenUpdating = False
'Setting up the new workbook
Set targetWorkbook = Application.ActiveWorkbook

' get the first Savoya manifest in .csv format after the prompt.
MsgBox ("You will now be prompted to enter the 2 Manifest files to be compared." & vbNewLine & "Step 1: Enter OLD Manifest" & vbNewLine & "Step 2: Enter Updated Manifest")

filter = "Text files (*.csv),*.csv"
caption = "Please Select the 1st manifest "
customerFilename = Application.GetOpenFilename(filter, , caption)

Set customerWorkbook = Application.Workbooks.Open(customerFilename)

' assume range is wide range in sheet1. This could be changed later, and copy to sheet1
Dim targetSheet As Worksheet
Set targetSheet = targetWorkbook.Worksheets(1)
Dim sourceSheet As Worksheet
Set sourceSheet = customerWorkbook.Worksheets(1)

targetSheet.Range("A1", "T1000").Value = sourceSheet.Range("A1", "T1000").Value

' Close customer workbook
customerWorkbook.Close
'Add a new Sheet2
Sheets.Add After:=ActiveSheet
ActiveSheet.Select
ActiveSheet.name = "Sheet2"

'get the second Savoya manifest in .csv format after the prompt.
filter = "Text files (*.csv),*.csv"
caption = "Please Select the 2nd manifest "
customerFilename = Application.GetOpenFilename(filter, , caption)

Set customerWorkbook = Application.Workbooks.Open(customerFilename)

' assume range is wide range in sheet1. This could be changed later, and copy to sheet2
'Dim targetSheet As Worksheet
Set targetSheet = targetWorkbook.Worksheets(2)
'Dim sourceSheet As Worksheet
Set sourceSheet = customerWorkbook.Worksheets(1)

targetSheet.Range("A1", "T1000").Value = sourceSheet.Range("A1", "T1000").Value

' Close customer workbook
customerWorkbook.Close
Application.ScreenUpdating = False


'This is the macro to compare and highlight the differences from both the sheets
Set w1 = Worksheets("Sheet1")
Set w2 = Worksheets("Sheet2")

LR = w1.Cells(Rows.Count, 1).End(xlUp).Row

With w1.Range("P1:P" & LR)
  .FormulaR1C1 = "=RC[-15]&RC[-14]&RC[-13]&RC[-12]&RC[-11]&RC[-10]&RC[-9]&RC[-8]&RC[-7]&RC[-6]&RC[-5]&RC[-4]&RC[-3]&RC[-2]&RC[-1]"
  .Value = .Value
End With
LR = w2.Cells(Rows.Count, 1).End(xlUp).Row
With w2.Range("P1:P" & LR)
  .FormulaR1C1 = "=RC[-15]&RC[-14]&RC[-13]&RC[-12]&RC[-11]&RC[-10]&RC[-9]&RC[-8]&RC[-7]&RC[-6]&RC[-5]&RC[-4]&RC[-3]&RC[-2]&RC[-1]"
  .Value = .Value
End With
For Each c In w2.Range("P1", w2.Range("P" & Rows.Count).End(xlUp))
  FR = 0
  On Error Resume Next
  FR = Application.Match(c, w1.Columns(16), 0)
  On Error GoTo 0
  If FR = 0 Then
    c.Offset(, -15).Resize(, 2).Interior.ColorIndex = 27
  End If
Next c

LR = w2.Cells(Rows.Count, 1).End(xlUp).Row

With w2.Range("P1:P" & LR)
  .FormulaR1C1 = "=RC[-15]&RC[-14]&RC[-13]&RC[-12]&RC[-11]&RC[-10]&RC[-9]&RC[-8]&RC[-7]&RC[-6]&RC[-5]&RC[-4]&RC[-3]&RC[-2]&RC[-1]"
  .Value = .Value
End With
LR = w1.Cells(Rows.Count, 1).End(xlUp).Row
With w1.Range("P1:P" & LR)
  .FormulaR1C1 = "=RC[-15]&RC[-14]&RC[-13]&RC[-12]&RC[-11]&RC[-10]&RC[-9]&RC[-8]&RC[-7]&RC[-6]&RC[-5]&RC[-4]&RC[-3]&RC[-2]&RC[-1]"
  .Value = .Value
End With
For Each c In w1.Range("P1", w1.Range("P" & Rows.Count).End(xlUp))
  FR = 0
  On Error Resume Next
  FR = Application.Match(c, w2.Columns(16), 0)
  On Error GoTo 0
  If FR = 0 Then
    c.Offset(, -15).Resize(, 2).Interior.ColorIndex = 27
  End If
Next c

LR = w1.Cells(Rows.Count, 1).End(xlUp).Row

With w1.Range("P1:P" & LR)
  .FormulaR1C1 = "=RC[-15]&RC[-14]"
  .Value = .Value
End With
LR = w2.Cells(Rows.Count, 1).End(xlUp).Row
With w2.Range("P1:P" & LR)
  .FormulaR1C1 = "=RC[-15]&RC[-14]"
  .Value = .Value
End With
For Each c In w2.Range("P1", w2.Range("P" & Rows.Count).End(xlUp))
  FR = 0
  On Error Resume Next
  FR = Application.Match(c, w1.Columns(16), 0)
  On Error GoTo 0
  If FR = 0 Then
    c.Offset(, -15).Resize(, 2).Interior.ColorIndex = 50
  End If
Next c

LR = w2.Cells(Rows.Count, 1).End(xlUp).Row

With w2.Range("P1:P" & LR)
  .FormulaR1C1 = "=RC[-15]&RC[-14]"
  .Value = .Value
End With
LR = w1.Cells(Rows.Count, 1).End(xlUp).Row
With w1.Range("P1:P" & LR)
  .FormulaR1C1 = "=RC[-15]&RC[-14]"
  .Value = .Value
End With
For Each c In w1.Range("P1", w1.Range("P" & Rows.Count).End(xlUp))
  FR = 0
  On Error Resume Next
  FR = Application.Match(c, w2.Columns(16), 0)
  On Error GoTo 0
  If FR = 0 Then
    c.Offset(, -15).Resize(, 2).Interior.ColorIndex = 48
  End If
Next c


w1.Columns(16).ClearContents
w1.Columns(16).Delete Shift:=xlToLeft
w2.Columns(16).ClearContents
w2.Columns(16).Delete Shift:=xlToLeft
Application.ScreenUpdating = True

'Gives the user a note
MsgBox ("GREY - Deletion" & vbNewLine & "GREEN - Addition" & vbNewLine & "YELLOW - Modification")


End Sub



