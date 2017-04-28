Attribute VB_Name = "PrivateJetTracking"
Option Explicit

Function InAir(DistNM As Integer, SpotTime As Date)
'   Function will calculate the appropriate time that
'   a flight should depart given DistNM (the distance
'   from airport of origin to destination in nautical
'   miles) and SpotTime (the time our chauffeur is
'   scheduled to make it onsite at the airport). Any
'   flight that hasn't departed by InAir will likely be
'   late to schedule.

        ' If distance is between 0 and 450 NM, then plane
        ' should be in flight at least by scheduled onsite time.
        If DistNM < 451 Then
            InAir = SpotTime
            
            ' If distance is between 451 and 650 NM, then plane
            ' should be in flight at least 30 minutes before onsite time.
            ElseIf DistNM < 651 Then
                InAir = DateAdd("n", -30, SpotTime)
            
                ' If distance is between 651 and 1150 NM, then plane
                ' should be in flight at least 60 minutes before onsite time.
                ElseIf DistNM < 1151 Then
                    InAir = DateAdd("n", -60, SpotTime)
                
                    ' If distance is between 1151 and 1600 NM, then plane
                    ' should be in flight at least 120 minutes before onsite time.
                    ElseIf DistNM < 1601 Then
                        InAir = DateAdd("n", -120, SpotTime)
                    
                        ' If distance is between 1601 and 1750 NM, then plane
                        ' should be in flight at least 150 minutes before onsite time.
                        ElseIf DistNM < 1751 Then
                            InAir = DateAdd("n", -150, SpotTime)
        
                            ' If distance is 1751 NM or more, then plane
                            ' should be in flight at least 240 minutes before onsite time.
                            Else
                                InAir = DateAdd("n", -240, SpotTime)
                                
        End If
    
End Function
Function Plan(DistNM As Integer, SpotTime As Date)
'   Function will calculate the latest time that
'   a flight should file a plan given DistNM (the distance
'   from airport of origin to destination in nautical
'   miles) and SpotTime (the time our chauffeur is
'   scheduled to make it onsite at the airport).  Any flight
'   that doesn't have a Plan filed by this time will likely
'   be late to schedule.

        ' If distance is between 0 and 450 NM, then plan
        ' should be filed at least 60 minutes before take off.
        If DistNM < 451 Then
            Plan = DateAdd("n", -60, SpotTime)
            
            ' If distance is between 451 and 650 NM, then plan
            ' should be filed at least 90 minutes before take off.
            ElseIf DistNM < 651 Then
                Plan = DateAdd("n", -90, SpotTime)
            
                ' If distance is between 651 and 1150 NM, then plan
                ' should be filed at least 120 minutes before take off.
                ElseIf DistNM < 1151 Then
                    Plan = DateAdd("n", -120, SpotTime)
                
                    ' If distance is between 1151 and 1600 NM, then plan
                    ' should be filed at least 150 minutes before take off.
                    ElseIf DistNM < 1601 Then
                        Plan = DateAdd("n", -150, SpotTime)
                    
                        ' If distance is between 1601 and 1750 NM, then plan
                        ' should be filed at least 180 minutes before take off.
                        ElseIf DistNM < 1751 Then
                            Plan = DateAdd("n", -180, SpotTime)
                            
                            ' If distance is 1751 NM or more, then plan
                            ' should be filed at least 240 minutes before take off.
                            Else
                                Plan = DateAdd("n", -240, SpotTime)
                                
        End If
    
End Function

Sub Private_Flight_Tracking()
'
' Private_Flight_Tracking Macro
'
' Keyboard Shortcut: Ctrl+q
'

' PrivateFlightTracking Macro Version 1
' Format private flight tracking report.
' Originally Created by Nicole Livingston on August 5, 2013.
' Updated by Michael Welch on July 27, 2015
    'Fixed ETA column to read correctly
    'Fixed link to distance file
    'Updated time formating to keep info
'
' To use:
'   1.  Go to Dispatch Main Board in the Savoya Internal System.
'   2.  On the top of the page, hover your mouse over the "Operations" menu.
'   3.  Click on "Reporting"
'   4.  Make sure the date range is correct.
'   5.  Select Report Type: "Reservation History"
'   6.  Select Report Name: "Savoya - Daily Trips of Private Air Tail#"
'   7.  Click on "Run Report"
'   8.  Copy the table that is generated and paste into a new Excel workbook.
'   9.  Press Ctrl+q to automatically format the Table
'

Application.EnableEvents = False
Application.ScreenUpdating = False
Application.DisplayAlerts = False

' Create Sheet2
ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)

' Select Sheet1
Worksheets("Sheet1").Activate

Dim cell As Range

    ' Delete rows that are quotes, not confirmed reservations
    With ActiveSheet
        .AutoFilterMode = False
        With Range("H1", Range("H" & Rows.Count).End(xlUp))
            .AutoFilter 1, "Quote"
            On Error Resume Next
            .Offset(1).SpecialCells(12).EntireRow.Delete
        End With
        .AutoFilterMode = False
    End With

     'Select entire current sheet
    Cells.Select
    'Unformat sheet completely
    Selection.Style = "Normal"
    
    'Add a new column for "Notes"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    'Add a new column for "ETA"
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    'Add a new column for "Tracking"
    Columns("J:J").Select
    Selection.ClearContents
    
    'Name the Columns
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Notes"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Name"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Spot"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "ETA"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Tracking"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Origin"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Arrival"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Distance"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Plan"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "InAir"
    
         
 ' Select all cells with data.
    ActiveSheet.Range("A1").CurrentRegion
    
  
  
    ' Name the table "Table1" and format it using
    ' Excel's built-in Table Style "Medium13"
    With ActiveSheet.ListObjects.Add(xlSrcRange, ActiveSheet _
        .Range("A1").CurrentRegion, , xlYes)
        .name = "Table1"
        .TableStyle = "TableStyleMedium13"
    End With
    
    ' Select all cells again
    Cells.Select
    
    ' Format Font and alignment
    With Selection.Font
        .name = "Calibri"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' Select column D, "Spot".  Format as a time "hh:mm".
    Columns("D:D").Select
    Selection.NumberFormat = "h:mm;@"
    Selection.Replace What:=" PM", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        
    ' Select columns E, "Tail", and H, "Origin".
    ' Remove all "TBA", "TBD", and "N/A" placeholders.
    Range("E:E,H:H").Select
    
    Selection.Replace What:="TBA", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="TBD", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="N/A", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    ' Select columns E, "Tail", and F, "ETA". Change font size to 12pt.
    ' and make bold for better visibility.
    Columns("E:F").Select
    Selection.Font.Size = 12
    Selection.Font.Bold = True
    
    ' Apply formula "Plan" to column K, "Plan".
    Range("K2").Formula = "=PERSONAL.XLSB!Plan(J2,D2)"

    ' Format column K, "Plan", to time "hh:mm".
    Columns("K:K").Select
    Selection.NumberFormat = "h:mm;@"
    
    ' Apply formula "InAir" to column L, "InAir".
    Range("L2").Formula = "=PERSONAL.XLSB!InAir(J2,D2)"

     ' Format column L, "InAir", to time "hh:mm".
    Columns("L:L").Select
    Selection.NumberFormat = "h:mm;@"
    
    ' Make all letters in column E, "Tail", uppercase.
    ' Also, if "Tail" begins with an integer, then add the
    ' letter "N" to the beginning of it.
  
    Dim lastColE As Integer
    Dim lastrowE As Integer
   
    
    lastColE = ActiveSheet.Range("E2").End(xlDown).Column
    
    
    lastrowE = ActiveSheet.Cells(65536, lastColE).End(xlUp).Row
    
    
    For Each cell In _
        ActiveSheet.Range("E2", ActiveSheet.Cells(lastrowE, lastColE))
        cell.Value = UCase(cell.Value)
        If Left(cell.Text, 1) Like "#" = "True" Then
            cell.Value = "N" & cell
        Else
            cell.Value = cell
        End If
    Next cell
    
    
    ' Name the tail column "Tail"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Tail"
    
    ' Make all letters in column H, "Origin", uppercase.
    ' Also, remove "K" from begining of origins in the USA.
    Dim LastColH As Integer
    Dim lastrowH As Integer
    
    LastColH = ActiveSheet.Range("H2").End(xlToRight).Column
    
    
    lastrowH = ActiveSheet.Cells(65536, LastColH).End(xlUp).Row

    For Each cell In _
        ActiveSheet.Range("H2", ActiveSheet.Cells(lastrowH, LastColH))
        cell.Value = UCase(cell.Value)
        If Left(cell.Value, 1) = "K" And Len(cell) = "4" Then
            cell.Value = Right(cell.Text, 3)
        Else
            cell.Value = cell
        End If
    Next cell

    ' Wrap text in "Notes" column
    Columns("A:A").Select
    With Selection
        .WrapText = True
    End With

    ' Fit all cells to contents
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    ' Use saved distance data to fill in "Distance" Column
    Dim wb As Workbook, wb2 As Workbook
    Dim LastColB1 As Integer
    Dim LastColH1 As Integer
    Dim LastColI1 As Integer
    Dim LastColJ1 As Integer
    Dim LastRow1 As Integer
    
    Set wb = ActiveWorkbook
    Set wb2 = Workbooks.Open("P:\Operations\Operations Private Flight Tracking\FlightDistances1.XLSM", True, True)
    
    wb.Worksheets("Sheet2").Range("A1:B30000").Value = _
        wb2.Worksheets("Sheet1").Range("A1", Worksheets("Sheet1"). _
        Range("A1").End(xlDown).End(xlToRight)).Value

    LastColB1 = wb.Worksheets("Sheet1").Range("B2").End(xlToRight).Column
    LastColH1 = wb.Worksheets("Sheet1").Range("H2").End(xlToRight).Column
    LastColI1 = wb.Worksheets("Sheet1").Range("I2").End(xlToRight).Column
    LastColJ1 = wb.Worksheets("Sheet1").Range("J2").End(xlToRight).Column
    LastRow1 = wb.Worksheets("Sheet1").Cells(65536, LastColB1).End(xlUp).Row

    Dim RefRng As Range
    Set RefRng = wb.Worksheets("Sheet2").Range("A2", wb.Worksheets("Sheet2") _
        .Range("A2").End(xlDown).End(xlToRight))
    
    Dim HCol As Range
    Set HCol = wb.Worksheets("Sheet1").Range("H2", wb.Worksheets("Sheet1") _
        .Cells(LastRow1, LastColH1))
    
    Dim iCol As Range
    Set iCol = wb.Worksheets("Sheet1").Range("I2", wb.Worksheets("Sheet1") _
        .Cells(LastRow1, LastColI1))
    
    Dim JCol As Range
    Set JCol = wb.Worksheets("Sheet1").Range("J2", wb.Worksheets("Sheet1") _
        .Cells(LastRow1, LastColJ1))
    
    Dim NewRng As Range
    Set NewRng = Application.Union(HCol, iCol, JCol)
    
    Dim NewRows As Long
    NewRows = NewRng.Rows.Count
    
    Dim j As Integer
    j = 1
    
    Dim K As Integer
    K = 1 + NewRows
    
    Dim x As Variant
    
        Do Until j = K
            x = Application.VLookup((NewRng.Cells(j, 1) & NewRng.Cells(j, 2)), RefRng, 2, False)
                    
                If IsError(x) = False Then
                    NewRng(j, 3) = x
                Else
                    NewRng(j, 3).ClearContents
                End If
                j = j + 1
            Loop
    
    wb.Worksheets("Sheet2").Range("A1:B30000").ClearContents
    wb2.Close
    
    
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.DisplayAlerts = True
    
  
    'Fixes time format issue
    
Cells.Find(What:="0.", After:=[A1], LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False).Activate
    
    ActiveCell.NumberFormat = ("hhmm")
    
Cells.Find(What:="0.", After:=[A1], LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False).Activate
    
    ActiveCell.NumberFormat = ("hhmm")
    
Cells.Find(What:="0.", After:=[A1], LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False).Activate
    
    ActiveCell.NumberFormat = ("hhmm")
    
Cells.Find(What:="0.", After:=[A1], LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False).Activate
    
    ActiveCell.NumberFormat = ("hhmm")
    
Cells.Find(What:="0.", After:=[A1], LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False).Activate
    
    ActiveCell.NumberFormat = ("hhmm")
    
Cells.Find(What:="0.", After:=[A1], LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False).Activate
    
    ActiveCell.NumberFormat = ("hhmm")
    
Cells.Find(What:="0.", After:=[A1], LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False).Activate
    
    ActiveCell.NumberFormat = ("hhmm")
    
Cells.Find(What:="0.", After:=[A1], LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False).Activate
    
    ActiveCell.NumberFormat = ("hhmm")
    
Cells.Find(What:="0.", After:=[A1], LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False).Activate
    
    ActiveCell.NumberFormat = ("hhmm")
    
Cells.Find(What:="0.", After:=[A1], LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False).Activate
    
    ActiveCell.NumberFormat = ("hhmm")

    
    
    
End Sub








