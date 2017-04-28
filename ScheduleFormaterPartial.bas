Attribute VB_Name = "ScheduleFormaterPartial"
Sub schedule_format()
'Version 1 created 10/10/2016 by Michael Welch
'Formats schedule for sending out to agents
Dim i As Integer
Dim TotTab As Range
Dim init As Workbook
Dim main As Workbook

'creates copy of schedule

Set init = ActiveWorkbook
Set main = Workbooks.Add

main.theme.ThemeColorScheme.Load ("P:\Operations\Group Department\Macros\theme")
    
    init.Activate
    
    ActiveSheet.UsedRange.Copy
    
    main.Activate
    
    Range("A1").Insert

'Sets Day of Week Range

SundayStartrow = 6
SundayEndrow = Cells(SundayStartrow, 2).End(xlDown).Row

MondayStartrow = SundayEndrow + 7
MondayEndrow = Cells(MondayStartrow, 2).End(xlDown).Row

TuesdayStartrow = MondayEndrow + 7
TuesdayEndrow = Cells(TuesdayStartrow, 2).End(xlDown).Row

WednesdayStartrow = TuesdayEndrow + 7
WednesdayEndrow = Cells(WednesdayStartrow, 2).End(xlDown).Row

ThursdayStartrow = WednesdayEndrow + 7
ThursdayEndrow = Cells(ThursdayStartrow, 2).End(xlDown).Row

FridayStartrow = ThursdayEndrow + 7
FridayEndrow = Cells(FridayStartrow, 2).End(xlDown).Row

SaturdayStartrow = FridayEndrow + 7
SaturdayEndrow = Cells(SaturdayStartrow, 2).End(xlDown).Row


' Adds Day of Week


For i = SundayStartrow To SundayEndrow

Cells(i, 1) = "Sunday"

Next i


For i = MondayStartrow To MondayEndrow

Cells(i, 1) = "Monday"

Next i


For i = TuesdayStartrow To TuesdayEndrow

Cells(i, 1) = "Tuesday"

Next i


For i = WednesdayStartrow To WednesdayEndrow

Cells(i, 1) = "Wednesday"

Next i


For i = ThursdayStartrow To ThursdayEndrow

Cells(i, 1) = "Thursday"

Next i


For i = FridayStartrow To FridayEndrow

Cells(i, 1) = "Friday"

Next i


For i = SaturdayStartrow To SaturdayEndrow

Cells(i, 1) = "Saturday"

Next i


' Removes Blanks

On Error Resume Next

Rows("1:2").Delete

    Range(Cells(2, 1), Cells(1000, 1)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    

'Orders rows by day of week

endrow1 = Cells(2, 1).End(xlDown).Row

    ActiveSheet.Sort.SortFields.Add Key:= _
        Range(Cells(2, 2), Cells(endrow1, 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:= _
        Range(Cells(2, 1), Cells(endrow1, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday", DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(Cells(2, 1), Cells(endrow1, 26))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With



If MsgBox("Do you wish to email these now?", vbYesNo) = vbYes Then Mail_Selection_Range_Outlook_Body


End Sub



Sub Mail_Selection_Range_Outlook_Body()
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Updated by Michael Welch 10/11/16
'Don't forget to copy the function RangetoHTML in the module.
'Working in Excel 2000-2016
    Dim empRng As Range
    Dim rng As Range
    Dim OutApp As Object
    Dim OutMail As Object
    Dim main As Workbook
    Dim init As Workbook
    
  Application.ScreenUpdating = False
    
    Set init = ActiveWorkbook
    Set main = Workbooks.Add
    
'init.Theme.ThemeColorScheme.Save ("C:\Users\mwelch\Desktop\theme")
main.theme.ThemeColorScheme.Load ("P:\Operations\Group Department\Macros\theme")
    
    init.Activate
    
    ActiveSheet.UsedRange.Copy
    
    main.Activate
    
    Range("A1").Insert
   ' Set main = ActiveWorkbook
   
   
'To add text to email
    
    Dim strbody As String

Set emailMain = Workbooks.Open("P:\Operations\Group Department\Macros\EmployeeEmail.csv")
main.Activate

'Sets up individual emails



endRow = Cells(1, 2).End(xlDown).Row

empTot = 0

'set number of employees

For i = 2 To endRow

If Cells(i, 2) <> Cells(i - 1, 2) Then empTot = empTot + 1

Next i

'emails the schedule

For x = 1 To empTot

schLen = 1

emp = Cells(2, 2)

For i = 2 To endRow

If Cells(i, 2) = emp Then

schLen = schLen + 1

Else: GoTo empset

End If

Next i

empset:


Set empRng = Range(Cells(1, 2), Cells(schLen, 26))


'Build the string you want to add

    StrBody1 = "Howdy " & emp & "," & "<br>" & "<br>" & _
              "Below are your assignments for the week.  Let me know if you have any questions." & "<br>" & "<br>" & "<br>"

    StrBody2 = "<br>" & "<br>" & _
                "Gray - Phones and To Do's" & "<br>" & _
                "Green - Portal" & "<br>" & _
                "Blue - CS" & "<br>" & _
                "Purple - Meetings" & "<br>" & "<br>" & _
                "Regards," & "<br>" & "<br>" & _
                "Josh"

'grab email

emailMain.Activate

endrowEmail = Cells(1, 1).End(xlDown).Row

For i = 1 To endrowEmail

If Cells(i, 1) = emp Then

empEmail = Cells(i, 3)

GoTo endeMail

End If

Next i

endeMail:

main.Activate

If empEmail = "" Then empEmail = InputBox("Email not found.  Please enter email address")

For i = 2 To schLen

Cells(i, 2) = Cells(i, 1)

Next i

Columns("B:B").AutoFit

' send email

    Set rng = Nothing
    On Error Resume Next
    'Only the visible cells in the selection
    Set rng = empRng '.SpecialCells(xlCellTypeVisible)
    'You can also use a fixed range if you want
    'Set rng = Sheets("YourSheet").Range("D4:D12").SpecialCells(xlCellTypeVisible)
    'On Error GoTo 0

    'If rng Is Nothing Then
     '   MsgBox "The selection is not a range or the sheet is protected" & _
     '        vbNewLine & "please correct and try again.", vbOKOnly
     '  Exit Sub
   ' End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = empEmail
        .CC = ""
        .BCC = ""
        .Subject = "Assignments for Next Week"
        .HTMLBody = StrBody1 & RangetoHTML(rng) & StrBody2
        .send   'or use .Display
    End With
    
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
    
    Rows("2:" & schLen & "").Delete
    
Next x
    
    emailMain.Close
    main.Close savechanges:=False
    
    Application.ScreenUpdating = True
    
End Sub


Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
    


    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    
    TempWB.theme.ThemeColorScheme.Load ("P:\Operations\Group Department\Macros\theme")
    
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
       ' .Theme.ThemeColorScheme.Load ("C:\Users\mwelch\Desktop\theme")
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function



