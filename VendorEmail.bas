Attribute VB_Name = "VendorEmail"
Sub Vendor_email()
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Updated by Michael Welch 10/11/16 to organize and send schedule emails
'Please Contact Michael as mwelch@savoya.com if you experiance any issues
'Working in Excel 2016
    Dim OutApp As Object
    Dim OutMail As Object
    Dim StrBody1 As String
    Dim StrBody2 As String
    Dim direct As String
    Dim endRow As Integer
    Dim i, x As Integer

    
  Application.ScreenUpdating = False
    

'Creates workbooks to send to
    
    
'Assigns Director
    
    direct = InputBox("Who is sending this?")
    
'copies main schedule to the vendor list for organizing
    

'Sets up individual emails


endRow = Cells(1, 2).End(xlDown).Row


'starts the loop to email

For x = 2 To endRow
vendEmail = Cells(x, 5)

vend = Cells(x, 1)


'Build the string you want to add as text

    StrBody1 = "Hello " & vend & " Team!" & "<br>" & "<br>" & _
              "I just wanted to check in with you regarding sedans in your fleet." & "<br>" _
              & "We wanted to make sure that we had your vehicles down correctly in our system." & "<br>" _
              & "Could you please confirm if you have E-Class sedans, or let us know if you are useing something else?" & _
              "<br>" & "Also, please confirm your sedan rate below so that we know we are up to date.  Thanks!"

    StrBody2 = "<br>" & "<br>" & _
                "E-Class or other Sedan:" & "<br>" & _
                "Hourly Rate:" & "<br>" & _
                "Hourly Minimum:" & "<br>" & _
                "Cancellation Policy:" & "<br>" & "<br>" & _
                "Regards," & "<br>" & "<br>" & _
                direct




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
        .display
        '.SentOnBehalfOfName = "vendors@savoya.com"
        .To = "mwelch@savoya.com" 'vendemail
        .CC = ""
        .BCC = ""
        .Subject = "Savoya Sedan Inquiry"
        .HTMLBody = StrBody1 & StrBody2 & .HTMLBody
        .send   'or use .Display
    End With
    
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
    

Next x
    
    'closes temp files and returns screen to normal
    
    
    Application.ScreenUpdating = True
    
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

Columns("A:A").AutoFit

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


Set empRng = Range(Cells(1, 1), Cells(schLen, 26))


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



