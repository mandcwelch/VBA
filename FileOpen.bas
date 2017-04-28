Attribute VB_Name = "FileOpen"
Sub FileOpen()
Dim userid As String

userid = InputBox("Please enter your user ID")

Call Shell("explorer.exe C:\Users\" & userid & "\AppData\Roaming\Microsoft\AddIns", vbNormalFocus)

End Sub

Sub FileUpload()
Dim userid As String
Dim try As Boolean

On Error GoTo fileerror

userid = InputBox("Please enter your user ID" & vbCrLf & "(Your first initial and last name)")

FileCopy "P:\Operations\Group Department\Macros\AddIn\UpdateAddin.xlam", _
"C:\Users\" & userid & "\AppData\Roaming\Microsoft\AddIns\Updateaddin.xlam"

MsgBox ("To install AddIn, click on the File tab and go to Options/Customize Ribbon. On the right hand side click the Developer check box.  Go back to the workbook, select the Developer tab and click on Add Ins.  Select the updateaddin box and click okay.  The Add Ins tab should open with all your macros.")

Exit Sub

fileerror:

MsgBox ("Your ID did not work.  Please try again.")

FileUpload2


End Sub


Sub FileUpload2()
Dim userid As String
Dim try As Boolean

On Error GoTo fileerror2
try:

userid = InputBox("Please enter your user ID" & vbCrLf & "(Your first initial and last name)")

FileCopy "P:\Operations\Group Department\Macros\AddIn\UpdateAddin.xlam", _
"C:\Users\" & userid & "\AppData\Roaming\Microsoft\AddIns\Updateaddin.xlam"

MsgBox ("To install AddIn, click on the File tab and go to Options/Customize Ribbon. On the right hand side click the Developer check box.  Go back to the workbook, select the Developer tab and click on Add Ins.  Select the updateaddin box and click okay.  The Add Ins tab should open with all your macros.")

Exit Sub


fileerror2:

MsgBox ("That ID is not working.  Please check in with Michael to have Add In installed.")

Exit Sub

End Sub




