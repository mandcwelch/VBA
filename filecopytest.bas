Attribute VB_Name = "filecopytest"
Sub filecopytest()
Dim userid As String

'tryAgain:

userid = InputBox("Please enter your user ID" & vbCrLf & "(Your first initial and last name)")

'On Error GoTo fileerror

 
    FileCopy "P:\Operations\Group Department\Macros\AddIn\SavoyaAddIn.xlam", "C:\Users\AppData\Roaming\Microsoft\AddIns\SavoyaAddIn.xlam"

MsgBox ("To install AddIn, click on the File tab and go to Options/Customize Ribbon. On the right hand side click the Developer check box.  Go back to the workbook, select the Developer tab and click on Add Ins.  Select the updateaddin box and click okay.  The Add Ins tab should open with all your macros.")

'Exit Sub

'fileerror:
'MsgBox ("Your ID did not work.  Please try again")
'GoTo tryAgain

End Sub
