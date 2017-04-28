Attribute VB_Name = "Vendor_Motor_Inquiry"
Option Explicit

Sub Motor_Coach_Email()

'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Updated by Michael Welch 4/27/1 to send out motor inquiries
'Please Contact Michael as mwelch@savoya.com if you experiance any issues
'Working in Excel 2016
    Dim OutApp As Object
    Dim OutMail As Object
    Dim StrBody1 As String
    Dim StrBody2 As String
    Dim direct As String
    Dim endRow As Integer
    Dim i, x As Integer
    Dim vend, vendEmail, ownerEmail As String

    
  Application.ScreenUpdating = False

'Sets up individual emails

endRow = Cells(1, 2).End(xlDown).Row

'starts the loop to email

For x = 2 To endRow
vend = ""
vendEmail = ""


vend = Cells(x, 1)
vendEmail = Cells(x, 2)


'Build the string you want to add as text

    StrBody1 = "Hello " & vend & " Team!" & "<br>" & "<br>" & _
              "I hope 2017 has been going great for you!  I am checking in to confirm mini bus and motor coach rates " & "<br>" _
              & "for the remainder of 2017.  Please provide the following according what you have in your fleet." & "<br>" _


    StrBody2 = "<br>" & "<br>" & _
                "<strong>" & "Mini Buses" & "</strong>" & "<br>" & _
                "Mini Bus Sizes:" & "<br>" & _
                "Hourly Rate:" & "<br>" & _
                "Hourly Minimum:" & "<br>" & _
                "Cancellation Policy:" & "<br>" & _
                "Deposit:" & "<br>" & _
                "<br>" & "<br>" & _
                "<strong>" & "Motor Coaches" & "</strong>" & "<br>" & _
                "Motor Coach Sizes:" & "<br>" & _
                "Hourly Rate:" & "<br>" & _
                "Hourly Minimum:" & "<br>" & _
                "Cancellation Policy:" & "<br>" & _
                "Deposit:" & "<br>" & "<br>" & "<br>" & _
                "Is gratuity provided in the rates above?" & "<br>" & _
                "If not, can that be applied to the final invoice?" & "<br>" & "<br>" & _
                "Are there any additional fees, fuel surcharges, or deposits that are not included above rates?" & "<br>" _
                & "<br>" & "<br>" & "<br>" & "Thanks for your help with those!  Have a great day!"
                

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
        .To = vendEmail
        .CC = ""
        .BCC = ""
        .Subject = "Savoya Mini/Motor Coach Inquiry"
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
    
    Application.ScreenUpdating = True

End Sub
