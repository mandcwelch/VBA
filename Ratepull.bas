Attribute VB_Name = "Ratepull"
Sub ratePull()

Dim inter As String
Dim dom As String
Dim yr As Integer

If MsgBox("Is this a domestic city?", vbYesNo) = vbNo Then

'shell(windowsexplored: "P:\Operations\Group Department\Pricing\Rates by Market\INTERNATIONAL RATES")
End

Else

dom = InputBox("What city are you looking for? (in 'city, ST' format")
yr = InputBox("What year is this in?")



End Sub
