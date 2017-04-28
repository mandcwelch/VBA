Attribute VB_Name = "beepbeep"
Sub beepbeep()
'
Application.Wait (Now + TimeValue("0:00:10"))
Beep
Application.Wait (Now + TimeValue("0:00:10"))
Beep
Application.Wait (Now + TimeValue("0:00:09"))
Beep
Application.Wait (Now + TimeValue("0:00:09"))
Beep
Application.Wait (Now + TimeValue("0:00:08"))
Beep
Application.Wait (Now + TimeValue("0:00:07"))
Beep
Application.Wait (Now + TimeValue("0:00:06"))
Beep
Application.Wait (Now + TimeValue("0:00:05"))
Beep
Application.Wait (Now + TimeValue("0:00:03"))
Beep
Application.Wait (Now + TimeValue("0:00:01"))
Beep
Application.Wait (Now + TimeValue("0:00:01"))
Beep
Application.Wait (Now + TimeValue("0:00:01"))
Beep
Application.Wait (Now + TimeValue("0:00:01"))
Beep
Beep
Application.Wait (Now + TimeValue("0:00:01"))
Beep
Beep
Application.Wait (Now + TimeValue("0:00:01"))
Beep
Beep
Application.Wait (Now + TimeValue("0:00:01"))
Beep
Application.Wait (Now + TimeValue("0:00:00"))
Beep
Application.Wait (Now + TimeValue("0:00:00"))
Beep
Application.Wait (Now + TimeValue("0:00:00"))
Beep
Application.Wait (Now + TimeValue("0:00:00"))
Beep
Application.Wait (Now + TimeValue("0:00:00"))
Beep
Application.Wait (Now + TimeValue("0:00:01"))
Beep

End Sub


Sub Beep2()

For i = 1 To 100

If Cells(i, 1) = 1 Then
Beep
Application.Wait (Now + TimeValue("0:00:01"))

ElseIf Cells(i, 1) = 2 Then
Beep
Application.Wait (Now + TimeValue("0:00:02"))

ElseIf Cells(i, 1) = 3 Then
Beep
Application.Wait (Now + TimeValue("0:00:03"))

End If

Next i

End Sub

