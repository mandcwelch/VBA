Attribute VB_Name = "QuickCompare"
Sub quickcompare()

Dim i As Integer
Dim col1 As Integer
Dim col2 As Integer
Dim endrange As Integer

col1 = InputBox("What is the number of the first column to check")

col2 = InputBox("What is the number of the second column to check")

Cells(2, col1).Select
endrange = Selection.End(xlDown).Row

For i = 2 To endrange

If Cells(i, col1) <> Cells(i, col2) Then

Cells(i, col1).Interior.color = RGB(250, 255, 0)
Cells(i, col2).Interior.color = RGB(250, 255, 0)

End If

Next i

End Sub
