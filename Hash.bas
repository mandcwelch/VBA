Attribute VB_Name = "Hash"
Sub HashColumn()
Attribute HashColumn.VB_ProcData.VB_Invoke_Func = "h\n14"
'
' HashColumn Macro
'
' Keyboard Shortcut: Ctrl+h
'
Dim i As Integer
Dim x As Integer
Dim endrange As Integer

endrange = Cells(2, 2).End(xlDown).Row

   Range(Cells(2, 2), Cells(endrange, 2)).TextToColumns destination:=Range("B2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

Range(Cells(2, 1), Cells(endrange, 1)).Value = Range(Cells(2, 3), Cells(endrange, 3)).Value
Range(Cells(2, 3), Cells(endrange, 3)).Delete

For i = 2 To endrange

x = Len(Cells(i, 1))
If Right(Cells(i, 1), 3) = " Mr" Then Cells(i, 1) = Left(Cells(i, 1), x - 3)
If Right(Cells(i, 1), 3) = " MR" Then Cells(i, 1) = Left(Cells(i, 1), x - 3)
If Right(Cells(i, 1), 3) = " Dr" Then Cells(i, 1) = Left(Cells(i, 1), x - 3)
If Right(Cells(i, 1), 3) = " DR" Then Cells(i, 1) = Left(Cells(i, 1), x - 3)
If Right(Cells(i, 1), 3) = " Ms" Then Cells(i, 1) = Left(Cells(i, 1), x - 3)
If Right(Cells(i, 1), 3) = " MS" Then Cells(i, 1) = Left(Cells(i, 1), x - 3)
If Right(Cells(i, 1), 4) = " Mrs" Then Cells(i, 1) = Left(Cells(i, 1), x - 4)
If Right(Cells(i, 1), 4) = " MRS" Then Cells(i, 1) = Left(Cells(i, 1), x - 4)

Next i

End Sub
