Attribute VB_Name = "Minutedivision"
Sub Min_Div()

For Each rCell In Selection.Cells

rCell.Value = Round(rCell / 60, 0)

Next rCell

End Sub



