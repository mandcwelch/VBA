Attribute VB_Name = "changetime"
Sub changetime()
Dim cell As Range
Dim rng As Range

Set rng = Selection
Selection.NumberFormat = general
For Each cell In rng.Cells

lng = Len(cell)

' Midnight numbers

If Left(cell, 2) = "12" And Right(cell, 1) = "A" And Len(cell) = 5 Then
    
    cell = Left(cell, lng - 1)
    cell.Value = cell.Value - 1200
    cell.NumberFormat = "0000"
End If

If Left(cell, 2) = "12" And Right(cell, 1) = "a" And Len(cell) = 5 Then
    
    cell = Left(cell, lng - 1)
    cell.Value = cell.Value - 1200
    cell.NumberFormat = "0000"
End If

' AM numbers

If Right(cell, 1) = "A" Then cell = Left(cell, lng - 1)
If Right(cell, 1) = "a" Then cell = Left(cell, lng - 1)

'Noon numbers

If Left(cell, 2) = "12" And Right(cell, 1) = "P" Then cell = Left(cell, lng - 1)
If Left(cell, 2) = "12" And Right(cell, 1) = "p" Then cell = Left(cell, lng - 1)

' PM Numbers

If Right(cell, 1) = "P" Then

    cell = Left(cell, lng - 1)

    cell.Value = cell.Value + 1200
    
End If

If Right(cell, 1) = "p" Then

    cell = Left(cell, lng - 1)

    cell.Value = cell.Value + 1200
    
End If

Next cell

End Sub
