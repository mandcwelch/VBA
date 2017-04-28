Attribute VB_Name = "QuickCheck"
Sub quickcheck()

Dim cCell As Range


check = InputBox("Please enter what you are looking for")

ischeck = False

For Each cCells In Selection.Cells

If cCells = check Then ischeck = True

Next cCells

If ischeck = False Then MsgBox ("Sorry, there was no " & check & ".")
If ischeck = trie Then MsgBox ("Yes, there is a " & check & ".")


End Sub
