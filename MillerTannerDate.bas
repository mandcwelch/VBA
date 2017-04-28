Attribute VB_Name = "MillerTannerDate"
Sub MillerTannerDate()

Dim rCell As Range
    Dim rRng As Range

Selection.NumberFormat = Text
    
    Set rRng = Selection

    For Each rCell In rRng.Cells



Selection.Replace What:="Jan", Replacement:="01", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="Feb", Replacement:="02", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="Mar", Replacement:="03", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="Apr", Replacement:="04", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="May", Replacement:="05", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="Jun", Replacement:="06", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="Jul", Replacement:="07", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="Aug", Replacement:="08", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="Sep", Replacement:="09", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="Oct", Replacement:="10", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="Nov", Replacement:="11", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="Dec", Replacement:="12", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
Selection.Replace What:="-", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

rCell = (Mid(rCell, 3, 2)) & "/" & (Left(rCell, 2)) & "/" & (Right(rCell, 2))

Next rCell


End Sub
