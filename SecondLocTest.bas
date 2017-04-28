Attribute VB_Name = "SecondLocTest"
Sub Transfer_Distance()

Dim tot As Range

If Range("A1") = "" Then
MsgBox ("Please paste the cities, with state if possible, in the leftmost column and the airport addresses in the top row, with the Service Area Name in the first cell.")
End
End If

On Error Resume Next

'Application.ScreenUpdating = False

endCol = Cells(1, 1).End(xlToRight).Column
endRow = Cells(1, 1).End(xlDown).Row

'check each column

'Set tot = Range(Cells(1, 1), Cells(endrow, endCol))
'tot.Replace What:=" ", Replacement:="+", LookAt:=xlPart, _
 '   SearchOrder:=xlByRows, MatchCase:=False

For i = 2 To endCol

origin = Cells(1, i)



    For x = 2 To endRow

    destination = Cells(x, 1)

    xmlstrdur = "URL;https://maps.googleapis.com/maps/api/directions/xml?origin=" & origin & ",&destination=" & destination & "&alternatives=false?key=AIzaSyA6LSCQbtuOfDnDH0NCbH_WaYDjCcNQ0cU"

'g = InputBox("Hello", , xmlstrdur)

'Shell ("C:\Program Files (x86)\Google\Chrome\Application\Chrome.exe -url https://maps.googleapis.com/maps/api/directions/xml?origin=" & origin & ",&destination=" & destination & "&alternatives=false?key=AIzaSyA6LSCQbtuOfDnDH0NCbH_WaYDjCcNQ0cU")
'Shell ("C:\Users\USERNAME\AppData\Local\Google\Chrome\Application\Chrome.exe -url https://www.gmail.com")
    
 '   With ActiveSheet.QueryTables.Add(Connection:=xmlstrdur, destination:=Range("M1"))


 '   .name = "1"
  '  .FieldNames = True
   ' .RowNumbers = False
    '.FillAdjacentFormulas = False
'    .PreserveFormatting = True
 '   .RefreshOnFileOpen = False
  '  .BackgroundQuery = True
   ' .RefreshStyle = xlInsertDeleteCells
    '.SavePassword = False
'    .SaveData = True
 '   .AdjustColumnWidth = True
  '  .RefreshPeriod = 0
   ' .WebSelectionType = xlSpecifiedTables
    '.WebFormatting = xlWebFormattingNone
'    .WebTables = "2"
 '   .WebPreFormattedTextToColumns = True
  '  .WebConsecutiveDelimitersAsOne = True
   ' .WebSingleBlockTextImport = False
    '.WebDisableDateRecognition = False
'    .WebDisableRedirections = False
 '   .Refresh BackgroundQuery:=False
 
'     ActiveWorkbook.XmlImport URL:= _
 '       "https://maps.googleapis.com/maps/api/directions/xml?origin=" & origin & ",&destination=" & destination & "&alternatives=false?key=AIzaSyA6LSCQbtuOfDnDH0NCbH_WaYDjCcNQ0cU" _
  '      , ImportMap:=Nothing, Overwrite:=True, destination:=Range("$M$1")
    '  ActiveWorkbook.XmlImport URL:= _
        "https://maps.googleapis.com/maps/api/directions/xml?origin=Frankfurt Airport,&destination=Wiesbaden, Hesse&alternatives=false?key=AIzaSyA6LSCQbtuOfDnDH0NCbH_WaYDjCcNQ0cU" _
        , ImportMap:=Nothing, Overwrite:=True, destination:=Range("$M$1")
     ActiveWorkbook.XmlImport URL:= _
        "http://dev.virtualearth.net/REST/v1/Routes? wayPoint.1=Detroit&viaWaypoint.2=Chicago&heading=heading&optimize=optimize&avoid=avoidOptions&distanceBeforeFirstTurn=distanceBeforeFirstTurn&routeAttributes=routeAttributes&maxSolutions=maxSolutions&tolerances=tolerance1,tolerance2,tolerancen&distanceUnit=distanceUnit&mfa=mfa&key=BingMapsKey" _
        , ImportMap:=Nothing, Overwrite:=True, destination:=Range("$M$1")
 
    'End With

    If Range("M2") <> "OK" Then
    Cells(x, i) = "Error"
    Else:
    'dur = Cells.Find("/route/leg/duration/value").Column
    Cells(x, i) = Range("AC2")
    'Cells(x, i) = Cells(10, dur)
    Cells(x, i) = Round((Cells(x, i) / 1609.344), 2)
    End If
    Columns("M:BK").Delete


'Sleep 500
'Cells(4, 5) = xmlstrdur
Application.Wait (Now + TimeValue("00:00:01"))

'Application.Wait (Now + 1 / (24 * 60 * 60))

Next x

Next i

tot.Replace What:="+", Replacement:=" ", LookAt:=xlPart, _
       SearchOrder:=xlByRows, MatchCase:=False

Application.ScreenUpdating = True

End Sub



