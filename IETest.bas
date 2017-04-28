Attribute VB_Name = "IETest"
Sub IETest()
Dim ie As New InternetExplorer
ie.Visible = True

Do
ie.navigate "http://www.unitedstateszipcodes.org/" & Range("B1").Value & "/"
Loop Until ie.readyState = READYSTATE_COMPLETE

Dim doc As HTMLDocument
Set doc = ie.document

Dim county As String
Dim city As String

county = doc.getElementsByTagName("td")(1).innerText
city = doc.getElementsByTagName("td")(0).innerText
State = Right(doc.getElementsByTagName("dd")(0).innerText, 2)

Range("B2") = county
Range("B3") = city
Range("B4") = State



MsgBox sDD


End Sub



Sub NewTransferTime()


Dim tot As Range
Dim ie As New InternetExplorer
Application.ScreenUpdating = False
endCol = Cells(1, 1).End(xlToRight).Column
endRow = Cells(1, 1).End(xlDown).Row

'Set tot = Range(Cells(1, 1), Cells(endrow, endCol))
'tot.Replace What:=" ", Replacement:="%", LookAt:=xlPart, _
 '    SearchOrder:=xlByRows, MatchCase:=False



ie.Visible = True

For i = 2 To endCol

origin = Cells(1, i)



For x = 2 To endRow

    destination = Cells(x, 1)
Do
ie.navigate "https://maps.googleapis.com/maps/api/directions/xml?origin=" & origin & ",&destination=" & destination & "&alternatives=false?key=AIzaSyA6LSCQbtuOfDnDH0NCbH_WaYDjCcNQ0cU"
Loop Until ie.readyState = READYSTATE_COMPLETE

Dim doc As
Set doc = ie.document

Dim transfer As String
transfer = ("//leg/duration/value")
'("DirectionResponse/route/leg/duration/text")(0).innerText
 
 MsgBox transfer
    
    
    Next x
    Next i

Application.Wait (Now + TimeValue("00:00:01"))

'Application.Wait (Now + 1 / (24 * 60 * 60))


tot.Replace What:="%", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False

Application.ScreenUpdating = True

End Sub




Sub othertest()

Dim request As New XMLHTTP30
Dim results As New DOMDocument30
Dim statusnode As IXMLDOMNode
Dim DistanceNode As IXMLDOMNode


If Range("A1") = "" Then
MsgBox ("Please paste the cities, with state if possible, in the leftmost column and the airport addresses in the top row, with the Service Area Name in the first cell.")
End
End If

On Error Resume Next

'Application.ScreenUpdating = False

endCol = Cells(1, 1).End(xlToRight).Column
endRow = Cells(1, 1).End(xlDown).Row

'check each column

For i = 2 To endCol

origin = Cells(1, i)



    For x = 2 To endRow

    destination = Cells(x, 1)

origin = WorksheetFunction.Substitute(origin, " ", "+")
destination = WorksheetFunction.Substitute(destination, " ", "+")

request.Open "GET", "https://maps.googleapis.com/maps/api/directions/xml?origin=" & origin & ",&destination=" & destination & "&alternatives=false?key=AIzaSyA6LSCQbtuOfDnDH0NCbH_WaYDjCcNQ0cU"
Application.Wait (Now + TimeValue("00:00:01"))
request.send
Application.Wait (Now + TimeValue("00:00:01"))
results.LoadXML request.responseText
Application.Wait (Now + TimeValue("00:00:01"))
'Shell ("C:\Program Files (x86)\Google\Chrome\Application\Chrome.exe -url https://maps.googleapis.com/maps/api/directions/xml?origin=" & origin & ",&destination=" & destination & "&alternatives=false?key=AIzaSyA6LSCQbtuOfDnDH0NCbH_WaYDjCcNQ0cU")
Set statusnode = results.SelectSingleNode("//status")

'Based on the status node result, proceed accordingly.
    Select Case statusnode.Text
            
        Case "OK"   'The response contains a valid result.

            Set DurationNode = results.SelectSingleNode("//leg/duration/value")
            Cells(x, i) = DurationNode.Text
            Cells(x, i) = Round(Cells(x, i) / 60)
    
        Case "INVALID_REQUEST"  'The provided request was invalid.
                                'Common causes of this status include an invalid parameter or parameter value.
            Cells(x, i) = "Invalid request"
        
        Case "NOT_FOUND"    'At least one of the locations specified in the requests's origin,
                            'destination, or waypoints could not be geocoded.
            Cells(x, i) = "Origin/destination could not be geocoded"
                    
        Case "ZERO_RESULTS" 'No route could be found between the origin and destination.
            Cells(x, i) = "Could not find route"
                                
        Case "MAX_WAYPOINTS_EXCEEDED"   'Too many waypoints were provided in the request The maximum allowed waypoints is 8, plus the origin, and destination.
                                        '(Google Maps API for Business customers may contain requests with up to 23 waypoints.)
            Cells(x, i) = "Too many waypoints"
            
        Case "OVER_QUERY_LIMIT" 'The requestor has exceeded limit.
            Cells(x, i) = "Requestor has exceeded limit"
            
        Case "REQUEST_DENIED"   'The service denied use of the directions service.
            Cells(x, i) = "Invalid sensor parameter"
        
        Case "UNKNOWN_ERROR"    'The request could not be processed due to a server error.
            Cells(x, i) = "Server error"
        
        Case Else   'Just in case...
            Cells(x, i) = "Error"
        
    End Select

    Set statusnode = Nothing
    Set DurationNode = Nothing
    Set results = Nothing
    Set request = Nothing

Application.Wait (Now + TimeValue("00:00:00"))


Next x

Next i






End Sub


