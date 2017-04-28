Attribute VB_Name = "getduration"
Public Function getduration(Start As String, dest As String)
    Dim firstVal As String, secondVal As String, lastVal As String
    firstVal = "http://maps.googleapis.com/maps/api/distancematrix/json?origins="
    secondVal = "&destinations="
    lastVal = "&mode=car&language=en&sensor=false"
    Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
    URL = firstVal & Replace(Start, " ", "+") & secondVal & Replace(dest, " ", "+") & lastVal
    objHttp.Open "GET", URL, False
    objHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHttp.send ("")
    If InStr(objHttp.responseText, """duration"" : {") = 0 Then GoTo ErrorHandl
    Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = "duration(?:.|\n)*?""value"".*?([0-9]+)": regex.Global = False
    Set matches = regex.Execute(objHttp.responseText)
    tmpVal = Replace(matches(0).SubMatches(0), ".", Application.International(xlListSeparator))
    getduration = CDbl(tmpVal)
    Exit Function
ErrorHandl:
    getduration = -1
End Function
