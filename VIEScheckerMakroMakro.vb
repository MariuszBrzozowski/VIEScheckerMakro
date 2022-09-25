Sub ExecuteGetRequest()
Dim apiURL, requestString, ticker, endpoint, reqType, params As String
Dim request As MSXML2.ServerXMLHTTP60
JsonConverter.JsonOptions.EscapeSolidus = True

'count not empty cells in the column with requests
totalRequests = Application.CountA(Range("E:E"))

'loop starts from 2 to exclude column name and go with requests only
For i = 2 To totalRequests
    'get reqyest string
    requestString = Range("E" & i).Value
   
    Dim tbl As ListObject
    Dim rng As Range
    Set tbl = ActiveSheet.ListObjects("Table1")
    Set rng = tbl.ListColumns(2).DataBodyRange
   
    apiURL = requestString
   
    Set request = New ServerXMLHTTP60
    request.Open "GET", apiURL, False
    request.Send
   
    'save raw json response
    Range("F" & i).Value = request.ResponseText
   
    'parse json response
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(request.ResponseText)
    'save json response values in to proper columns
    Range("G" & i).Value = Json("isValid")
    Range("H" & i).Value = Json("requestDate")
    Range("I" & i).Value = Json("requestIdentifier")
    Range("J" & i).Value = Json("userError")
   
    request.Abort
Next i

End Sub