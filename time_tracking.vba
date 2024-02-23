Sub AutomatedTimeTrackingWithAPI()
    Dim ws As Worksheet
    Dim currentDate As Date
    Dim employeeName As String
    Dim currentColumn As Integer
    Dim lastRow As Integer
    Dim apiUrl As String
    Dim httpRequest As Object
    Dim jsonResponse As String
    
    Set ws = ThisWorkbook.Sheets("TimeSheet")
    
    apiUrl = "https://your-time-tracking-api.com/data"
    
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", apiUrl, False
    httpRequest.send
    
    jsonResponse = "{""employee"": ""John"", ""entry"": ""08:00"", ""exit"": ""17:00""}"
    
    employeeName = JsonParser(jsonResponse, "employee")
    
    currentDate = Date
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If WorksheetFunction.CountIf(ws.Range("A1:A" & lastRow), employeeName) = 0 Then
        ws.Cells(lastRow + 1, 1).Value = employeeName
    End If
    
    currentColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    
    If WorksheetFunction.CountIf(ws.Range("1:1"), currentDate) = 0 Then
        ws.Cells(1, currentColumn).Value = currentDate
    End If
    
    ws.Cells(WorksheetFunction.Match(employeeName, ws.Range("A:A"), 0), currentColumn).Value = Format(Now, "hh:mm:ss")
    
    MsgBox "Time tracking record for " & employeeName & " on " & Format(currentDate, "dd/mm/yyyy") & " completed successfully!"
End Sub

Function JsonParser(jsonString As String, key As String) As String
    
    Dim startIdx As Long
    Dim endIdx As Long
    
    startIdx = InStr(jsonString, """" & key & """") + Len(key) + 3
    endIdx = InStr(startIdx, jsonString, """") - startIdx
    
    If startIdx > 0 And endIdx > 0 Then
        JsonParser = Mid(jsonString, startIdx, endIdx)
    Else
        JsonParser = "Not found"
    End If
End Function
