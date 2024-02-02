Sub AutomatedTimeTrackingWithAPI()
    Dim ws As Worksheet
    Dim currentDate As Date
    Dim employeeName As String
    Dim currentColumn As Integer
    Dim lastRow As Integer
    Dim apiUrl As String
    Dim httpRequest As Object
    Dim jsonResponse As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("TimeSheet")
    
    ' Replace the API URL below with your actual URL
    apiUrl = "https://your-time-tracking-api.com/data"
    
    ' Example of HTTP request using MSXML2.XMLHTTP
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", apiUrl, False
    httpRequest.send
    
    ' Example JSON response
    jsonResponse = "{""employee"": ""John"", ""entry"": ""08:00"", ""exit"": ""17:00""}"
    
    ' Parse JSON to retrieve data
    ' Example parsing:
    employeeName = JsonParser(jsonResponse, "employee")
    
    ' Get the current date
    currentDate = Date
    
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Check if the employee name is already in the list
    If WorksheetFunction.CountIf(ws.Range("A1:A" & lastRow), employeeName) = 0 Then
        ' Add the employee name to the next empty row in column A
        ws.Cells(lastRow + 1, 1).Value = employeeName
    End If
    
    ' Find the last column in row 1
    currentColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    
    ' Check if the column for the current date already exists, if not, create it
    If WorksheetFunction.CountIf(ws.Range("1:1"), currentDate) = 0 Then
        ' Add the current date to the next empty column in row 1
        ws.Cells(1, currentColumn).Value = currentDate
    End If
    
    ' Automatically fill in entry and exit times
    ws.Cells(WorksheetFunction.Match(employeeName, ws.Range("A:A"), 0), currentColumn).Value = Format(Now, "hh:mm:ss")
    
    ' Success message
    MsgBox "Time tracking record for " & employeeName & " on " & Format(currentDate, "dd/mm/yyyy") & " completed successfully!"
End Sub

Function JsonParser(jsonString As String, key As String) As String
    ' Fictitious function for simple JSON parsing
    ' Replace this with actual JSON parsing, if necessary
    
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
