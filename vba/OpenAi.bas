Option Explicit

Function ChatGpt(prompt As String, model As String, effort As String, apiKey As String) As String
    Dim http As Object
    Dim jsonRequest As String
    Dim jsonResponse As String
    Dim data As JsonData, item As JsonData

    On Error GoTo ErrHandler

    ' Create HTTP request
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Build JSON request body
    jsonRequest = "{""model"":""" & model & """,""reasoning"":{""effort"":""" & effort & """},""input"":""" & Replace(prompt, """", "\""") & """}"

    ' Send request
    http.Open "POST", "https://api.openai.com/v1/responses", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.Send jsonRequest

    ' Read the response
    jsonResponse = http.responseText
    
    ' Parse json
    Set data = ParseJSON(jsonResponse)
    
    ' Throw error if json is not valid
    If Not data.IsValid Then
        ChatGpt = "Error: invalid JSON: " & jsonResponse
        Exit Function
    End If
    
    ' Extract text response
    Set item = data.GetChildByPath("output.1.content.0.text")
    
    ' Return response if it exists, otherwise throw error
    If item.IsScalar Then
        ChatGpt = item.ScalarValue
        Exit Function
    Else
        ChatGpt = "Error: unable to find text field in JSON: " & jsonResponse
        Exit Function
    End If

ErrHandler:
    ChatGpt = "Error: " & Err.Description
End Function


