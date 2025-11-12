Attribute VB_Name = "OpenAi"
Option Explicit

' Author: Wouter Grimme <coding@woutergrimme.nl>
' Source: https://github.com/droppiq/chatgpt-excel
' License: GNU AGPLv3

Function CHATGPT(prompt As String, model As String, effort As String, apiKey As String) As String
    Dim http As Object
    Dim jsonRequest As String
    Dim jsonResponse As String
    Dim data As JsonData, item As JsonData

    On Error GoTo ErrHandler

    ' Build JSON request
    jsonRequest = _
        "{" & _
            """model"": """ & model & """," & _
            """reasoning"": {" & _
                """effort"": """ & effort & """" & _
            "}," & _
            """input"": """ & Replace(prompt, """", "\""") & """" & _
        "}"

    ' Send HTTP request
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", "https://api.openai.com/v1/responses", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.Send jsonRequest

    ' Check http status
    If http.Status <> 200 Then
        CHATGPT = "HTTP Error " & http.Status & ": " & http.StatusText & vbCrLf & http.responseText
        Exit Function
    End If

    ' Get the response
    jsonResponse = http.responseText
    
    ' Parse JSON response
    Set data = ParseJSON(jsonResponse)
    
    ' Throw error if json is not valid
    If Not data.IsValid Then
        CHATGPT = "Error: invalid JSON: " & jsonResponse
        Exit Function
    End If
    
    ' Extract text response
    Set item = data.GetChildByPath("output.1.content.0.text")
    
    ' Return response if it exists, otherwise throw error
    If item.IsScalar Then
        CHATGPT = item.ScalarValue
        Exit Function
    Else
        CHATGPT = "Error: unable to find text field in JSON: " & jsonResponse
        Exit Function
    End If

ErrHandler:
    CHATGPT = "Error: " & Err.Description
End Function
