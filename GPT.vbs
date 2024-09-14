' Initialize variables
Dim objHTTP, url, apiKey, data, response, userInput, defaultText, scriptPath
Dim objFSO, objFile

' Get the script file path
Set objFSO = CreateObject("Scripting.FileSystemObject")
scriptPath = objFSO.GetAbsolutePathName(WScript.ScriptFullName)

' Set the API endpoint and API key
url = "https://api.openai.com/v1/chat/completions"
apiKey = "YOUR_API_KEY"

' Initialize default text
defaultText = "Please enter your text:"

' Function to extract a value from a JSON string
Function ExtractJsonValue(json, key)
    Dim regex, matches, match
    Set regex = New RegExp
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = """" & key & """:\s*""([^""]*)"""
    
    Set matches = regex.Execute(json)
    If matches.Count > 0 Then
        ExtractJsonValue = matches(0).SubMatches(0)
    Else
        ExtractJsonValue = ""
    End If
End Function

' Loop until the user cancels the input box
Do
    ' Create an Input Box to get user input
    userInput = InputBox(defaultText, "ChatGPT VBS Interface", "")

    ' Check if the user clicked OK or entered some text
    If userInput = "" Then
        ' Exit loop if userInput is empty (Cancel was clicked)
        Exit Do
    Else
        ' Create the HTTP object
        Set objHTTP = CreateObject("MSXML2.XMLHTTP")

        ' Create the JSON request payload
        data = "{""model"": ""gpt-4"", ""messages"": [{""role"": ""user"", ""content"": """ & userInput & """}]}"

        ' Open an HTTP POST connection
        objHTTP.Open "POST", url, False

        ' Set the request headers
        objHTTP.setRequestHeader "Content-Type", "application/json"
        objHTTP.setRequestHeader "Authorization", "Bearer " & apiKey

        ' Send the request with the JSON payload
        objHTTP.send data

        ' Wait for the response
        Do While objHTTP.readyState <> 4
            WScript.Sleep 100
        Loop

        ' Get the response text
        response = objHTTP.responseText

        ' Extract the answer from the JSON response
        answer = ExtractJsonValue(response, "content")

        ' Append the question and answer to the defaultText for the next loop iteration
        defaultText = defaultText & vbCrLf & "Q: " & userInput & vbCrLf & "A: " & answer & vbCrLf & vbCrLf & "Please enter your text:"
    End If
Loop

' Clean up
Set objHTTP = Nothing
Set objFSO = Nothing
Set objFile = Nothing

' Display a message indicating the script has finished
MsgBox "Chat session ended."
