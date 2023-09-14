' Configure the MSXML2.ServerXMLHTTP object
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
url = "https://api.telegram.org/bot5676063743:AAEyntU6QDEJd3mQMR1lii95nKPXlliX_tY/sendMessage" '(1) Replace <YourBotToken> with bot token

' Open a POST request
objHTTP.open "POST", url, False
objHTTP.setRequestHeader "Content-Type", "application/json"

' Call the input box to enter the message
message = InputBox("Please enter the message text:")

' Ensure that a message text is entered
If message <> "" Then
    ' Construct the data to be sent as a JSON body
    jsonPayload = "{""chat_id"": 930980136, ""text"": """ & message & """}" '(2) Replace chat ID message text

    ' Send the data
    objHTTP.send jsonPayload

    WScript.Echo "Sending message to URL: " & url
    WScript.Echo "JSON Payload: " & jsonPayload

    ' Check the response status
    If objHTTP.status = 200 Then
        WScript.Echo "Message sent successfully!"
        WScript.Echo "Response: " & objHTTP.responseText
    Else
        WScript.Echo "Error sending message. Status code: " & objHTTP.status
        WScript.Echo "Response: " & objHTTP.responseText
    End If
Else
    WScript.Echo "No message text entered."
End If

' Close the HTTP object
Set objHTTP = Nothing
