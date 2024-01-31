Attribute VB_Name = "Modul1"
Sub test()

End Sub
Function EscapeJsonString(inputStr As String) As String
    Dim outputStr As String
    Dim char As String
    Dim i As Integer

    outputStr = ""
    
    For i = 1 To Len(inputStr)
        char = Mid(inputStr, i, 1)
        Select Case char
            Case """"
                outputStr = outputStr & "\""" ' Replace double quote with escaped double quote
            Case "\"
                outputStr = outputStr & "\\" ' Replace backslash with escaped backslash
            Case "/", "/"
                outputStr = outputStr & "\/" ' Replace forward slash with escaped forward slash (optional)
            Case Chr(8)
                outputStr = outputStr & "\b" ' Replace backspace with escaped character (optional)
            Case Chr(12)
                outputStr = outputStr & "\f" ' Replace form feed with escaped character (optional)
            Case vbCr
                outputStr = outputStr & "\r" ' Replace carriage return with escaped character (optional)
            Case vbLf
                outputStr = outputStr & "\n" ' Replace line feed with escaped character (optional)
            Case vbTab
                outputStr = outputStr & "\t" ' Replace tab with escaped character (optional)
            Case Else
                outputStr = outputStr & char ' Keep other characters as they are
        End Select
    Next i

    EscapeJsonString = outputStr
End Function




Function GPT(prompt As String, Optional data As String)

Dim baseURL, name, name1 As String
baseURL = "http://192.168.178.39:8000"
name = "/v1/gpt-prompt"
name1 = "/"

Dim status, response, req As String
prompt = EscapeJsonString(prompt)
data = EscapeJsonString(data)
req = "{""prompt"": """ & prompt & """"
If Not IsMissing(data) Then
    req = req & ", ""data"": """ & data & """"
End If
req = req & "}"
GPT = "LÄDT..."
With CreateObject("MSXML2.XMLHTTP")
    .Open "POST", baseURL & name, False
    .send req
    status = .status
    response = .responseText
    response = Right(response, Len(response) - 1)
    response = Left(response, Len(response) - 1)
    GPT = response
End With
End Function
