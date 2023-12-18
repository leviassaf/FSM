Attribute VB_Name = "modSisenseApi"
Option Explicit

' https://sisense.dev/guides/restApi/using-rest-api.html

Sub GetDataFromSisense()
    ' Set your Sisense API endpoint and credentials
    Dim sisenseEndpoint As String
    Dim username As String
    Dim password As String
    
    sisenseEndpoint = "https://sisense.trusteer.il.ibm.com/app/account/login"
    username = "lassaff"
    password = "s^hsHull#F$8#Dhf)S"
    
    ' Create and configure the XMLHttpRequest object
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Build the API request URL
    Dim apiURL As String
    apiURL = sisenseEndpoint
    
    ' Open the connection to Sisense API
    xhr.Open "GET", apiURL, False
    xhr.setRequestHeader "Content-Type", "application/json"
    xhr.setRequestHeader "Authorization", "Basic " & Base64Encode(username & ":" & password)
    
    ' Send the request
    xhr.send
    
    ' Check if the request was successful (status code 200)
    If xhr.Status = 200 Then
        ' Print the response to the immediate window (you can modify this part to handle the data as needed)
        SaveTextFile (xhr.responseText)
    Else
        ' Print an error message if the request was not successful
        Debug.Print "Error: " & xhr.Status & " - " & xhr.statusText
    End If
End Sub

Function Base64Encode(sText As String) As String
    Dim arrData() As Byte
    arrData = StrConv(sText, vbFromUnicode)
    
    Dim objXML As Object
    Set objXML = CreateObject("MSXML2.DOMDocument")
    
    Dim objNode As Object
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    Base64Encode = objNode.Text
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Sub SaveTextFile(Text As String)
    Dim fso As Object
    Dim oFile As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.CreateTextFile("C:\temp\Sisense.txt")
    oFile.WriteLine Text
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
End Sub
