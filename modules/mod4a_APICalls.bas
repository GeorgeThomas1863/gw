*****************************************

MOD 4a_APICalls

*****************************************

Option Explicit

Private Declare PtrSafe Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
    (ByVal lpszAgent As String, ByVal dwAccessType As Long, _
    ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, _
    ByVal dwFlags As Long) As LongPtr

Private Declare PtrSafe Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
    (ByVal hInternet As LongPtr, ByVal lpszUrl As String, _
    ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, _
    ByVal dwFlags As Long, ByVal dwContext As Long) As LongPtr

Private Declare PtrSafe Function InternetReadFile Lib "wininet.dll" _
    (ByVal hFile As LongPtr, ByVal lpBuffer As String, _
    ByVal dwNumberOfBytesToRead As Long, lNumberOfBytesRead As Long) As Long

Private Declare PtrSafe Function InternetCloseHandle Lib "wininet.dll" _
    (ByVal hInet As LongPtr) As Long

Private Declare PtrSafe Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
    (ByVal hConnect As LongPtr, ByVal lpszVerb As String, _
    ByVal lpszObjectName As String, ByVal lpszVersion As String, _
    ByVal lpszReferer As String, ByVal lplpszAcceptTypes As LongPtr, _
    ByVal dwFlags As Long, ByVal dwContext As Long) As LongPtr

Private Declare PtrSafe Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" _
    (ByVal hRequest As LongPtr, ByVal lpszHeaders As String, _
    ByVal dwHeadersLength As Long, ByVal lpOptional As LongPtr, _
    ByVal dwOptionalLength As Long) As Long

Private Declare PtrSafe Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
    (ByVal hInternet As LongPtr, ByVal lpszServerName As String, _
    ByVal nServerPort As Integer, ByVal lpszUsername As String, _
    ByVal lpszPassword As String, ByVal dwService As Long, _
    ByVal dwFlags As Long, ByVal dwContext As Long) As LongPtr

' Constants
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_SERVICE_HTTP = 3
Private Const INTERNET_FLAG_SECURE = &H800000
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const INTERNET_DEFAULT_HTTPS_PORT = 443
Private Const INTERNET_DEFAULT_HTTP_PORT = 80

'complex API call works after login
Function WinINetReq(strURL As String, token As String, Optional method As String = "GET", Optional params As String = "") As String
    Dim hOpen As LongPtr, hConnect As LongPtr, hRequest As LongPtr
    Dim buffer As String, res As String, headers As String, serverName As String, urlPath As String
    Dim bytesRead As Long, pos As Long, port As Integer, dwFlags As Long, sendRes As Long
    Dim useHttps As Boolean, postDataBytes() As Byte
    
    ' Parse URL to get server and path
    If InStr(1, strURL, "https://", vbTextCompare) > 0 Then
        useHttps = True
        serverName = Mid(strURL, 9)
    ElseIf InStr(1, strURL, "http://", vbTextCompare) > 0 Then
        useHttps = False
        serverName = Mid(strURL, 8)
    Else
        WinINetReq = "FAIL: Invalid URL scheme"
        Exit Function
    End If
    
    ' Extract server name and path
    pos = InStr(serverName, "/")
    If pos > 0 Then
        urlPath = Mid(serverName, pos)
        serverName = Left(serverName, pos - 1)
    Else
        urlPath = "/"
    End If
    
    ' Set port
    If useHttps Then
        port = INTERNET_DEFAULT_HTTPS_PORT
    Else
        port = INTERNET_DEFAULT_HTTP_PORT
    End If
    
    ' Build headers
    headers = "Authorization: Bearer " & token & vbCrLf
    If method = "POST" And Len(params) > 0 Then
        headers = headers & "Content-Type: application/json" & vbCrLf
    End If
    
    ' Initialize Internet
    hOpen = InternetOpen("VBA WinINet", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hOpen = 0 Then
        WinINetReq = "FAIL: InternetOpen failed"
        Exit Function
    End If
    
    ' Connect to server
    hConnect = InternetConnect(hOpen, serverName, port, vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    If hConnect = 0 Then
        InternetCloseHandle hOpen
        WinINetReq = "FAIL: InternetConnect failed"
        Exit Function
    End If
    
    dwFlags = INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE Or INTERNET_FLAG_KEEP_CONNECTION
    If useHttps Then
        dwFlags = dwFlags Or INTERNET_FLAG_SECURE
    End If
    
    hRequest = HttpOpenRequest(hConnect, method, urlPath, "HTTP/1.1", vbNullString, 0, dwFlags, 0)
    If hRequest = 0 Then
        InternetCloseHandle hConnect
        InternetCloseHandle hOpen
        WinINetReq = "FAIL: HttpOpenRequest failed"
        Exit Function
    End If
    
    Debug.Print "RUNNING WININET API REQ"
    Debug.Print "INPUT PARAMS: " & params
    Debug.Print "HEADERS: " & headers
    
    ' Send request
    If Len(params) > 0 Then
        postDataBytes = StrConv(params, vbFromUnicode)
        sendRes = HttpSendRequest(hRequest, headers, Len(headers), VarPtr(postDataBytes(0)), UBound(postDataBytes) + 1)
    Else
        sendRes = HttpSendRequest(hRequest, headers, Len(headers), 0, 0)
    End If
    
    If sendRes = 0 Then
        InternetCloseHandle hRequest
        InternetCloseHandle hConnect
        InternetCloseHandle hOpen
        WinINetReq = "FAIL: HttpSendRequest failed"
        Exit Function
    End If
    
    ' Read response
    res = ""
    Do
        buffer = Space(8192)
        If InternetReadFile(hRequest, buffer, Len(buffer), bytesRead) = 0 Then
            Exit Do
        End If
        If bytesRead = 0 Then Exit Do
        res = res & Left(buffer, bytesRead)
    Loop
    
    ' Cleanup
    InternetCloseHandle hRequest
    InternetCloseHandle hConnect
    InternetCloseHandle hOpen
    
    WinINetReq = res
End Function

'XML GET
Function GetXML(strURL As String, token As String) As String
    Dim xmlHTTP As New MSXML2.XMLHTTP60
       
    'SEND GET req
    xmlHTTP.Open "GET", strURL, False
    xmlHTTP.setRequestHeader "Content-Type", "application/json"
    xmlHTTP.setRequestHeader "Authorization", "Bearer " + token
    xmlHTTP.send
    
    GetXML = xmlHTTP.responseText
End Function

'XML POST
Function PostXML(strURL As String, params As String, token As String) As String
    Dim xmlHTTP As New MSXML2.XMLHTTP60
    
    'SEND POST req
    xmlHTTP.Open "POST", strURL, False
    xmlHTTP.setRequestHeader "Content-Type", "application/json"
    xmlHTTP.setRequestHeader "Authorization", "Bearer " + token
    xmlHTTP.send params
    
    PostXML = xmlHTTP.responseText
End Function