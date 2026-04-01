Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Maximum age in seconds before a heartbeat is considered stale
Private Const HEARTBEAT_MAX_AGE      As Double = 5#
' Seconds to wait for the kernel to start before giving up
Private Const LAUNCH_WAIT_SEC        As Long   = 15
' Milliseconds between polling checks
Private Const POLL_MS                As Long   = 300
' Default HTTP timeout in seconds passed to Python requests
Private Const DEFAULT_TIMEOUT        As Long   = 30
' Network path to the Anaconda installer — update this before use
Private Const ANACONDA_INSTALLER_SRC As String = "\\YOUR_NETWORK_PATH\Anaconda3-installer.exe"
' Port the Jupyter notebook server listens on
Private Const JUPYTER_PORT           As Long   = 8888

' ============================================================
' Path helpers
' ============================================================

' Returns the shared temp directory used for bridge I/O
Private Function BridgeDir() As String
    BridgeDir = Environ("TEMP") & "\JupyterBridge"
End Function

' Returns the path to the command file VBA writes for Python to read
Private Function CmdFilePath() As String
    CmdFilePath = BridgeDir() & "\cmd.json"
End Function

' Returns the path to the response file Python writes for VBA to read
Private Function RespFilePath() As String
    RespFilePath = BridgeDir() & "\resp.txt"
End Function

' Returns the path to the heartbeat file Python updates each poll cycle
Private Function HeartbeatPath() As String
    HeartbeatPath = BridgeDir() & "\heartbeat.txt"
End Function

' Returns the Anaconda installation directory for this user.
' Falls back to C:\anaconda3 if the username contains spaces, since the
' Anaconda (NSIS) silent installer /D= flag cannot handle paths with spaces.
Private Function AnacondaDir() As String
    Dim strUser As String
    strUser = Environ("USERNAME")
    If InStr(strUser, " ") > 0 Then
        AnacondaDir = "C:\anaconda3"
    Else
        AnacondaDir = "C:\Users\" & strUser & "\anaconda3"
    End If
End Function

' Returns the expected path to the Anaconda jupyter.exe for this user
Private Function JupyterExePath() As String
    JupyterExePath = AnacondaDir() & "\Scripts\jupyter.exe"
End Function

' Returns the base URL for the local Jupyter notebook server API
Private Function JupyterApiBase() As String
    JupyterApiBase = "http://localhost:" & JUPYTER_PORT
End Function

' Returns the path where the IPython startup script should be installed
Private Function StartupScriptPath() As String
    StartupScriptPath = Environ("USERPROFILE") & "\.ipython\profile_default\startup\00_bridge.py"
End Function

' ============================================================
' FSO helper
' ============================================================

' Returns a fresh FileSystemObject instance
Private Function GetFSO() As Object
    Set GetFSO = CreateObject("Scripting.FileSystemObject")
End Function

' ============================================================
' HTTP helpers (localhost only — no Windows auth required)
' ============================================================

' Issues a synchronous GET and returns the HTTP status code, or 0 on connection error
Private Function HttpGetStatus(sUrl As String) As Long
    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    On Error GoTo 0
    If http Is Nothing Then Exit Function  ' MSXML not available; return 0 = not up
    On Error Resume Next
    http.Open "GET", sUrl, False
    http.Send
    HttpGetStatus = http.Status
    On Error GoTo 0
End Function

' Issues a synchronous POST with a JSON body; return value is ignored
Private Sub HttpPost(sUrl As String, sBody As String)
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "POST", sUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send sBody
End Sub

' ============================================================
' Utility helpers
' ============================================================

' Returns the age in seconds of the file at sPath based on its last-modified timestamp
Private Function FileAgeSec(sPath As String) As Double
    Dim fso As Object
    Set fso = GetFSO()
    FileAgeSec = (Now - fso.GetFile(sPath).DateLastModified) * 86400#
End Function

' Returns True if the Python bridge is running and has written a recent heartbeat
Private Function IsBridgeAlive() As Boolean
    Dim fso As Object
    Set fso = GetFSO()
    If Not fso.FileExists(HeartbeatPath()) Then Exit Function
    If FileAgeSec(HeartbeatPath()) >= HEARTBEAT_MAX_AGE Then Exit Function
    IsBridgeAlive = True
End Function

' Escapes a string value for safe embedding inside a JSON double-quoted string
Private Function EscapeJsonStr(sVal As String) As String
    Dim s As String
    s = sVal
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr(13), "\r")
    s = Replace(s, Chr(10), "\n")
    s = Replace(s, Chr(9), "\t")
    EscapeJsonStr = s
End Function

' ============================================================
' Public JSON builder
' ============================================================

' Builds a flat JSON object string from alternating key/value pairs (all values treated as strings)
Public Function BJ(ParamArray kv() As Variant) As String
    Dim i       As Long
    Dim strOut  As String

    If ((UBound(kv) - LBound(kv) + 1) Mod 2) <> 0 Then
        Err.Raise vbObjectError + 999, "BJ", "BJ() requires an even number of arguments (key/value pairs)"
    End If

    strOut = "{"
    For i = LBound(kv) To UBound(kv) Step 2
        If i > LBound(kv) Then strOut = strOut & ","
        strOut = strOut & """" & EscapeJsonStr(CStr(kv(i))) & """:" & _
                          """" & EscapeJsonStr(CStr(kv(i + 1))) & """"
    Next i
    strOut = strOut & "}"
    BJ = strOut
End Function

' ============================================================
' Startup script installer
' ============================================================

' Writes the Python polling bridge script to the IPython startup directory,
' creating any missing directory levels along the way
Private Sub WriteStartupScript()
    Dim fso     As Object
    Dim ts      As Object
    Dim strPath As String
    Dim strDir  As String
    Dim strCode As String

    Set fso = GetFSO()

    ' Ensure IPython startup directory exists, creating each level as needed
    strDir = Environ("USERPROFILE") & "\.ipython"
    If Not fso.FolderExists(strDir) Then fso.CreateFolder strDir

    strDir = strDir & "\profile_default"
    If Not fso.FolderExists(strDir) Then fso.CreateFolder strDir

    strDir = strDir & "\startup"
    If Not fso.FolderExists(strDir) Then fso.CreateFolder strDir

    ' Build the Python source string (exactly matches startup_bridge.py)
    strCode = "import os" & vbLf & _
              "import json" & vbLf & _
              "import time" & vbLf & _
              "import threading" & vbLf & _
              "import requests" & vbLf & _
              "" & vbLf & _
              "BRIDGE_DIR = os.path.join(os.environ[""TEMP""], ""JupyterBridge"")" & vbLf & _
              "CMD_FILE   = os.path.join(BRIDGE_DIR, ""cmd.json"")" & vbLf & _
              "RESP_FILE  = os.path.join(BRIDGE_DIR, ""resp.txt"")" & vbLf & _
              "HB_FILE    = os.path.join(BRIDGE_DIR, ""heartbeat.txt"")" & vbLf & _
              "" & vbLf & _
              "POLL_INTERVAL = 0.5  # seconds" & vbLf & _
              "" & vbLf & _
              "" & vbLf & _
              "def _execute_request(cmd):" & vbLf & _
              "    method  = cmd.get(""method"", ""GET"").upper()" & vbLf & _
              "    url     = cmd[""url""]" & vbLf & _
              "    headers = cmd.get(""headers"") or {}" & vbLf & _
              "    params  = cmd.get(""params"") or {}" & vbLf & _
              "    body    = cmd.get(""body"")" & vbLf & _
              "    timeout = cmd.get(""timeout"", 30)" & vbLf & _
              "" & vbLf & _
              "    if method == ""GET"":" & vbLf & _
              "        resp = requests.get(url, headers=headers, params=params, timeout=timeout)" & vbLf & _
              "    elif method == ""POST"":" & vbLf & _
              "        resp = requests.post(url, headers=headers, params=params, json=body, timeout=timeout)" & vbLf & _
              "    else:" & vbLf & _
              "        raise ValueError(""Unsupported method: "" + method)" & vbLf & _
              "" & vbLf & _
              "    return resp.status_code, resp.text" & vbLf & _
              "" & vbLf & _
              "" & vbLf & _
              "def _poll_loop():" & vbLf & _
              "    os.makedirs(BRIDGE_DIR, exist_ok=True)" & vbLf & _
              "" & vbLf & _
              "    while True:" & vbLf & _
              "        try:" & vbLf & _
              "            with open(HB_FILE, ""w"") as f:" & vbLf & _
              "                f.write(str(time.time()))" & vbLf & _
              "" & vbLf & _
              "            if os.path.exists(CMD_FILE):" & vbLf & _
              "                try:" & vbLf & _
              "                    with open(CMD_FILE, ""r"") as f:" & vbLf & _
              "                        cmd = json.load(f)" & vbLf & _
              "                    os.remove(CMD_FILE)" & vbLf & _
              "                    status, body = _execute_request(cmd)" & vbLf & _
              "                except Exception as e:" & vbLf & _
              "                    status, body = -1, str(e)" & vbLf & _
              "" & vbLf & _
              "                _tmp = RESP_FILE + "".tmp""" & vbLf & _
              "                with open(_tmp, ""w"") as f:" & vbLf & _
              "                    f.write(str(status) + ""\n"" + body)" & vbLf & _
              "                os.replace(_tmp, RESP_FILE)" & vbLf & _
              "" & vbLf & _
              "        except Exception:" & vbLf & _
              "            pass" & vbLf & _
              "" & vbLf & _
              "        time.sleep(POLL_INTERVAL)" & vbLf & _
              "" & vbLf & _
              "" & vbLf & _
              "_bridge_thread = threading.Thread(target=_poll_loop, daemon=True)" & vbLf & _
              "_bridge_thread.start()" & vbLf

    strPath = StartupScriptPath()
    Set ts = fso.CreateTextFile(strPath, True)
    ts.Write strCode
    ts.Close
End Sub

' ============================================================
' Kernel lifecycle
' ============================================================

' Launches the Jupyter notebook server in a hidden window with no browser and no token auth
Private Sub LaunchJupyterServer()
    Shell """" & JupyterExePath() & """ notebook --no-browser --port=" & JUPYTER_PORT & _
          " --NotebookApp.token= --NotebookApp.password=" & _
          " --ServerApp.token= --ServerApp.password=", vbHide
End Sub

' Returns True if the Jupyter notebook server is responding on the configured port
Private Function IsJupyterServerUp() As Boolean
    IsJupyterServerUp = (HttpGetStatus(JupyterApiBase() & "/api") = 200)
End Function

' Polls IsJupyterServerUp for up to lMaxSec seconds; returns True if server comes up in time
Private Function WaitForServerUp(lMaxSec As Long) As Boolean
    If lMaxSec <= 0 Then Exit Function
    Dim i As Long
    For i = 1 To (lMaxSec * 1000) \ POLL_MS
        DoEvents
        If IsJupyterServerUp() Then
            WaitForServerUp = True
            Exit Function
        End If
        Sleep POLL_MS
    Next i
End Function

' Creates a new Jupyter kernel via the REST API, which triggers the IPython startup script
Private Sub CreateJupyterKernel()
    HttpPost JupyterApiBase() & "/api/kernels", "{}"
End Sub

' Polls IsBridgeAlive up to lMaxSec seconds; returns True if the bridge comes alive in time
Private Function WaitForHeartbeat(lMaxSec As Long) As Boolean
    If lMaxSec <= 0 Then Exit Function
    Dim i As Long
    For i = 1 To (lMaxSec * 1000) \ POLL_MS
        DoEvents
        If IsBridgeAlive() Then
            WaitForHeartbeat = True
            Exit Function
        End If
        Sleep POLL_MS
    Next i
End Function

' Ensures the Python bridge is running, installing Anaconda and launching Jupyter if needed
Private Sub EnsureBridgeReady()
    Dim fso As Object
    Dim wsh As Object

    ' Fast path: bridge is already running
    If IsBridgeAlive() Then Exit Sub

    ' Slow path: set everything up
    Set fso = GetFSO()

    ' 1. Install Anaconda if Jupyter is not present
    If Not fso.FileExists(JupyterExePath()) Then
        If Not fso.FileExists(ANACONDA_INSTALLER_SRC) Then
            Err.Raise vbObjectError + 1001, "EnsureBridgeReady", _
                "Anaconda is not installed and the installer was not found at: " & _
                ANACONDA_INSTALLER_SRC
        End If
        MsgBox "Anaconda is not installed. Installing now — this may take 5-15 minutes." & _
               vbCrLf & "Please do not close this application.", vbInformation, "Installing Anaconda"
        Set wsh = CreateObject("WScript.Shell")
        wsh.Run """" & ANACONDA_INSTALLER_SRC & """ /S /D=" & AnacondaDir(), 0, True
        If Not fso.FileExists(JupyterExePath()) Then
            Err.Raise vbObjectError + 1002, "EnsureBridgeReady", _
                "Anaconda installation completed but Jupyter was not found at: " & _
                JupyterExePath() & ". Verify the installer at: " & ANACONDA_INSTALLER_SRC
        End If
    End If

    ' 2. Write startup script if missing
    If Not fso.FileExists(StartupScriptPath()) Then
        WriteStartupScript
    End If

    ' 3. Ensure bridge directory exists
    If Not fso.FolderExists(BridgeDir()) Then
        fso.CreateFolder BridgeDir()
    End If

    ' 4. Launch Jupyter notebook server if not already running
    If Not IsJupyterServerUp() Then
        LaunchJupyterServer
        If Not WaitForServerUp(LAUNCH_WAIT_SEC) Then
            Err.Raise vbObjectError + 1003, "EnsureBridgeReady", _
                "Jupyter notebook server did not respond within " & LAUNCH_WAIT_SEC & _
                " seconds. Port " & JUPYTER_PORT & " may be in use by another process."
        End If
    End If

    ' 5. Create a kernel — triggers the IPython startup script and polling thread
    CreateJupyterKernel

    ' 6. Wait for the polling thread to prove itself via heartbeat
    If Not WaitForHeartbeat(LAUNCH_WAIT_SEC) Then
        Err.Raise vbObjectError + 1004, "EnsureBridgeReady", _
            "Jupyter kernel was created but the bridge did not start polling within " & _
            LAUNCH_WAIT_SEC & " seconds. Verify: (1) the startup script exists at " & _
            StartupScriptPath() & ", and (2) the ""requests"" package is installed " & _
            "in the active Anaconda environment."
    End If
End Sub

' ============================================================
' Command/response I/O
' ============================================================

' Writes the JSON command file for the Python bridge to consume, clearing any stale response first
Private Sub WriteCmdFile(sMethod As String, sUrl As String, _
    sHeadersJson As String, sParamsJson As String, sBodyJson As String)

    Dim fso     As Object
    Dim ts      As Object
    Dim strJson As String

    Set fso = GetFSO()

    ' Clear any stale response from a previous call
    If fso.FileExists(RespFilePath()) Then fso.DeleteFile RespFilePath()

    strJson = "{" & _
              """method"":" & """" & EscapeJsonStr(UCase(sMethod)) & """," & _
              """url"":" & """" & EscapeJsonStr(sUrl) & """," & _
              """headers"":" & sHeadersJson & "," & _
              """params"":" & sParamsJson & "," & _
              """body"":" & sBodyJson & "," & _
              """timeout"":" & CStr(DEFAULT_TIMEOUT) & _
              "}"

    Set ts = fso.CreateTextFile(CmdFilePath(), True)
    ts.Write strJson
    ts.Close
End Sub

' Polls for the response file up to lTimeoutSec seconds; returns the response body on success
Private Function WaitForResponse(lTimeoutSec As Long) As String
    Dim fso     As Object
    Dim ts      As Object
    Dim i       As Long
    Dim lStatus As Long
    Dim strBody As String

    Set fso = GetFSO()

    For i = 1 To (lTimeoutSec * 1000) \ POLL_MS
        DoEvents
        If fso.FileExists(RespFilePath()) Then
            ' Read response: line 1 = status code, remainder = body
            Set ts = fso.OpenTextFile(RespFilePath(), 1)  ' 1 = ForReading
            lStatus = CLng(ts.ReadLine())
            If ts.AtEndOfStream Then
                strBody = ""
            Else
                strBody = ts.ReadAll()
            End If
            ts.Close
            fso.DeleteFile RespFilePath()

            ' Evaluate result
            If lStatus < 0 Then
                Err.Raise vbObjectError + 1005, "BridgeCall", _
                    "Bridge error: " & strBody
            End If
            If lStatus >= 400 Then
                Err.Raise vbObjectError + 1006, "BridgeCall", _
                    "HTTP " & lStatus & ": " & strBody
            End If

            WaitForResponse = strBody
            Exit Function
        End If
        Sleep POLL_MS
    Next i

    Err.Raise vbObjectError + 1007, "BridgeCall", _
        "Timeout: no response received after " & lTimeoutSec & " seconds"
End Function

' ============================================================
' Public API
' ============================================================

' Makes an HTTP request via the Python bridge; returns the response body as a string
Public Function BridgeCall(sMethod As String, sUrl As String, _
    Optional sHeadersJson As String = "{}", _
    Optional sParamsJson  As String = "{}", _
    Optional sBodyJson    As String = "null") As String

    On Error GoTo ErrHandler

    EnsureBridgeReady
    WriteCmdFile sMethod, sUrl, sHeadersJson, sParamsJson, sBodyJson
    BridgeCall = WaitForResponse(DEFAULT_TIMEOUT)
    Exit Function

ErrHandler:
    Debug.Print "BridgeCall FAILED | Err " & Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
