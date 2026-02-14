*****************************************

MOD 2b_S

*****************************************

Option Compare Database

Option Explicit

Function RunSBulkSearch(str As String, token As String) As String
    Dim json As Object
    Dim strURL As String, params As String, res As String, resBulk As String
    Dim searchCheck As String, searchCheckText As String, searchStr As String, returnStr As String
    Dim loopNum As Long, r As Long, i As Long

    searchStr = Trim(str)
    
    'null input
    If searchStr = "" Then RunSBulkSearch = "": Exit Function
    
    res = SearchSBatch(searchStr, token, 0)
    'If Trim(res) = "" Then ThrowError 1970, "BLANK SEARCH RETURN; INPUT STR: " & str: Exit Function
    If Trim(res) = "" Then RunSBulkSearch = "": Exit Function
    If InStr(UCase(Left(res, 50)), "FORBIDDEN") > 0 Then ThrowError 1968, "FORBIDDEN SEARCH RETURN; INPUT STR: " & str & " | INPUT TOKEN: " & token: Exit Function
    
    Set json = ParseJson(res)
    If json Is Nothing Then Exit Function
    
    r = Val(json("numFound"))
    Debug.Print "SEARCH: " & searchStr & " | NUMBER OF HITS: " & r
    
    If r < 500 Then RunSBulkSearch = res: Exit Function
    
    'if stupid number of returns throw error
    If r > 10000 Then RunSBulkSearch = "[TOO MANY TO COUNT; SEARCHED STRING TOO COMMON]"

    'check if want to proceed
    searchCheckText = DefinePopupText(Trim(r), Trim(searchStr), "searchCheck")
    searchCheck = MsgBox(searchCheckText, vbYesNo + vbDefaultButton1, "KEEP GOING??")
    
    Debug.Print "SEARCH CHECK: " & searchCheck
    
    If searchCheck = vbNo Or searchCheck = 7 Then RunSBulkSearch = "[SKIPPED]": Exit Function

    'floor round
    loopNum = Int(r / 500)
    
    If loopNum < 1 Then RunSBulkSearch = res: Exit Function
    
    returnStr = res & "!;"
    For i = 1 To loopNum
        resBulk = SearchSBatch(searchStr, token, i)
        
        If Trim(resBulk) <> "" Then
            returnStr = returnStr & resBulk & "!;"
        End If
    Next
    
    'remove trailing delim
    returnStr = Trim(Left(returnStr, Len(returnStr) - 2))
    
    'removes things, combines item array correctly
    RunSBulkSearch = FixSBulkReturn(returnStr)
End Function

Function SearchSBatch(str As String, token As String, startNum As Long) As String
    Dim searchStr As String, strURL As String, params As String
    Dim i As Long
    
    searchStr = Trim(str)
    i = 500 * startNum
    
    If searchStr = "" Then SearchSBatch = "": Exit Function
    
    strURL = "https://S-api.Fnet.F/services/search/api/search/v1"
    
    'PROPERLY ESCAPED
    params = "{""q"": ""\" & Chr(34) & searchStr & "\" & Chr(34) & """,""limit"": 500, ""start"": " & i & "}"
    
    SearchSBatch = WinINetReq(strURL, token, "POST", params)
End Function

'++++++++++++++++++++++++++++

Function ParseSSearch(str As String, searchStr As String) As String
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim json As Object, itemObj As Object
    Dim resStr As String, dateStr As String, uniqueIdStr As String, titleStr As String
    Dim i As Long
    
    'search result
    resStr = Trim(str)
    
    'If resStr = "" Then ThrowError 1970, "PARSE RETURN EMPTY; Parse Input: " & str & " | Search STR: " & searchStr: Exit Function
    If resStr = "" Then ParseSSearch = "": Exit Function
    
    'skip
    If InStr(LCase(resStr), "[skip") > 0 Then ParseSSearch = "": Exit Function
    
    'ExportJsonToFile resStr, "test1.json"
    
    Set json = ParseJson(resStr)
    'If json.Count = 0 Then ThrowError 1970, "PARSE RETURN EMPTY; Parse Input: " & str & " | Search STR: " & searchStr: Exit Function
    If json.Count = 0 Then ParseSSearch = "": Exit Function
    
    'could put in other function but unnecessary
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tempSSearchResults", dbOpenDynaset)
    
    For i = 1 To json("items").Count
        Set itemObj = json("items")(i)
        
        'unique id for each EC for highlighting tracking
        'might want to change to only display 1 hit per serial
        uniqueIdStr = Trim("U" & itemObj("uniqueID") & "_" & Trim(i))
        titleStr = Trim(itemObj("title"))
        If Len(titleStr) > 255 Then titleStr = Trim(Left(titleStr, 254))
        
        rs.AddNew
        rs!SId = uniqueIdStr 'only for highlighting / tracking
        rs!selector = Trim(searchStr)
        rs!docId = itemObj("uniqueID")
        rs!docType = itemObj("recordType")
        rs!docSubType = itemObj("recordSubType")
        rs!case = itemObj("UCFN")
        rs!serial = itemObj("itemNumber")
        rs!caseSerialFull = itemObj("UCFN") & " Serial " & itemObj("itemNumber")
        rs!office = itemObj("caseOfficeCode")
        rs!docTitle = titleStr
        rs!author = itemObj("primaryAuthor")
        rs!createdDate = FixDate(itemObj("createdDate"))
        rs!link = "https://S.Fnet.F/apps/desktop/#/main/serial/" & itemObj("uniqueID")
        
        rs.Update
    Next

    'return number of hits
    ParseSSearch = json("items").Count
End Function

'open S link
Function OpenLink(strURL As String) As String
    Dim chromePath As String
    
    chromePath = """C:\Program Files\Google\Chrome\Application\chrome.exe"""
    OpenLink = shell(chromePath & " """ & strURL & """", vbNormalFocus)
    
End Function

'UTIL for testing
Sub ExportJsonToFile(str As String, fileName As String)
    Dim fso As Object, file As Object, readFile As Object
    Dim outputPath As String, readStr As String
    Dim chunkSize As Long, i As Long
    
    'SET OUTPUT PATH
    outputPath = "C:\Users\RREMEDIO\OneDrive - F\All\TOOLS\GrayWolfe\" & fileName

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(outputPath) Then
        Set readFile = fso.OpenTextFile(outputPath, 1, False)
        readStr = Trim(readFile.ReadAll)
        
        If Right(Trim(readStr), 1) = "]" Then
            readStr = Trim(Left(readStr, Len(readStr) - 1))
        End If
        
        'overwrite
        Set file = fso.CreateTextFile(outputPath, True)
        file.Write readStr
        file.Write ","
        
    Else
        Set file = fso.CreateTextFile(outputPath, True)
        file.Write "["
    End If
    
    'write to file loop
    chunkSize = 100
    For i = 1 To Len(str) Step chunkSize
        On Error Resume Next
        file.Write Trim(Mid(str, i, chunkSize))
    Next
    
    file.Write "]"
    file.Close
End Sub
