****************************

MOD 1a_RunSearch

****************************

Option Compare Database

Option Explicit

Sub RunSearch(str As String, searchTargetInput As String, delimInput As String)
    Dim searchArr() As String, searchStr As String, delimStr As String, searchTarget As String
    Dim resGW As String, resS As String, openArgsStr As String
    
    delimStr = CleanDelimInput(Trim(str), Trim(delimInput))
    
    'blank input error
    If Trim(str) = "" Or InStr(Trim(str), "Search one or multiple") = 1 Then ThrowError 1950, str: Exit Sub

    searchStr = CleanUserInput(Trim(str), delimStr)
    searchTarget = CleanTargetInput(Trim(searchTargetInput))
    resGW = "!!"
    resS = "!!$$" 'for split in display
    
    Debug.Print "++++++++++++"
    Debug.Print "CLEAN SEARCH INPUT: " & searchStr & vbLf & "-----------------"
    Debug.Print "CLEAN SEARCH TARGET: " & searchTarget & vbLf & "-----------------"
    
    Select Case searchTarget

    Case "graywolfe"
        resGW = SearchGrayWolfe(searchStr)
    
    Case "S"
        resS = SearchS(searchStr)
    
    Case "both"
        resGW = SearchGrayWolfe(searchStr)
        resS = SearchS(searchStr)
    
    End Select
    
    openArgsStr = searchStr & "$$" & resGW & "$$" & resS
    Debug.Print "^^^^^" & vbLf & "OPEN ARGS STR FROM SEARCH: " & openArgsStr

    'display results
    DoCmd.OpenForm "frmResultsDisplay", acNormal, , , , , openArgsStr
End Sub


Function SearchGrayWolfe(str As String) As String
    Dim db As DAO.Database, rs As DAO.Recordset, rsCount As DAO.Recordset
    Dim searchArr() As String
    Dim splitStr As String, searchStr As String, selectorCleanStr As String
    Dim strSQL As String, recordCount As Long
    Dim i As Long, x As Long, t As Long

    'replace all delims
    splitStr = Replace(Trim(str), "+++", "!!")
    searchArr = Split(splitStr, "!!")

    'shouldnt happen
    If UBound(searchArr) = -1 Then ThrowError 1961, str: Exit Function

    Set db = CurrentDb

    x = 0 'gw hit counter
    For i = LBound(searchArr) To UBound(searchArr)
        searchStr = Trim(searchArr(i))

        If searchStr <> "" And LCase(searchStr) <> "null" Then
            selectorCleanStr = BuildSelectorClean(searchStr)

            Set rs = SearchSelectorTbl(searchStr, selectorCleanStr)
            If Not rs.EOF Then
                rs.MoveFirst
                rs.MoveLast
                recordCount = rs.recordCount

                x = x + recordCount
            End If

            FillTempGWSearchResults rs, searchStr, selectorCleanStr
        End If
    Next

    'count distinct targets from search results (handles shared selectors across multiple targets)
    t = 0
    strSQL = "SELECT COUNT(*) AS cnt FROM (SELECT DISTINCT targetId FROM [tempGWSearchResults] WHERE [targetId] IS NOT NULL AND [targetId] <> '') AS subQ"
    Set rsCount = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If Not rsCount.EOF Then t = Nz(rsCount!cnt, 0)
    rsCount.Close

    'target hits !! selector hits
    SearchGrayWolfe = Trim(t & "!!" & x)
End Function

Function SearchS(str As String) As String
    Dim searchArr() As String
    Dim splitStr As String, searchStr As String, token As String, tokenInput As String
    Dim SStr As String, rateStr As String, parseStr As String, returnStr As String, selectorCleanStr As String
    Dim hitCount As Long, i As Long, x As Long, y As Long, t As Long
    
    'pull token first
    tokenInput = Trim(Form_frmMainMenu.txtSToken)
    If Trim(tokenInput) = "" Or Trim(LCase(tokenInput)) = Trim(LCase(DefineFormDefaults("txtToken"))) Then
        tokenInput = Trim(Form_frmMainMenu.txtSTokenAdd)
    End If
        
    token = CheckToken(tokenInput)
    
    Debug.Print "TOKEN: " & token
    Debug.Print "INPUTSTR: " & str
    
    splitStr = Replace(Trim(str), "+++", "!!")

    searchArr = Split(splitStr, "!!")
    
    If UBound(searchArr) < 0 Then ThrowError 1969, "Search input is empty array; Input string: " & str: Exit Function
    
    returnStr = "" 'unique hits
    x = 0 ' counter for rate limit
    y = 0 ' number of items positive
    t = 0 ' total S hits
    For i = LBound(searchArr) To UBound(searchArr)
        searchStr = Trim(searchArr(i))
        
        'prevent dups
        If searchStr <> "" And LCase(searchStr) <> "null" And InStr(LCase(returnStr), LCase(searchStr) & "!!") = 0 Then
            
            selectorCleanStr = BuildSelectorClean(searchStr)
        
'            Debug.Print "^^^^^^^^^^^^^^^^" & vbLf & "!!!SEARCH STR:"
'            Debug.Print searchStr
'            Debug.Print "^^^^^^^^^^^^^^^"
            
            'rate limit for api
           If x Mod 500 = 0 And x > 0 Then
               rateStr = DefinePopupText(Trim(x), Trim(UBound(searchArr)), "rateLimitText")
               MsgBox rateStr, , "#TOO_MUCH_WINNING"
               Sleep 10000 '10 seconds
           End If
            
            'in mod2b S
            SStr = RunSBulkSearch(selectorCleanStr, token)
            Debug.Print "S SEARCH STR: " & SStr
 
            parseStr = ParseSSearch(SStr, searchStr)
            Debug.Print "PARSE STR: " & parseStr
            
            hitCount = Val(parseStr)
            FillTempGWSHits selectorCleanStr, hitCount
            
            'count hits
            t = t + hitCount
            If hitCount <> 0 Then
                'track unique hits searched
                returnStr = returnStr & searchStr & "!!"
                y = y + 1
            End If
        End If
    Next
    
    'trailing delim
    If Trim(returnStr) <> "" Then
        returnStr = Trim(Left(returnStr, Len(returnStr) - 2))
    End If
    
    'number of hits !! unique items searched $$ returnStr
    SearchS = t & "!!" & y & "$$" & returnStr
End Function
