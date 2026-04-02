*****************************************

MOD 2a_Sharepoint_Tables

*****************************************

Option Compare Database

Option Explicit

'RESET / GET DATA
Sub ResetLocalTablesSP()
    Dim tblArr As Variant, listArr As Variant
    Dim tblName As String, listName As String
    Dim i As Long
    
    tblArr = Array("localNorks", "localSelectors", "localTargets", "tempSchema", "tempGWSearchResults", "tempSSearchResults")
    listArr = Array("Norks", "Selectors", "Targets")
    
    'clear everything
    For i = 0 To 5
        tblName = tblArr(i)
        ClearTableList tblName
    Next
    
    'get SP data
    For i = 0 To 2
        tblName = tblArr(i)
        listName = listArr(i)
        
        SendDataLocalSP listName, tblName, tblName
        UpdateLocalTableCount tblName
    Next
    
    UpdateLocalTableCount "selectorsTargetId"
End Sub

'++++++++++++++++++++++++++++++

'ADD DATA SECTION

'single move data function, returns record count
Sub SendDataLocalSP(fromTbl As String, toTbl As String, localTbl As String, Optional howMany As Long = 0)
    Dim db As DAO.Database
    Dim fieldStr As String, strSQL As String, localStr As String, selectorStr As String
    
    'build fields from local table
    fieldStr = BuildLocalFieldStr(localTbl)
    
    Set db = CurrentDb
   
    'Debug.Print fieldStr
        
    strSQL = "INSERT INTO [" & toTbl & "] (" & fieldStr & ") SELECT " & fieldStr & " FROM " & "[" & fromTbl & "]"
    
    If howMany <> 0 Then
        strSQL = "INSERT INTO [" & toTbl & "] (" & fieldStr & ") SELECT TOP " & howMany & " " & fieldStr & " " & _
        "FROM " & "[" & fromTbl & "] ORDER BY [" & fromTbl & "].[ID] DESC"
    End If
    'Debug.Print "STR SQL: " & strSQL
        
    db.Execute strSQL, dbFailOnError
End Sub


Sub SendDataUpdateSP(fromTbl As String, toTbl As String, matchStr As String, updateStr As String)
    Dim db As DAO.Database
    Dim strSQL As String
    
    Set db = CurrentDb
    
    strSQL = "UPDATE " & toTbl & " INNER JOIN " & fromTbl & " " & _
    "ON " & toTbl & "." & matchStr & " = " & fromTbl & "." & matchStr & " " & _
    "SET " & toTbl & "." & updateStr & " = " & fromTbl & "." & updateStr & " " & _
    "WHERE " & toTbl & "." & updateStr & " IS NULL OR " & toTbl & "." & updateStr & " = ''"
    
    'Debug.Print "SEND DATA UPDATE STR SQL: " & strSQL
    
    db.Execute strSQL, dbFailOnError
End Sub

'+++++++++++++++++++++++++

'SEARCH SECTION

Function SearchSelectorTbl(str As String, Optional cleanStr As String = "") As DAO.Recordset
    Dim searchStr As String, selectorCleanStr As String

    searchStr = Trim(str)
    selectorCleanStr = Trim(cleanStr)
    If selectorCleanStr = "" Then selectorCleanStr = BuildSelectorClean(searchStr)
    If searchStr = "" Or selectorCleanStr = "" Then ThrowError 1961, str: Exit Function

    'SEARCH BY SELECTOR CLEAN
    Set SearchSelectorTbl = OpenRS("SELECT * FROM [localSelectors] WHERE [selectorClean] = ?", selectorCleanStr)
End Function


Function SearchAddStrTbl(str As String) As String
    Dim arr() As String, rs As DAO.Recordset
    Dim splitStr As String, searchStr As String, selectorCleanStr As String
    Dim targetIdStr As String, targetName As String, targetStr As String, targetItem As String
    Dim recordCount As Long, x As Long, i As Long, t As Long
    
    'treat everything as separate (split on row AND item)
    splitStr = Replace(str, "+++", "!!")
    arr = Split(splitStr, "!!")
    
    If UBound(arr) = -1 Then SearchAddStrTbl = "":  Exit Function
    
    'targetStr = "" 'track unique targets
    't = 0 'gw targets
    x = 0 'selector hits
    For i = LBound(arr) To UBound(arr)
        searchStr = Trim(arr(i))
        If searchStr <> "" And LCase(searchStr) <> "null" Then
            selectorCleanStr = BuildSelectorClean(searchStr)
       
            'Debug.Print "****SEARCH STR: " & searchStr
            Set rs = SearchSelectorTbl(searchStr, selectorCleanStr)
            
            'count hits
            If Not rs.EOF Then
                rs.MoveFirst
                rs.MoveLast
                recordCount = rs.recordCount
                
                If recordCount <> 0 Then
                    x = x + 1
                End If
            End If
            
            FillTempGWSearchResults rs, searchStr, selectorCleanStr
            
            'targetItem = SearchTargetItemGWTbl(searchStr)
            targetItem = SearchSelectorForTargetIdTbl(selectorCleanStr)
            If targetItem <> "" And InStr(LCase(targetStr), LCase(targetItem) & "!!") = 0 Then
                'targetStr = targetStr & targetItem & "!!" 'track uniques
                t = t + 1
            End If
        End If
    Next
    
'    If Trim(targetStr) <> "" Then
'        targetStr = Trim(Left(targetStr, Len(targetStr) - 2))
'    End If
    
    'return is target hits !! selector hits $$ targetStr
    'SearchAddStrTbl = t & "!!" & x & "$$" & targetStr
    SearchAddStrTbl = Trim(t & "!!" & x)
End Function


Function SearchTargetNameTbl(str As String) As String
    Dim rs As DAO.Recordset
    Dim targetIdStr As String

    targetIdStr = Trim(str)

    Set rs = OpenRS("SELECT * FROM [localTargets] WHERE [targetId] = ?", targetIdStr)

    'targetId not found
     If (rs.EOF And rs.BOF) Then SearchTargetNameTbl = "": Exit Function

     'return target name
     SearchTargetNameTbl = Nz(rs!targetName, "")
End Function

'returns targetId of selectorId in Selector list
Function SearchSelectorForTargetIdTbl(str As String) As String
    Dim rs As DAO.Recordset
    Dim selectorCleanStr As String

    selectorCleanStr = Trim(str)

    Set rs = OpenRS("SELECT * FROM [localSelectors] WHERE [selectorClean] = ?", selectorCleanStr)

    'throw error on null return?
    If (rs.EOF And rs.BOF) Then SearchSelectorForTargetIdTbl = "": Exit Function

    SearchSelectorForTargetIdTbl = Nz(rs!targetId, "")
End Function

Function SearchAllTargetHitsGWTbl() As String
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim strSQL As String, returnStr As String, itemStr As String, nameStr As String
    Dim i As Long
    
    Set db = CurrentDb
    
    'query lang per claude
    strSQL = "SELECT * FROM [tempGWSearchResults] WHERE [targetId] IS NOT NULL AND [targetId] <> ''"
    
    Set rs = db.OpenRecordset(strSQL)
    
    If rs.EOF Then SearchAllTargetHitsGWTbl = "": Exit Function
    
    rs.MoveLast
    rs.MoveFirst
    
    'arrow garbage below, do NOT care
    returnStr = ""
    For i = 1 To rs.recordCount
        itemStr = Nz(rs!targetId, "")
        nameStr = Nz(rs!targetName, "")
        If Trim(itemStr) <> "" Then
            'name set
            If Trim(nameStr) <> "" And InStr(LCase(nameStr), "[real name not set") = 0 Then
                'dupes
                If InStr(LCase(returnStr), LCase(nameStr) & "!!") = 0 Then
                    returnStr = returnStr & nameStr & "!!"
                End If
            Else
                'name not set return targetId
                If InStr(LCase(returnStr), LCase(itemStr) & "!!") = 0 Then
                    returnStr = returnStr & itemStr & "!!"
                End If
            End If
        End If
        rs.MoveNext
    Next
     
    If Trim(returnStr) <> "" Then
        returnStr = Trim(Left(returnStr, Len(returnStr) - 2))
    End If
    
    SearchAllTargetHitsGWTbl = returnStr
End Function

'++++++++++++++++++++++++++++++++++++++++++++

Function FillLocalSelectors(str As String, selectorType As String, Optional targetId As String = "", Optional dataSource As String = "") As String
    Dim db As DAO.Database, rs As DAO.Recordset, rsSearch As DAO.Recordset
    Dim inputStr As String, selectorCleanStr As String, selectorIdStr As String

    inputStr = Trim(str)

    'null input
    If inputStr = "" Or UCase(inputStr) = "NULL" Then Exit Function

    'CALC SELECTOR CLEAN HERE (in clean)
    selectorCleanStr = BuildSelectorClean(inputStr, selectorType)
    'Debug.Print "^^^ SELECTORCLEANSTR: " & selectorCleanStr

    'check if already in SP
    'when targetId provided (manual add), only block if exact selector+target combo exists
    'when targetId empty (import flow), block any duplicate selector value
    If Trim(targetId) <> "" And Trim(UCase(targetId)) <> "NULL" Then
        Set rsSearch = OpenRS("SELECT * FROM [localSelectors] WHERE [selectorClean] = ? AND [targetId] = ?", selectorCleanStr, Trim(targetId))
    Else
        Set rsSearch = OpenRS("SELECT * FROM [localSelectors] WHERE [selectorClean] = ?", selectorCleanStr)
    End If

    'record set has records (already in selectors list)
    If Not (rsSearch.EOF And rsSearch.BOF) Then FillLocalSelectors = "SELECTOR ALREADY KNOWN | ID: " & _
    Nz(rsSearch!selectorId, "NULL"): rsSearch.Close: Exit Function

    rsSearch.Close
    Set db = CurrentDb
    Set rs = db.OpenRecordset("localSelectors", dbOpenDynaset)
    
    selectorIdStr = "S" & DefineUniqueId()
    
    'insert values
    rs.AddNew
    rs!selectorId = selectorIdStr
    rs!selector = inputStr
    rs!selectorClean = selectorCleanStr
    rs!dateCreated = Now()
    rs!lastUpdated = Now()
    rs!createdBy = Trim(LCase(Environ("USERNAME")))
    rs!lastUpdatedBy = Trim(LCase(Environ("USERNAME")))
    rs!selectorType = Trim(LCase(selectorType))
    rs!Title = selectorIdStr 'sp req
    If Trim(dataSource) <> "" Then
        rs!dataSource = dataSource
    Else
        rs!dataSource = "GrayWolfe Upload"
    End If

    If Trim(UCase(targetId)) <> "NULL" And Trim(targetId) <> "" Then rs!targetId = targetId
    rs.Update
    rs.Close
    
    FillLocalSelectors = selectorIdStr
    
    'Debug.Print "INPUTSTR: " & inputStr & " | SELECTORTYPE: " & selectorType & " successfully added to LOCAL SELECTORS table"
End Function

Function FillLocalTargets() As String
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim targetIdStr As String, targetNameStr As String
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("localTargets", dbOpenDynaset)
    
    'Debug.Print "ADDING TO TARGETS LIST, NAMESTr: " & nameStr
    
    'build target id
    targetIdStr = "T" & DefineUniqueId()
    
    'always set name to not known?
    targetNameStr = "[REAL name NOT set / known]"
    
    'insert values
    rs.AddNew
    rs!targetId = targetIdStr
    rs!targetName = targetNameStr
    rs!Title = targetIdStr
    rs!dateCreated = Now()
    rs!lastUpdated = Now()
    rs!createdBy = Trim(LCase(Environ("USERNAME")))
    rs!lastUpdatedBy = Trim(LCase(Environ("USERNAME")))
    rs!dataSource = "GrayWolfe Upload"

    rs.Update
    
    FillLocalTargets = targetIdStr
End Function

'takes recordset as input, uses to fill tbl
Sub FillTempGWSearchResults(rsInput As DAO.Recordset, Optional searchStr As String = "", Optional cleanStr As String = "")
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim typeStr As String, targetIdStr As String, targetNameStr As String, selectorCleanStr As String
    Dim recordCount As Long, i As Long
    
    'write to temp table
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tempGWSearchResults", dbOpenDynaset)
    
    'empty inputs
    If rsInput.EOF Then
        rs.AddNew
        rs!selector = searchStr
        rs!selectorClean = cleanStr
        rs!selectorId = searchStr 'set id to whats searched to make uniuqe
        rs!inGrayWolfe = "NO"
        'rs!SHits = "N/A" 'overwritten later

        
        rs.Update
        Exit Sub
    End If
        
    'handle hits
    rsInput.MoveLast
    rsInput.MoveFirst
    recordCount = rsInput.recordCount

    'fix type format
    typeStr = StrConv(Nz(rsInput!selectorType, ""), vbProperCase)
    If Trim(LCase(typeStr)) = "ip" Then typeStr = "IP"

    targetNameStr = ""
    For i = 1 To recordCount
        If Not rsInput.EOF Then
        
            'calc connnected target (get the name)
            targetIdStr = Nz(rsInput!targetId, "")
'            Debug.Print "^^^^^^^^^^"
'            Debug.Print "TARGET ID STR: " & targetIdStr
            
            'not a target
            If Trim(targetIdStr) = "" Then
                targetNameStr = "[NO known connections]"
            Else
                targetNameStr = SearchTargetNameTbl(targetIdStr)
                
                If Trim(targetNameStr) = "" Then targetNameStr = "[REAL name NOT set / known]"
            End If
            
'            Debug.Print "TARGET NAME STR: " & targetNameStr
'            Debug.Print "^^^^^^^^^^^^"
            'calc connected nork later
        
            rs.AddNew
            rs!selector = searchStr
            rs!selectorId = rsInput!selectorId
            rs!selectorClean = cleanStr
            
            'rs!SHits = "N/A" 'overwritten later
            
            rs!inGrayWolfe = "YES"
            'rs!addedToGrayWolfe = "N/A"
            rs!targetName = targetNameStr

            rs!dateCreated = rsInput!dateCreated
            rs!createdBy = rsInput!createdBy
            rs!lastUpdated = rsInput!lastUpdated
            rs!lastUpdatedBy = rsInput!lastUpdatedBy
            rs!targetId = rsInput!targetId
            rs!norkId = rsInput!norkId
            rs!username = rsInput!username
            rs!displayName = rsInput!displayName
            rs!selectorType = typeStr
    
            rs.Update
        End If
        If i < recordCount Then rsInput.MoveNext
    Next
    rsInput.Close
    rs.Close
End Sub

Sub FillTempGWSHits(str As String, hits As Long)
    Dim rs As DAO.Recordset
    Dim selectorCleanStr As String

    selectorCleanStr = Trim(str)
    'If searchStr = "" Or hits = 0 Then Exit Sub
    If selectorCleanStr = "" Then Exit Sub

    Set rs = OpenRS("SELECT * FROM [tempGWSearchResults] WHERE [selectorClean] = ?", selectorCleanStr)

    If Not rs.EOF Then
        ExecSQL "UPDATE [tempGWSearchResults] SET [SHits] = " & hits & " WHERE [selectorClean] = ?", selectorCleanStr
    End If
End Sub



Sub FillTempSchema(str As String, colMax As Long)
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim arr() As String, rowArr() As String
    Dim rowStr As String, addStr As String, addType As String
    Dim rowMax As Long, i As Long, j As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tempSchema", dbOpenDynaset)
    
    'calc max columns
    'rowMax = (colMax * 2) + 1
    'Debug.Print "ROW MAX: " & rowMax
    
    arr = Split(str, "+++")

    For i = LBound(arr) To UBound(arr)
        rowStr = Trim(arr(i))

        If rowStr <> "" Then
            rowArr = Split(rowStr, "!!")

            rs.AddNew

            'For j = LBound(rowArr) To UBound(rowArr)
            For j = LBound(rowArr) To colMax - 1
                If j > UBound(rowArr) Then
                    addStr = "NULL"
                Else
                    addStr = Trim(rowArr(j))
                End If
                
                addType = CleanSelectorInput(DetectSelectorType(addStr))
                'Debug.Print "ADD STR: " & addStr
                'Debug.Print "ADD TYPE: " & addType

                'add to temp table
                If addStr = "" Then addStr = "NULL"
                If addType = "" Then addType = "NULL"

                rs.Fields("ColumnStr" & j + 1).Value = addStr
                rs.Fields("ColumnType" & j + 1).Value = addType
            Next

            rs.Update
        End If
    Next
    
    rs.Close 'close after adding data, then reopen per claude
End Sub

Sub FillTargetDisplayForm(targetId As String)
    Dim rs As DAO.Recordset, ctl As Control
    Dim typeArr As Variant, typeItem As String
    Dim targetIdStr As String
    Dim returnStr As String, itemStr As String, ctlName As String
    Dim recordCount As Long, i As Long, j As Long

    targetIdStr = Trim(targetId)
    'Debug.Print "TARGET ID FILL: " & targetIdStr

    'null input
    If targetIdStr = "" Then Exit Sub

    'below is Array("persona name", "street address", "email", "phone", "ip", "other")
    typeArr = DefineFormTypeArr()

    For i = LBound(typeArr) To UBound(typeArr)
        typeItem = Trim(typeArr(i))
        'Debug.Print "TYPE ITEM: " & typeItem

        Set rs = OpenRS("SELECT * FROM [localSelectors] WHERE [targetId] = ? AND [selectorType] = ?", targetIdStr, typeItem)
        
        returnStr = ""
        If Not rs.EOF Then
    
            rs.MoveLast
            rs.MoveFirst
            recordCount = rs.recordCount
            'Debug.Print "RECORD COUNT: " & recordCount
            
            'returnStr = ""
            For j = 1 To recordCount
                itemStr = Nz(rs!selector, "")
                
                If Trim(itemStr) <> "" Then returnStr = returnStr & itemStr & " | "
                
                If j < recordCount Then rs.MoveNext
            Next
            
            'Debug.Print "RETURN STR: " & returnStr
            
            'remove trailing delim and populate
            If Trim(returnStr) <> "" Then
                returnStr = Trim(Left(returnStr, Len(returnStr) - 2))
            
                ctlName = TargetFormDisplayMap(typeItem)
                'Debug.Print "CTL NAME: " & ctlName
                'Debug.Print "RETURN STR: " & returnStr
                If ctlName <> "" Then
                    Forms("frmTargetDetails").Controls(ctlName).Value = returnStr
                End If
            End If
        End If
    Next

End Sub

Sub FillTargetLaptopCount(targetId As String)
    Dim rs As DAO.Recordset
    Dim targetIdStr As String, itemStr As String
    Dim recordCount As Long, i As Long

    targetIdStr = Trim(targetId)

    'null input
    If targetIdStr = "" Then Exit Sub

    Set rs = OpenRS("SELECT * FROM [localTargets] WHERE [targetId] = ?", targetIdStr)

    'set to 0
    If rs.EOF Then
        Forms("frmTargetDetails").txtLaptopCount.Value = 0
        Form_frmTargetDetails.laptopsCurrent = 0
        Exit Sub
    End If
    
    rs.MoveLast
    rs.MoveFirst
    recordCount = rs.recordCount
    
    For i = 1 To recordCount
        itemStr = Nz(rs!laptopCount, 0)
        
        If i < recordCount Then rs.MoveNext
    Next
    
    Forms("frmTargetDetails").txtLaptopCount.Value = itemStr
    
    'update the tracking variable
    Form_frmTargetDetails.laptopsCurrent = Val(itemStr)
End Sub


'+++++++++++++++++++++++++++++++++++

'BUILD THINGS SECTION
Function BuildLocalFieldStr(tblName As String) As String
    Dim db As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field
    Dim returnStr As String
    Dim i As Long
    
    Set db = CurrentDb
    
    Set tdf = db.TableDefs(tblName)
    If tdf Is Nothing Then ThrowError 1964, "CANT GET LOCAL FIELD NAMES FROM TABLE: " & tblName: Exit Function
    
    returnStr = ""
    For i = 0 To tdf.Fields.Count - 1
        Set fld = tdf.Fields(i)
        
        'If Trim(fld.Name) <> "" And Not fld.Attributes And Not dbAutoIncrField Then
        If Trim(fld.Name) <> "" And Trim(fld.Name) <> "ID" Then
        'If Trim(fld.Name) <> "" Then
            returnStr = returnStr & "[" & fld.Name & "], "
        End If
    Next
    
    'remove trailing delim
    BuildLocalFieldStr = Trim(Left(returnStr, Len(returnStr) - 2))
End Function




'+++++++++++++++++++++++++++++++++++++++++++++

'UPDATE / GET DATA SECTION

Function UpdateLocalTableCount(tblName As String) As Long
    Dim x As Long
    
    x = GetLocalTableCount(tblName)
    
    'half assed version below, real (polymorphic) answer not worth it
    Select Case tblName
    
    Case "localNorks"
        Form_frmMainMenu.localNorksCount = x
    
    Case "localSelectors"
        Form_frmMainMenu.localSelectorsCount = x
    
    Case "localTargets"
        Form_frmMainMenu.localTargetsCount = x
    
'    Case "targetsName"
'        Form_frmMainMenu.localTargetsNameCount = x
        
    Case "selectorsTargetId"
        Form_frmMainMenu.localSelectorsTargetIdCount = x
    
    End Select
    
    
    UpdateLocalTableCount = x
End Function

Function GetLocalTableCount(tbl As String) As Long
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim strSQL As String, tblName As String
    
    tblName = Trim(tbl)
    
    Select Case tblName
    
    'WANT USER TO MANUALLY SET
'    Case "targetsName"
'        strSQL = "SELECT COUNT(*) as Count FROM [localTargets] WHERE [targetName] IS NOT NULL AND [targetName] <> ''"
        
    Case "selectorsTargetId"
        strSQL = "SELECT COUNT(*) as Count FROM [localSelectors] WHERE [targetId] IS NOT NULL AND [targetId] <> ''"
        
    Case "knownTargets"
        strSQL = "SELECT COUNT(*) as Count FROM [tempGWSearchResults] WHERE [targetName] IS NOT NULL AND [targetName] <> '[REAL name NOT set / known]'"
    Case Else
        strSQL = "SELECT COUNT(*) AS Count FROM [" & tblName & "]"
        
    End Select
    
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    GetLocalTableCount = rs!Count
End Function


'delimited string as input
Sub UpdateSelectorsTblTargetId(str As String, targetId As String)
    Dim inputArr() As String, targetIdStr As String, selectorStr As String
    Dim i As Long

'    Debug.Print "!!!!!!" & vbLf & "UPDATE SELECTORS TBL TARGET ID:" & vbLf
'    Debug.Print "INPUT STR: " & str
'    Debug.Print "TARGET ID STR: " & targetId

    inputArr = Split(Trim(str), "!!")
    targetIdStr = Trim(targetId)

    If targetIdStr = "" Or UBound(inputArr) = -1 Then Exit Sub

    For i = LBound(inputArr) To UBound(inputArr)
        selectorStr = Trim(inputArr(i))

        'only update selectors that have no target assigned (preserve shared selectors)
        ExecSQL "UPDATE [localSelectors] SET [targetId] = ? WHERE [selector] = ? AND ([targetId] IS NULL OR [targetId] = '')", targetIdStr, selectorStr
    Next

End Sub

Sub UpdateGWSearchTargetId(str As String, targetId As String)
    Dim rsTarget As DAO.Recordset
    Dim inputArr() As String, targetIdStr As String, targetNameStr As String
    Dim selectorStr As String
    Dim i As Long

    inputArr = Split(Trim(str), "!!")
    targetIdStr = Trim(targetId)

    If targetIdStr = "" Or UBound(inputArr) = -1 Then Exit Sub

    For i = LBound(inputArr) To UBound(inputArr)
        selectorStr = Trim(inputArr(i))

        'get target name for display
        Set rsTarget = OpenRS("SELECT * FROM [localTargets] WHERE [targetId] = ?", targetIdStr)
        If Not rsTarget.EOF Then
            targetNameStr = Nz(rsTarget!targetName, "")
        Else
            targetNameStr = ""
        End If

        'only update search results that have no target assigned (preserve shared selectors)
        ExecSQL "UPDATE [tempGWSearchResults] SET [targetId] = ?, [targetName] = ? WHERE [selector] = ? AND ([targetId] IS NULL OR [targetId] = '')", targetIdStr, targetNameStr, selectorStr
    Next

End Sub

Sub UpdateTargetsSelectorCountTbl(targetId As String)
    Dim rs As DAO.Recordset
    Dim targetIdStr As String
    Dim x As Long

    targetIdStr = Trim(targetId)

    'empty input
    If targetIdStr = "" Then ThrowError 1962, "UPDATE SELECTOR COUNT BLANK TARGET ID; INPUT: " & targetId: Exit Sub

    Set rs = OpenRS("SELECT * FROM [localSelectors] WHERE [targetId] = ?", targetIdStr)

    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
        x = rs.recordCount
    Else
        x = 0
    End If
    rs.Close

    'Debug.Print "SELECTOR COUNT X: " & x

    ExecSQL "UPDATE [localTargets] SET [selectorCount] = " & x & " WHERE [targetId] = ?", targetIdStr
End Sub

Sub UpdateAddToGW(str As String)
    Dim inputStr As String

    inputStr = Trim(str)
    If inputStr = "" Then Exit Sub

    ExecSQL "UPDATE [tempGWSearchResults] SET [addedToGrayWolfe] = 'YES', " & _
    "[lastUpdated] = ?, [lastUpdatedBy] = ? WHERE [selector] = ?", _
    CStr(Now()), Trim(LCase(Environ("USERNAME"))), inputStr
End Sub

Sub UpdateLastUpdated(str As String, tbl As String)
    Dim targetStr As String, tblStr As String

    targetStr = Trim(str)
    tblStr = Trim(tbl)
    If tblStr = "" Or targetStr = "" Then Exit Sub

    ExecSQL "UPDATE [" & tblStr & "] SET [lastUpdated] = ?, [lastUpdatedBy] = ? WHERE [targetId] = ?", _
    CStr(Now()), Trim(LCase(Environ("USERNAME"))), targetStr
End Sub

Sub UpdateTargetStatsForm(targetId As String, inputItem As String, fieldItem As String)
    Dim targetIdStr As String, inputStr As String, fieldStr As String

    targetIdStr = Trim(targetId)
    inputStr = Trim(inputItem)
    fieldStr = Trim(fieldItem)

    If targetId = "" Or fieldStr = "" Then Exit Sub

    ExecSQL "UPDATE [localTargets] SET [" & fieldStr & "] = ? WHERE [targetId] = ?", inputStr, targetIdStr

End Sub

'for Form UI
Sub UpdateTargetSelectors(targetId As String, inputItem As String, currentItem As String, selectorType As String, Optional dataSource As String = "")
    Dim arr() As String, itemStr As String
    Dim targetIdStr As String, inputStr As String, selectorTypeStr As String
    Dim currentStr As String, deleteStr As String, addStr As String
    Dim i As Long
    
    targetIdStr = Trim(targetId)
    inputStr = Trim(inputItem)
    currentStr = Trim(currentItem)
    selectorTypeStr = Trim(selectorType)
    
    'delete items
    deleteStr = DetectStrDiff(currentStr, inputStr, "!!")
    arr = Split(deleteStr, "!!")
    
    If UBound(arr) > -1 Then
        
        For i = LBound(arr) To UBound(arr)
            itemStr = arr(i)

            DeleteTargetSelector targetIdStr, itemStr
        Next
    End If
    
    'add items
    addStr = DetectStrDiff(inputStr, currentStr, "!!")
    arr = Split(addStr, "!!")
    
    If UBound(arr) > -1 Then
    
        For i = LBound(arr) To UBound(arr)
            itemStr = arr(i)

            FillLocalSelectors itemStr, selectorTypeStr, targetIdStr, dataSource
        Next
    End If

End Sub


'++++++++++++++++++++++++++++++++++++++++++++

'MERGE / BRIDGING SECTION

Function MergeTargets(keepTargetId As String, absorbTargetId As String) As Boolean
    Dim rsKeep As DAO.Recordset, rsAbsorb As DAO.Recordset
    Dim keepStr As String, absorbStr As String
    Dim keepName As String, absorbName As String, keepCase As String, absorbCase As String
    Dim keepLaptops As Long, absorbLaptops As Long

    keepStr = Trim(keepTargetId)
    absorbStr = Trim(absorbTargetId)

    If keepStr = "" Or absorbStr = "" Or keepStr = absorbStr Then
        MergeTargets = False
        Exit Function
    End If

    'validate both exist
    Set rsKeep = OpenRS("SELECT * FROM [localTargets] WHERE [targetId] = ?", keepStr)
    If rsKeep.EOF Then MergeTargets = False: rsKeep.Close: Exit Function

    Set rsAbsorb = OpenRS("SELECT * FROM [localTargets] WHERE [targetId] = ?", absorbStr)
    If rsAbsorb.EOF Then MergeTargets = False: rsKeep.Close: rsAbsorb.Close: Exit Function

    'move all selectors from absorbed to kept
    ExecSQL "UPDATE [localSelectors] SET [targetId] = ? WHERE [targetId] = ?", keepStr, absorbStr

    'remove duplicate selectors that existed in both targets
    ExecSQL "DELETE FROM [localSelectors] WHERE [ID] NOT IN (" & _
        "SELECT MIN([ID]) FROM [localSelectors] WHERE [targetId] = ? GROUP BY [selectorClean]" & _
    ") AND [targetId] = ?", keepStr, keepStr

    'copy metadata from absorbed if kept lacks it
    keepName = Nz(rsKeep!targetName, "")
    absorbName = Nz(rsAbsorb!targetName, "")
    If (keepName = "" Or InStr(LCase(keepName), "[real name not set") > 0) And absorbName <> "" And InStr(LCase(absorbName), "[real name not set") = 0 Then
        ExecSQL "UPDATE [localTargets] SET [targetName] = ? WHERE [targetId] = ?", absorbName, keepStr
    End If

    keepCase = Nz(rsKeep!caseNumber, "")
    absorbCase = Nz(rsAbsorb!caseNumber, "")
    If keepCase = "" And absorbCase <> "" Then
        ExecSQL "UPDATE [localTargets] SET [caseNumber] = ? WHERE [targetId] = ?", absorbCase, keepStr
    End If

    'sum laptop counts
    keepLaptops = Val(Nz(rsKeep!laptopCount, 0))
    absorbLaptops = Val(Nz(rsAbsorb!laptopCount, 0))
    If absorbLaptops > 0 Then
        ExecSQL "UPDATE [localTargets] SET [laptopCount] = " & (keepLaptops + absorbLaptops) & " WHERE [targetId] = ?", keepStr
    End If

    rsKeep.Close
    rsAbsorb.Close

    'recalc selector count
    UpdateTargetsSelectorCountTbl keepStr

    'delete absorbed target
    ExecSQL "DELETE FROM [localTargets] WHERE [targetId] = ?", absorbStr

    'update last updated
    UpdateLastUpdated keepStr, "localTargets"

    MergeTargets = True
End Function

Function CollectTargetIdsForRow(inputStr As String) As String
    Dim rs As DAO.Recordset
    Dim inputArr() As String, selectorCleanStr As String
    Dim dict As Object, targetIdItem As String
    Dim returnStr As String
    Dim i As Long

    Set dict = CreateObject("Scripting.Dictionary")

    inputArr = Split(Trim(inputStr), "!!")
    If UBound(inputArr) = -1 Then CollectTargetIdsForRow = "": Exit Function

    For i = LBound(inputArr) To UBound(inputArr)
        If Trim(inputArr(i)) <> "" And LCase(Trim(inputArr(i))) <> "null" Then
            selectorCleanStr = BuildSelectorClean(Trim(inputArr(i)))

            Set rs = OpenRS("SELECT DISTINCT targetId FROM [localSelectors] WHERE [selectorClean] = ? AND [targetId] IS NOT NULL AND [targetId] <> ''", selectorCleanStr)

            Do While Not rs.EOF
                targetIdItem = Nz(rs!targetId, "")
                If targetIdItem <> "" And Not dict.Exists(targetIdItem) Then
                    dict.Add targetIdItem, True
                End If
                rs.MoveNext
            Loop
            rs.Close
        End If
    Next

    'build return string
    If dict.Count = 0 Then CollectTargetIdsForRow = "": Exit Function

    returnStr = ""
    Dim key As Variant
    For Each key In dict.Keys
        returnStr = returnStr & key & "!!"
    Next

    'remove trailing delim
    CollectTargetIdsForRow = Trim(Left(returnStr, Len(returnStr) - 2))
End Function

'+++++++++++++++++++++++++++++++++++++++++++

Sub ClearTableList(tbl As String)
    Dim db As DAO.Database

    Set db = CurrentDb
    
    On Error Resume Next
    db.Execute "DELETE * FROM " & tbl, dbFailOnError
    On Error GoTo 0

    Set db = Nothing

End Sub

Sub DeleteTargetSelector(targetId As String, inputItem As String)
    Dim inputStr As String, targetIdStr As String

    inputStr = Trim(inputItem)
    targetIdStr = Trim(targetId)

    If targetIdStr = "" Then Exit Sub

    ExecSQL "DELETE FROM [localSelectors] WHERE [targetId] = ? AND [selector] = ?", targetIdStr, inputStr

End Sub

'deletes everythign except ID field
Sub DeleteTempFields(tblName As String)
    Dim db As DAO.Database, tdf As DAO.TableDef, tbl As DAO.TableDef
    Dim i As Long

    Set db = CurrentDb
    Set tbl = db.TableDefs(tblName)

    For i = tbl.Fields.Count - 1 To 0 Step -1
        If tbl.Fields(i).Name <> "ID" Then
            tbl.Fields.Delete tbl.Fields(i).Name
        End If
    Next

End Sub

Sub ExportBackupCSV()
    Dim fd As FileDialog
    Dim folderPath As String, timestamp As String, filePath As String
    Dim tblArr As Variant, tblName As String
    Dim i As Long

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select Backup Folder"
    If fd.Show = 0 Then Exit Sub
    folderPath = fd.SelectedItems(1)

    timestamp = Format(Now(), "YYYYMMDD_HHNNSS")
    tblArr = Array("localNorks", "localSelectors", "localTargets")

    For i = LBound(tblArr) To UBound(tblArr)
        tblName = tblArr(i)
        filePath = folderPath & "\" & tblName & "_" & timestamp & ".csv"
        DoCmd.TransferText acExportDelim, , tblName, filePath, True
    Next

    MsgBox "Exported " & UBound(tblArr) + 1 & " tables to:" & vbLf & folderPath, vbInformation, "Backup Complete"
End Sub

''''''''''''''''''''''
