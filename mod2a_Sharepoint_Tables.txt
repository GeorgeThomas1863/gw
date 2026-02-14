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
    
    Debug.Print "RESETTING DATA" & vbLf & "--------------------"
    
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
    
    'UpdateLocalTableCount "targetsName"
    UpdateLocalTableCount "selectorsTargetId"
    
'    Debug.Print "SP NORKS COUNT: " & Form_frmMainMenu.localNorksCount
'    Debug.Print "SP SELECTORS COUNT: " & Form_frmMainMenu.localSelectorsCount
'    Debug.Print "SP TARGETS COUNT: " & Form_frmMainMenu.localTargetsCount
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
    Debug.Print "SEND DATA STR SQL: " & strSQL
    
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

'MAKE PARAMATIZED LATER
Function SearchSelectorTbl(str As String, Optional cleanStr As String = "") As DAO.Recordset
    Dim db As DAO.Database, rs As DAO.Recordset, qdf As DAO.QueryDef, tdf As DAO.TableDef
    Dim searchStr As String, strSQL As String, returnStr As String, selectorCleanStr As String
    Dim i As Long, x As Long
    
    searchStr = Trim(str)
    selectorCleanStr = Trim(cleanStr)
    If selectorCleanStr = "" Then selectorCleanStr = BuildSelectorClean(searchStr)
    If searchStr = "" Or selectorCleanStr = "" Then ThrowError 1961, str: Exit Function
    
    Set db = CurrentDb
    
    'SEARCH BY SELECTOR CLEAN
    strSQL = "SELECT * FROM [localSelectors] WHERE [selectorClean] = '" & selectorCleanStr & "' "
    'Debug.Print "STRSQL: " & strSQL
    
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)
    'Debug.Print "SEARCH STR: " & searchStr & " | HITS: " & rs.recordCount
    Set SearchSelectorTbl = rs
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
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim strSQL As String, targetIdStr As String

    targetIdStr = Trim(str)

    Set db = CurrentDb

    strSQL = "SELECT * FROM [localTargets] WHERE [targetId] = '" & targetIdStr & "'"

    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    'targetId not found
     If (rs.EOF And rs.BOF) Then SearchTargetNameTbl = "": Exit Function

     'return target name
     SearchTargetNameTbl = Nz(rs!targetName, "")
End Function

'returns targetId of selectorId in Selector list
Function SearchSelectorForTargetIdTbl(str As String) As String
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim strSQL As String, selectorCleanStr As String
    Dim i As Long
    
    selectorCleanStr = Trim(str)
    
    Set db = CurrentDb
    
    strSQL = "SELECT * FROM [localSelectors] WHERE [selectorClean] = '" & selectorCleanStr & "'"
    
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    'throw error on null return?
    'If rs.EOF And rs.BOF Then ThrowError 1962, "CANT FIND SELECTOR TO GET TARGET": Exit Function
    If (rs.EOF And rs.BOF) Then SearchSelectorForTargetIdTbl = "": Exit Function
    
    SearchSelectorForTargetIdTbl = Nz(rs!targetId, "")
End Function

'return a string of selectors with given target id
Function SearchTargetIdForSelectorsTbl(targetId As String)
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim strSQL As String, targetIdStr As String
    Dim returnStr As String, selectorStr As String, typeStr As String
    Dim recordCount As Long, i As Long
    
    targetIdStr = Trim(targetId)
    
    If targetIdStr = "" Or targetIdStr = "Null" Then SearchTargetIdForSelectorsTbl = "": Exit Function
    
    Set db = CurrentDb
    
    strSQL = "SELECT * FROM [localSelectors] WHERE [targetId] = '" & targetIdStr & "'"
    
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If (rs.EOF And rs.BOF) Then SearchTargetIdForSelectorsTbl = "": Exit Function
    
    rs.MoveFirst
    rs.MoveLast
    recordCount = rs.recordCount
    rs.MoveFirst
    
    returnStr = ""
    For i = 1 To recordCount
        If Not rs.EOF Then
            selectorStr = Nz(rs!selector, "")
            typeStr = Nz(rs!selectorType, "")
        
            If selectorStr <> "" And typeStr <> "" Then
                returnStr = returnStr & typeStr & "$$" & selectorStr & "!!"
            End If
        End If
        If i < recordCount Then rs.MoveNext
    Next
    
    SearchTargetIdForSelectorsTbl = Trim(Left(returnStr, Len(returnStr) - 2))
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

Function FillLocalSelectors(str As String, selectorType As String, Optional targetId As String = "") As String
    Dim db As DAO.Database, rs As DAO.Recordset, rsSearch As DAO.Recordset
    Dim inputStr As String, selectorCleanStr As String, selectorIdStr As String, strSQL As String
    
    inputStr = Trim(str)
    
    'null input
    If inputStr = "" Or UCase(inputStr) = "NULL" Then Exit Function
    
    'CALC SELECTOR CLEAN HERE (in clean)
    selectorCleanStr = BuildSelectorClean(inputStr, selectorType)
    'Debug.Print "^^^ SELECTORCLEANSTR: " & selectorCleanStr
    
    Set db = CurrentDb
    
    'check if already in SP
    strSQL = "SELECT * FROM [localSelectors] WHERE [selectorClean] = '" & selectorCleanStr & "'"
    
    Set rsSearch = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    'record set has records (already in selectors list)
    If Not (rsSearch.EOF And rsSearch.BOF) Then FillLocalSelectors = "SELECTOR ALREADY KNOWN | ID: " & _
    Nz(rsSearch!selectorId, "NULL"): rsSearch.Close: Exit Function
    
    rsSearch.Close
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
    rs!dataSource = "GrayWolfe Upload"
    
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
    
    'add nameStr if not null
    'If Trim(UCase(nameStr)) <> "NULL" And Trim(nameStr) <> "" Then rs!targetName = nameStr

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
        Debug.Print "SEARCH STR ADD: " & searchStr
        
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
            rs!targetId = targetIdStr
                       
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
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim selectorCleanStr As String, strSQL As String
    
    selectorCleanStr = Trim(str)
    'If searchStr = "" Or hits = 0 Then Exit Sub
    If selectorCleanStr = "" Then Exit Sub
    
    Set db = CurrentDb
    
    strSQL = "SELECT * FROM [tempGWSearchResults] WHERE [selectorClean] = '" & selectorCleanStr & "'"
    'Debug.Print "SOURCE STR SQL: " & strSQL
    
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If Not rs.EOF Then
    
        strSQL = "UPDATE [tempGWSearchResults] SET [SHits] = " & hits & " WHERE [selectorClean] = '" & selectorCleanStr & "'"
        db.Execute strSQL
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
    Dim db As DAO.Database, rs As DAO.Recordset, ctl As Control
    Dim typeArr As Variant, typeItem As String
    Dim strSQL As String, targetIdStr As String
    Dim returnStr As String, itemStr As String, ctlName As String
    Dim recordCount As Long, i As Long, j As Long
    
    targetIdStr = Trim(targetId)
    'Debug.Print "TARGET ID FILL: " & targetIdStr
    
    'null input
    If targetIdStr = "" Then Exit Sub
    
    Set db = CurrentDb
    
    'below is Array("persona name", "street address", "email", "phone", "ip", "other")
    typeArr = DefineFormTypeArr()
    
    For i = LBound(typeArr) To UBound(typeArr)
        typeItem = Trim(typeArr(i))
        'Debug.Print "TYPE ITEM: " & typeItem
        
        strSQL = "SELECT * FROM [localSelectors] WHERE targetId = '" & targetIdStr & "' AND selectorType = '" & typeItem & "'"
        'Debug.Print "STR SQL: " & strSQL
        
        Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
        
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
                On Error Resume Next 'should remove
                    Forms("frmTargetDetails").Controls(ctlName).Value = returnStr
                End If
            End If
        End If
    Next

End Sub

Sub FillTargetLaptopCount(targetId As String)
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim targetIdStr As String, strSQL As String, itemStr As String
    Dim recordCount As Long, i As Long
    
    targetIdStr = Trim(targetId)
    
    'null input
    If targetIdStr = "" Then Exit Sub
    
    Set db = CurrentDb
    
    strSQL = "SELECT * FROM [localTargets] WHERE targetId = '" & targetIdStr & "'"
    
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    'set to 0
    If rs.EOF Then
        Forms("frmTargetDetails").txtLaptopCount.Value = 0
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



Sub BuildTempFields(colMax As Long, tblName As String)
    Dim db As DAO.Database, tbl As DAO.TableDef, fld As DAO.Field
    Dim i As Long
    
    Set db = CurrentDb
    Set tbl = db.TableDefs(tblName)
    
    'build temp tbl fields
    For i = 1 To colMax
        Set fld = tbl.CreateField("ColumnStr" & i, dbText, 255)
        tbl.Fields.Append fld

        Set fld = tbl.CreateField("ColumnType" & i, dbText, 255)
        tbl.Fields.Append fld
    Next
    
    db.TableDefs.Refresh
End Sub

Sub BuildTempSchema()
    Dim db As DAO.Database, tbl As DAO.TableDef, fld As DAO.Field
    Dim i As Long
    
    Set db = CurrentDb
    Set tbl = db.TableDefs("tempSchema")
    
    'build temp tbl fields
    For i = 1 To 8
        Set fld = tbl.CreateField("ColumnStr" & i, dbText, 255)
        tbl.Fields.Append fld

        Set fld = tbl.CreateField("ColumnType" & i, dbText, 255)
        tbl.Fields.Append fld
    Next
    
    db.TableDefs.Refresh
End Sub

Function BuildAddStr(colMax As Long) As String
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim strSQL As String, returnStr As String
    Dim itemStr As String, colStr As String, rowStr As String
    Dim recordCount As Long, i As Long, c As Long
    
    Set db = CurrentDb
    
    strSQL = "SELECT * FROM [tempSchema]"
    
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    If rs.EOF Then Exit Function
    
    rs.MoveLast
    rs.MoveFirst
    
    returnStr = ""
    For i = 1 To rs.recordCount
        rowStr = ""
        For c = 1 To colMax
            colStr = "ColumnStr" & c
            'Debug.Print "COL STR: " & colStr
            itemStr = Nz(rs.Fields(colStr), "")
            If Trim(itemStr) <> "" Then
                rowStr = rowStr & itemStr & "!!"
            End If
        Next
        If rowStr <> "" Then
            rowStr = Trim(Left(rowStr, Len(rowStr) - 2))
            returnStr = returnStr & rowStr & "+++"
        End If
        
        rs.MoveNext
     Next
     
     If returnStr = "" Then Exit Function
     
     BuildAddStr = Trim(Left(returnStr, Len(returnStr) - 3))
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
        Debug.Print "KNOWN TARGETS STR SQL: " & strSQL

    Case Else
        strSQL = "SELECT COUNT(*) AS Count FROM [" & tblName & "]"
        
    End Select
    
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    GetLocalTableCount = rs!Count
End Function


'delimited string as input
Sub UpdateSelectorsTblTargetId(str As String, targetId As String)
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim inputArr() As String, strSQL As String, targetIdStr As String, selectorStr As String, currentTargetId As String
    Dim i As Long
    
'    Debug.Print "!!!!!!" & vbLf & "UPDATE SELECTORS TBL TARGET ID:" & vbLf
'    Debug.Print "INPUT STR: " & str
'    Debug.Print "TARGET ID STR: " & targetId
    
    inputArr = Split(Trim(str), "!!")
    targetIdStr = Trim(targetId)
    
    If targetIdStr = "" Or UBound(inputArr) = -1 Then Exit Sub
    
    Set db = CurrentDb
    
    For i = LBound(inputArr) To UBound(inputArr)
        selectorStr = Trim(inputArr(i))
        
        strSQL = "SELECT * FROM [localSelectors] WHERE [selector] = '" & selectorStr & "'"
        
        Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
        If Not rs.EOF Then
            
            currentTargetId = Nz(rs!targetId, "")
            If currentTargetId <> targetIdStr Then
            
                'check if targetId WRONG / Diff
                'If currentTargetId <> "" And UCase(currentTargetId) <> "NULL" Then ThrowError 1963, "DIFFERENT TARGET ID; SELECTOR: " & selectorStr: Exit Sub
                
                'otherwise update it
                strSQL = "UPDATE [localSelectors] SET [targetId] = '" & targetIdStr & "' WHERE [selector] = '" & selectorStr & "'"
                db.Execute strSQL
            End If
        End If
    Next
                    
End Sub

Sub UpdateGWSearchTargetId(str As String, targetId As String)
    Dim db As DAO.Database, rs As DAO.Recordset, rsTarget As DAO.Recordset
    Dim inputArr() As String, strSQL As String, targetIdStr As String, targetNameStr As String
    Dim selectorStr As String, currentTargetId As String
    Dim i As Long
    
    inputArr = Split(Trim(str), "!!")
    targetIdStr = Trim(targetId)
    
    If targetIdStr = "" Or UBound(inputArr) = -1 Then Exit Sub
    
    Set db = CurrentDb
    
    For i = LBound(inputArr) To UBound(inputArr)
        selectorStr = Trim(inputArr(i))
        
        strSQL = "SELECT * FROM [tempGWSearchResults] WHERE [selector] = '" & selectorStr & "'"
        
        Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
        If Not rs.EOF Then
            
            currentTargetId = Nz(rs!targetId, "")
            If currentTargetId <> targetIdStr Then
            
                'check if targetId WRONG / Diff
                'If currentTargetId <> "" And UCase(currentTargetId) <> "NULL" Then ThrowError 1963, "DIFFERENT TARGET ID; SELECTOR: " & selectorStr: Exit Sub
                
                strSQL = "SELECT * FROM [localTargets] WHERE [targetId] = '" & targetIdStr & "'"
                Set rsTarget = db.OpenRecordset(strSQL, dbOpenSnapshot)
                targetNameStr = Nz(rsTarget!targetName, "")
'                Debug.Print "^^^^^^"
'                Debug.Print "TARGET NAME STR: " & targetNameStr
'                Debug.Print "^^^^^^"
'
                'otherwise update it
                strSQL = "UPDATE [tempGWSearchResults] SET [targetId] = '" & targetIdStr & "', " & _
                "[targetName] = '" & targetNameStr & "' WHERE [selector] = '" & selectorStr & "'"

                
                db.Execute strSQL
            End If
        End If
    Next

End Sub

Sub UpdateTotalSelectors()
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim strSQL As String, targetIdStr As String
    Dim i As Long
    
    Set db = CurrentDb
    
    'query lang per claude
    strSQL = "SELECT DISTINCT [targetId] FROM [localSelectors] WHERE [targetId] IS NOT NULL AND [targetId] <> ''"
    
    Set rs = db.OpenRecordset(strSQL)
    
    If rs.EOF Then Exit Sub
    
    rs.MoveLast
    rs.MoveFirst
    
    For i = 1 To rs.recordCount
        targetIdStr = Nz(rs!targetId, "")
        'Debug.Print "TARGET ID STR: " & targetIdStr
        If Trim(targetIdStr) <> "" Then
            UpdateTargetsSelectorCountTbl targetIdStr
        End If
        rs.MoveNext
    Next
End Sub

Sub UpdateTargetsSelectorCountTbl(targetId As String)
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim strSQL As String, targetIdStr As String
    Dim x As Long
    
    targetIdStr = Trim(targetId)
    
    'empty input
    If targetIdStr = "" Then ThrowError 1962, "UPDATE SELECTOR COUNT BLANK TARGET ID; INPUT: " & targetId: Exit Sub
    
    Set db = CurrentDb
    
    strSQL = "SELECT * FROM [localSelectors] WHERE [targetId] = '" & targetIdStr & "'"
    'Debug.Print "SOURCE STR SQL: " & strSQL
    
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
        x = rs.recordCount
    Else
        x = 0
    End If
    rs.Close
    
    'Debug.Print "SELECTOR COUNT X: " & x
    
    strSQL = "UPDATE [localTargets] SET [selectorCount] = " & x & " WHERE [targetId] = '" & targetIdStr & "'"
    
    db.Execute strSQL
End Sub

Sub UpdateAddToGW(str As String)
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim inputStr As String, strSQL As String, dateStr As String, userStr As String
    
    inputStr = Trim(str)
    If inputStr = "" Then Exit Sub
    
    Set db = CurrentDb
    
    dateStr = Now()
    userStr = Trim(LCase(Environ("USERNAME")))
    
    strSQL = "UPDATE [tempGWSearchResults] SET [addedToGrayWolfe] = 'YES', " & _
    "[lastUpdated] = '" & dateStr & "', " & _
    "[lastUpdatedBy] = '" & userStr & "' " & _
    "WHERE [selector] = '" & inputStr & "'"
    
    'Debug.Print "^^^^^^^^^^^^^"
    'Debug.Print "UPDATE ADD TO GW strSQL: " & strSQL
    
    db.Execute strSQL
End Sub

Sub UpdateLastUpdated(str As String, tbl As String)
    Dim db As DAO.Database, strSQL As String
    Dim targetStr As String, tblStr As String, dateStr As String, userStr As String
    
    targetStr = Trim(str)
    tblStr = Trim(tbl)
    If tblStr = "" Or targetStr = "" Then Exit Sub
    
    dateStr = Now()
    userStr = Trim(LCase(Environ("USERNAME")))
    
    Set db = CurrentDb
    
    strSQL = "UPDATE [" & tblStr & "] SET [lastUpdated] = '" & dateStr & "', " & _
    "[lastUpdatedBy] = '" & userStr & "' " & _
    "WHERE [targetId] = '" & targetStr & "'"
    
    'Debug.Print "^^^^" & vbLf & "LAST UPDATE SQL STR: " & strSQL
    
    db.Execute strSQL
End Sub

Sub UpdateTargetStatsForm(targetId As String, inputItem As String, fieldItem As String)
    Dim db As DAO.Database, strSQL As String
    Dim targetIdStr As String, inputStr As String, fieldStr As String
    
    targetIdStr = Trim(targetId)
    inputStr = Trim(inputItem)
    fieldStr = Trim(fieldItem)
    
    If targetId = "" Or fieldStr = "" Then Exit Sub
    
    Set db = CurrentDb
    
    strSQL = "UPDATE [localTargets] SET [" & fieldStr & "] = '" & inputStr & "' WHERE [targetId] = '" & targetIdStr & "'"
    
    'Debug.Print "UPDATE TARGET STATS STR SQL: " & strSQL
    
    db.Execute strSQL
    
    

End Sub

'for Form UI
Sub UpdateTargetSelectors(targetId As String, inputItem As String, currentItem As String, selectorType As String)
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
            Debug.Print "DELETING STR: " & itemStr
            
            DeleteTargetSelector targetIdStr, itemStr
        Next
    End If
    
    'add items
    addStr = DetectStrDiff(inputStr, currentStr, "!!")
    arr = Split(addStr, "!!")
    
    If UBound(arr) > -1 Then
    
        For i = LBound(arr) To UBound(arr)
            itemStr = arr(i)
            Debug.Print "ADDING STR: " & itemStr
            
            FillLocalSelectors itemStr, selectorTypeStr, targetIdStr
        Next
    End If

End Sub

Sub RequeryTargetForm()
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim strSQL As String
    
    Set db = CurrentDb
    
    strSQL = "SELECT * FROM localTargets"
    
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    Set Form_frmTargetDetails.subSelectors.Form.Recordset = rs
    
End Sub



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
    Dim db As DAO.Database
    Dim strSQL As String, inputStr As String, targetIdStr As String
    
    inputStr = Trim(inputItem)
    targetIdStr = Trim(targetId)
    
    If targetIdStr = "" Then Exit Sub
    
    Set db = CurrentDb
    
    strSQL = "DELETE FROM localSelectors WHERE [targetId] = '" & targetIdStr & "' AND [selector] = '" & inputStr & "'"
    'Debug.Print "DELETE TARGET SELECTOR SQL: " & strSQL
    
    db.Execute strSQL, dbFailOnError
    Set db = Nothing

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

''''''''''''''''''''''
