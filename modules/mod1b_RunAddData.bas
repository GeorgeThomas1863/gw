*****************************************

MOD 1b_RunAddData

*****************************************


Option Compare Database

Option Explicit

Sub RunAddData(str As String, importTypeInput As String, selectorTypeInput As String, delimInput As String)
    Dim addArr() As String, addStr As String, importType As String
    Dim selectorType As String, delimStr As String
    
    'throws error if delim not set
    delimStr = CleanDelimInput(Trim(str), Trim(delimInput))
    
    'blank input error
    If Trim(str) = "" Or InStr(Trim(str), "Input / Paste ") = 1 Then ThrowError 1953, str: Exit Sub

    'standaredizes delims / nulls
    addStr = CleanUserInput(Trim(str), delimStr)
    importType = CleanImportTypeInput(Trim(importTypeInput))
    
    'selector type calculated later for default import
    If importType = "unrelated" Then
        selectorType = CleanSelectorInput(Trim(selectorTypeInput))
    End If
    
'    Debug.Print "++++++++++++"
'    Debug.Print "ADD DATA INPUT " & addStr
'    Debug.Print "IMPORT TYPE " & importType
'    Debug.Print "SELECTOR TYPE " & selectorType
'    Debug.Print "DELIM STR " & delimStr
    
    Select Case importType

    Case "unrelated"
        RunUnrelatedImport addStr, selectorType
    
    Case "default"
        RunDefaultImport addStr

    End Select

End Sub

'+++++++++++++++++++++++++++++++++++

Sub RunUnrelatedImport(str As String, selectorTypeInput As String)
    Dim addArr() As String, searchArr() As String, addStr As String, splitStr As String
    Dim selectorTypeStr As String, searchStr As String, alertStr As String, SStr As String
    Dim res As String, openArgsStr As String
    Dim tblDiff As Long, i As Long, x As Long
    
    'returns gw search
    searchStr = SearchAddStrTbl(Trim(str))
        
    'treat everything as separate (split on row AND item)
    splitStr = Replace(Trim(str), "+++", "!!")
    addArr = Split(splitStr, "!!")
    
    x = 0
    For i = LBound(addArr) To UBound(addArr)
        addStr = Trim(addArr(i))
        If addStr <> "" And LCase(addStr) <> "null" Then
            'resets / recalcs each time
            selectorTypeStr = DetectSelectorType(addStr, Trim(selectorTypeInput))
            
            'handle wrong / null types
            If Trim(UCase(selectorTypeStr)) = "WRONG" Or Trim(UCase(selectorTypeStr)) = "NULL" Then
                selectorTypeStr = FixWrongSelectorType(addStr, selectorTypeInput, selectorTypeStr)
            End If
            
            If selectorTypeStr <> "SKIP" Then

                'Send to local here, count only hits
                res = FillLocalSelectors(addStr, selectorTypeStr)
                If InStr(UCase(res), "SELECTOR ALREADY KNOWN") = 0 And Trim(res) <> "" Then
                    UpdateAddToGW addStr
                    x = x + 1
                End If
            End If
        End If
    Next
    
    'upload everything to SP at once, upload selectors
    tblDiff = GetLocalTableCount("localSelectors") - Form_frmMainMenu.localSelectorsCount
    If tblDiff > 0 Then SendDataLocalSP "localSelectors", "Selectors", "localSelectors", tblDiff
    
    'check if search S set, and search if it is
    If Trim(LCase(Form_frmMainMenu.cboSAdd)) = "yes" Then
        SStr = SearchS(Trim(str))
    Else
        SStr = "!!$$"
    End If

    openArgsStr = Trim(str) & "$$" & searchStr & "$$" & SStr

    'display results
    DoCmd.OpenForm "frmResultsDisplay", acNormal, , , , , openArgsStr
    
End Sub

'+++++++++++++++++++++++++++

'builds schema, puts data in form; next step triggered by submit button in frmSchemaDetection
Sub RunDefaultImport(str As String)
    Dim arr() As String, schemaTypeStr As String
    Dim openArgsStr As String, colMax As Long

    'check for null input
    arr = Split(Trim(str), "+++")
    If UBound(arr) = -1 Then ThrowError 1962, "EMPTY SPLIT ARRAY FROM STR: " & str: Exit Sub
    
    colMax = DetectColMax(Trim(str))
    'Debug.Print "COLMAX: " & colMax
    
    'populate input data
    FillTempSchema Trim(str), colMax

    openArgsStr = DetectSchemaTypeStr(colMax)
    
    'display results / button for triggering default import
    DoCmd.OpenForm "frmSchemaDetection", acNormal, , , , , openArgsStr
    
    'set dataStr in schema form
    'Form_frmSchemaDetection.dataStr = str
    
End Sub

'TRIGGERED BY BUTTON
Sub RunAddSchemaData(typeInput As String, colMax As Long)
    Dim addArr() As String, searchArr() As String, addStr As String, searchStr As String
    Dim typeStr As String, rowStr As String, addSelectorsStr As String, addTargetsStr As String
    Dim targetReturnStr As String, alertStr As String, SStr As String, openArgsStr As String
    Dim tblDiff As Long, i As Long
       
    addStr = BuildAddStr(colMax)
    If addStr = "" Then ThrowError 1962, "ADD STR EMPTY IN RUN ADD SCHEMA DATA": Exit Sub
    
    'Debug.Print "COLMAX: " & colMax
    
    typeStr = Trim(typeInput)
    
    'gw selector hits
    searchStr = SearchAddStrTbl(addStr)
    'Debug.Print "^^^^" & vbLf & "SEARCH STR RETURN: " & searchStr
    
    'split on row
    addArr = Split(Trim(addStr), "+++")
    
    targetReturnStr = ""
    For i = LBound(addArr) To UBound(addArr)
        rowStr = Trim(addArr(i))
        addTargetsStr = "NULL"
        
        If rowStr <> "" Then
            addSelectorsStr = AddRelatedRowSelectors(rowStr, typeStr)
            addTargetsStr = AddRelatedRowTargets(rowStr, typeStr, addSelectorsStr)
            targetReturnStr = targetReturnStr & addTargetsStr & "!!"
        End If
    Next
    
    UpdateAddDataCounts addStr, targetReturnStr
    
    tblDiff = GetLocalTableCount("localSelectors") - Form_frmMainMenu.localSelectorsCount
    If tblDiff > 0 Then SendDataLocalSP "localSelectors", "Selectors", "localSelectors", tblDiff
    
    tblDiff = GetLocalTableCount("localTargets") - Form_frmMainMenu.localTargetsCount
    If tblDiff > 0 Then SendDataLocalSP "localTargets", "Targets", "localTargets", tblDiff
    
    tblDiff = GetLocalTableCount("selectorsTargetId") - Form_frmMainMenu.localSelectorsTargetIdCount
    If tblDiff > 0 Then SendDataUpdateSP "localSelectors", "Selectors", "selectorId", "targetId"

    'check if search S set, and search if it is
    If Trim(LCase(Form_frmMainMenu.cboSAdd)) = "yes" Then
        SStr = SearchS(Trim(addStr))
    Else
        SStr = "!!$$"
    End If
    
    openArgsStr = addStr & "$$" & searchStr & "$$" & SStr

    'display results
    DoCmd.OpenForm "frmResultsDisplay", acNormal, , , , , openArgsStr

        
End Sub

Function AddRelatedRowSelectors(inputStr As String, typeStr As String) As String
    Dim inputArr() As String, typeArr() As String, inputItem As String, typeItem As String
    Dim returnStr As String, addStr As String
    Dim i As Long
    
    'create arrs
    inputArr = Split(inputStr, "!!")
    typeArr = Split(typeStr, "!!")
    
    'Debug.Print "UBOUND INPUT ARR: " & UBound(inputArr)
    'Debug.Print "UBOUND TYPE ARR: " & UBound(typeArr)
    
    'check if inputs are fucked
    If UBound(inputArr) <> UBound(typeArr) Or UBound(inputArr) = -1 Or UBound(typeArr) = -1 Then _
    ThrowError 1962, "ARRAYS FOR RELATED IMPORT WRONG; inputStr: " & inputStr & " typeStr: " & typeStr: Exit Function
    
    returnStr = ""
    For i = LBound(inputArr) To UBound(inputArr)
        inputItem = Trim(inputArr(i))
        typeItem = Trim(typeArr(i))
        If inputItem <> "" And LCase(inputItem) <> "null" Then
            addStr = FillLocalSelectors(inputItem, typeItem, "")
            'Debug.Print "ADD STR: " & addStr
        
            'NOT in GW (selector NOT known)
            If InStr(UCase(addStr), "SELECTOR ALREADY KNOWN") = 0 Then
                returnStr = returnStr & addStr & "!!"
                UpdateAddToGW inputItem
            End If
        End If
    Next
    
    If Trim(returnStr) = "" Then AddRelatedRowSelectors = "": Exit Function
    
    'trailing delim
    AddRelatedRowSelectors = Trim(Left(returnStr, Len(returnStr) - 2))
End Function

'inputStr is user input, addSelectorStr is new selectors
Function AddRelatedRowTargets(inputStr As String, typeStr As String, addSelectorsStr As String) As String
    Dim inputArr() As String, targetArr() As String, selectorArr() As String, typeArr() As String
    Dim targetIdsStr As String, targetIdStr As String, tId As String, dupResult As String
    Dim i As Long, j As Long

    inputArr = Split(inputStr, "!!")

    'single entries
    If UBound(inputArr) < 1 Then Exit Function

    'collect ALL existing targetIds for selectors in this row
    targetIdsStr = CollectTargetIdsForRow(inputStr)

    If targetIdsStr = "" Then
        '0 targets found - create new
        targetIdStr = FillLocalTargets()

    Else
        targetArr = Split(targetIdsStr, "!!")

        If UBound(targetArr) = 0 Then
            '1 target found - use it
            targetIdStr = Trim(targetArr(0))

        Else
            '2+ targets found - add all selectors to ALL targets (no merge)
            targetIdStr = Trim(targetArr(0))

            'fill blank targetIds with first target (for new selectors from AddRelatedRowSelectors)
            UpdateSelectorsTblTargetId inputStr, targetIdStr

            'ensure every selector exists under every matched target
            selectorArr = Split(inputStr, "!!")
            typeArr = Split(typeStr, "!!")

            For i = 0 To UBound(targetArr)
                tId = Trim(targetArr(i))
                For j = 0 To UBound(selectorArr)
                    If Trim(selectorArr(j)) <> "" And LCase(Trim(selectorArr(j))) <> "null" Then
                        dupResult = FillLocalSelectors(Trim(selectorArr(j)), Trim(typeArr(j)), tId)
                    End If
                Next j
                UpdateTargetsSelectorCountTbl tId
            Next i

            UpdateGWSearchTargetId inputStr, targetIdStr
            AddRelatedRowTargets = targetIdStr
            Exit Function
        End If
    End If

    UpdateSelectorsTblTargetId inputStr, targetIdStr
    UpdateGWSearchTargetId inputStr, targetIdStr

    AddRelatedRowTargets = targetIdStr
End Function

Sub UpdateAddDataCounts(addStr As String, targetId As String)
    Dim inputArr() As String, rowArr() As String, targetArr() As String
    Dim rowStr As String, inputItem As String, targetIdStr As String
    Dim i As Long, j As Long
    
    inputArr = Split(addStr, "+++")
    targetArr = Split(targetId, "!!")
    
    
    For i = LBound(inputArr) To UBound(inputArr)
        rowStr = inputArr(i)
        
        If Trim(rowStr) <> "" Then
            rowArr = Split(rowStr, "!!")
            targetIdStr = targetArr(i) 'same target for each row

            'update selectors count
            UpdateTargetsSelectorCountTbl targetIdStr
        End If
     Next
End Sub
