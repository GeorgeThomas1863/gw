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

    Select Case importType

    Case "unrelated"
        RunUnrelatedImport addStr, selectorType

    Case "default"
        RunDefaultImport addStr

    End Select

End Sub

'+++++++++++++++++++++++++++++++++++

Sub RunUnrelatedImport(str As String, selectorTypeInput As String)
    Dim addArr() As String, addStr As String, splitStr As String
    Dim selectorTypeStr As String, searchStr As String, SStr As String
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

    'populate input data
    FillTempSchema Trim(str), colMax

    openArgsStr = DetectSchemaTypeStr(colMax)

    'display results / button for triggering default import
    DoCmd.OpenForm "frmSchemaDetection", acNormal, , , , , openArgsStr

End Sub

'TRIGGERED BY BUTTON
Sub RunAddSchemaData(typeInput As String, colMax As Long)
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim typeArr() As String
    Dim rowStr As String, rowInputStr As String
    Dim searchInputStr As String, searchStr As String, SStr As String
    Dim targetIdsStr As String, targetIdStr As String
    Dim targetArr() As String
    Dim value As String, typeItem As String, result As String
    Dim rowCount As Long, tblDiff As Long
    Dim r As Long, c As Long, tIdx As Long

    '1. Split confirmed types from schema form
    typeArr = Split(Trim(typeInput), "!!")
    If UBound(typeArr) < 0 Then ThrowError 1962, "TYPE ARR EMPTY IN RUN ADD SCHEMA DATA": Exit Sub

    '2. Load tempSchema recordset (For loop pattern)
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT " & BuildColumnSelectStr(colMax) & " FROM [tempSchema] ORDER BY [ID]", dbOpenSnapshot)

    If rs.EOF Then ThrowError 1962, "TEMP SCHEMA EMPTY IN RUN ADD SCHEMA DATA": rs.Close: Exit Sub

    rs.MoveLast
    rs.MoveFirst
    rowCount = rs.RecordCount

    '3. Pre-pass: build searchInputStr in same format as former BuildAddStr output
    '   (rows separated by +++, items by !!, no trailing separators)
    searchInputStr = ""
    For r = 1 To rowCount
        rowInputStr = ""
        For c = 1 To colMax
            value = Nz(rs.Fields("ColumnStr" & c).Value, "")
            If value <> "" And LCase(value) <> "null" And LCase(typeArr(c - 1)) <> "null" Then
                rowInputStr = rowInputStr & value & "!!"
            End If
        Next c
        If rowInputStr <> "" Then
            'remove trailing !!
            rowInputStr = Left(rowInputStr, Len(rowInputStr) - 2)
            searchInputStr = searchInputStr & rowInputStr & "+++"
        End If
        rs.MoveNext
    Next r
    'remove trailing +++
    If Len(searchInputStr) > 3 Then searchInputStr = Left(searchInputStr, Len(searchInputStr) - 3)

    rs.MoveFirst  'reset cursor to start before main loop

    searchStr = SearchAddStrTbl(searchInputStr)

    '4. Main loop — process each tempSchema row
    For r = 1 To rowCount

        'reset per-row state
        targetIdStr = ""
        rowStr = ""

        '-- Build selector list for this row (inlined AddRelatedRowSelectors) --
        For c = 1 To colMax
            value = Nz(rs.Fields("ColumnStr" & c).Value, "")
            typeItem = Trim(typeArr(c - 1))
            If value <> "" And LCase(value) <> "null" And typeItem <> "" And LCase(typeItem) <> "null" Then
                rowStr = rowStr & value & "!!"
                result = FillLocalSelectors(value, typeItem, "", BuildDataSource("Default Import"))
                If InStr(UCase(result), "SELECTOR ALREADY KNOWN") = 0 Then
                    UpdateAddToGW value
                End If
            End If
        Next c
        'remove trailing !!
        If Len(rowStr) > 2 Then rowStr = Left(rowStr, Len(rowStr) - 2)

        '-- Target linking (multi-column only) --
        If colMax >= 2 And rowStr <> "" Then
            targetIdsStr = CollectTargetIdsForRow(rowStr)

            If targetIdsStr = "" Then
                '0 existing targets — create new
                targetIdStr = FillLocalTargets()
                UpdateSelectorsTblTargetId rowStr, targetIdStr
                UpdateGWSearchTargetId rowStr, targetIdStr
                UpdateTargetsSelectorCountTbl targetIdStr

            Else
                targetArr = Split(targetIdsStr, "!!")

                If UBound(targetArr) = 0 Then
                    '1 existing target — join it
                    targetIdStr = Trim(targetArr(0))
                    UpdateSelectorsTblTargetId rowStr, targetIdStr
                    UpdateGWSearchTargetId rowStr, targetIdStr
                    UpdateTargetsSelectorCountTbl targetIdStr

                Else
                    '2+ existing targets — bridging
                    targetIdStr = Trim(targetArr(0))
                    UpdateSelectorsTblTargetId rowStr, targetIdStr
                    For tIdx = 0 To UBound(targetArr)
                        For c = 1 To colMax
                            value = Nz(rs.Fields("ColumnStr" & c).Value, "")
                            typeItem = Trim(typeArr(c - 1))
                            If value <> "" And LCase(value) <> "null" And typeItem <> "" And LCase(typeItem) <> "null" Then
                                FillLocalSelectors value, typeItem, Trim(targetArr(tIdx)), BuildDataSource("Bridge")
                            End If
                        Next c
                        UpdateTargetsSelectorCountTbl Trim(targetArr(tIdx))
                    Next tIdx
                    UpdateGWSearchTargetId rowStr, targetIdStr
                End If
            End If
        End If

        rs.MoveNext
    Next r

    rs.Close

    '5. SharePoint sync
    tblDiff = GetLocalTableCount("localSelectors") - Form_frmMainMenu.localSelectorsCount
    If tblDiff > 0 Then SendDataLocalSP "localSelectors", "Selectors", "localSelectors", tblDiff

    tblDiff = GetLocalTableCount("localTargets") - Form_frmMainMenu.localTargetsCount
    If tblDiff > 0 Then SendDataLocalSP "localTargets", "Targets", "localTargets", tblDiff

    SendDataUpdateSP "localSelectors", "Selectors", "selectorId", "targetId"

    '6. Optional S search
    If Trim(LCase(Form_frmMainMenu.cboSAdd)) = "yes" Then
        SStr = SearchS(searchInputStr)
    Else
        SStr = "!!$$"
    End If

    '7. Open results
    DoCmd.OpenForm "frmResultsDisplay", acNormal, , , , , searchInputStr & "$$" & searchStr & "$$" & SStr

End Sub

'builds SELECT column list for tempSchema query
Private Function BuildColumnSelectStr(colMax As Long) As String
    Dim colStr As String
    Dim i As Long

    colStr = ""
    For i = 1 To colMax
        colStr = colStr & "[ColumnStr" & i & "], "
    Next

    'remove trailing comma and space
    If Len(colStr) > 2 Then colStr = Left(colStr, Len(colStr) - 2)

    BuildColumnSelectStr = colStr
End Function
