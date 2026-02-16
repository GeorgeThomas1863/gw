*****************************************

Form_frmResultsDisplay

*****************************************

Option Explicit

Option Compare Database

Public SFilterStr As String
Public attachmentFilter As Boolean
Public fd302Filter As Boolean
Public lastYearFilter As Boolean
Public authoredFilter As Boolean

Public gwFilterStr As String
Public inSFilter As Boolean
Public notInSFilter As Boolean
Public gwTargetFilter As Boolean
Public gwSelectorFilter As Boolean
Public notInGWFilter As Boolean

'+++++++++

'ADD GW FILTERS

Private Sub Form_Load()
    Dim arr() As String, searchArr() As String, SArr() As String, SFilterArr() As String
    Dim targetArr() As String, splitStr As String, SStr As String, gwStr As String
    Dim filterStr As String, filterItem As String, targetStr As String, targetItem As String, targetFilterStr As String
    Dim i As Long, x As Long
    
    '''''''
    'OPEN ARGS:
    'itemsSearched $$ GW target hits !! GW selector hits $$ S hits !! in S $$ SFilter
    ''''''
    
    Me.tabResults.Value = 1
    ClearFilterS
    ClearFilterGW

    If IsNull(Me.OpenArgs) Then Exit Sub
    
    arr = Split(Me.OpenArgs, "$$")
    
    If UBound(arr) < 3 Then ThrowError 1971, "OPEN ARGS: " & Me.OpenArgs: Exit Sub
    
    PopulateStatsSearched arr(0)
    PopulateStatsGW arr(1)
    PopulateStatsS arr(2), arr(3)

End Sub

Private Sub PopulateStatsSearched(str As String)
    Dim arr() As String, splitStr As String, itemStr As String
    Dim x As Long, i As Long

    splitStr = Replace(Trim(str), "+++", "!!")
    arr = Split(splitStr, "!!")
    If UBound(arr) = -1 Then Exit Sub
    
    'count non null / blanks
    x = 0
    For i = LBound(arr) To UBound(arr)
        itemStr = Trim(arr(i))
        If itemStr <> "" And LCase(itemStr) <> "null" Then
            x = x + 1
        End If
    Next
    
    'number searched
    Me.lblSearchedCount.Caption = Trim(x)
End Sub

Private Sub PopulateStatsGW(str As String)
    Dim arr() As String, inputStr As String, searchedStr As String
    Dim targetStr As String
    
    inputStr = Trim(str)
    searchedStr = Me.lblSearchedCount.Caption
    
    'GW NOT searched
    If LCase(inputStr) = "null" Or inputStr = "" Or inputStr = "!!" Then
        Me.lblGWTargetsCount.Caption = "N/A"
        Me.lblInGWCount.Caption = "N/A"
        Me.lblNotInGWCount.Caption = "N/A"
        
        'open S results first
        Me.tabResults.Value = 0
        Exit Sub
    End If
    
    arr = Split(inputStr, "!!")
    
    'fill selector counts
    Me.lblGWTargetsCount.Caption = arr(0)
    Me.lblInGWCount.Caption = arr(1)
    Me.lblNotInGWCount.Caption = Trim(Val(searchedStr) - Val(arr(1)))
    
    'handle GW targets
    targetStr = SearchAllTargetHitsGWTbl()
    'Debug.Print "DISPLAY TARGET STR: " & targetStr
    
    PopulateTargetFilter targetStr

End Sub

Private Sub PopulateTargetFilter(str)
    Dim arr() As String, targetStr As String
    Dim targetItem As String, targetFilterStr As String
    Dim i As Long
    
    targetStr = Trim(str)
    
    'no targets
    If targetStr = "" Then
        targetFilterStr = """[No Known Targets Found]"""
        Me.cboGWHitsFilter.Visible = True
        Me.cboGWHitsFilter.RowSource = targetFilterStr
        Exit Sub
    End If
    
    arr = Split(targetStr, "!!")
    targetFilterStr = """[Display All GW Hits]"";"
    For i = LBound(arr) To UBound(arr)
        targetItem = Trim(arr(i))
        If targetItem <> "" And LCase(targetItem) <> "null" Then
            targetItem = """" & targetItem & """"
            targetFilterStr = targetFilterStr & targetItem & ";"
        End If
    Next
    
    If Trim(targetFilterStr) <> "" Then
        targetFilterStr = AlphabetizeStr(Trim(Left(targetFilterStr, Len(targetFilterStr) - 1)), ";")
    End If
    
    Me.cboGWHitsFilter.Visible = True
    Me.cboGWHitsFilter.RowSource = targetFilterStr

End Sub

Private Sub PopulateStatsS(str As String, filterStr As String)
    Dim arr() As String, inputStr As String, searchedStr As String
    
    inputStr = Trim(str)
    searchedStr = Me.lblSearchedCount.Caption
    
    'S not searched
    If inputStr = "" Or inputStr = "!!" Then
        Me.lblSHitsCount.Caption = "N/A"
        Me.lblInSCount.Caption = "N/A"
        Me.lblNotInSCount.Caption = "N/A"
        PopulateSFilter filterStr
        Exit Sub
    End If
    
    'S searched
    arr = Split(inputStr, "!!")
     
    Me.lblSHitsCount.Caption = Trim(arr(0))
    Me.lblInSCount.Caption = Trim(arr(1))
    Me.lblNotInSCount.Caption = Trim(Val(searchedStr) - Val(arr(1)))
    
    If Trim(filterStr) = "" Then Exit Sub
    
    PopulateSFilter filterStr
End Sub

Private Sub PopulateSFilter(str As String)
    Dim arr() As String, inputStr As String, filterStr As String, filterItem As String
    Dim i As Long
    
    inputStr = Trim(str)
    
    If inputStr = "" Then
        filterStr = """[No S Hits]"""
        Me.cboSHitsFilter.Visible = True
        Me.cboSHitsFilter.RowSource = filterStr
        Exit Sub
    End If
    
    arr = Split(inputStr, "!!")

    filterStr = """[Display All S Hits]"";"
    For i = LBound(arr) To UBound(arr)
        filterItem = Trim(arr(i))
        If filterItem <> "" And LCase(filterItem) <> "null" Then
            'handle commas
            filterItem = """" & filterItem & """"
            filterStr = filterStr & filterItem & ";"
        End If
    Next

    'Remove trailing semicolon
    filterStr = Trim(Left(filterStr, Len(filterStr) - 1))
    
    Me.cboSHitsFilter.Visible = True
    Me.cboSHitsFilter.RowSource = filterStr
End Sub

'++++++++++++++++++++

'S FILTERS
Private Sub cboSHitsFilter_AfterUpdate()
    
    If IsNull(Me.cboSHitsFilter) Then Exit Sub
    
    SFilterStr = Me.cboSHitsFilter
    ApplyFilterS
End Sub

'Buttons
Private Sub btnNoAttachments_Click()
    attachmentFilter = Not attachmentFilter
    ApplyFilterS
    
    If attachmentFilter Then AddBtnFilterStyles "btnNoAttachments": Exit Sub
    RemoveBtnFilterStyles "btnNoAttachments"
End Sub

Private Sub btn302_Click()
    fd302Filter = Not fd302Filter
    ApplyFilterS
    
    If fd302Filter Then AddBtnFilterStyles "btn302": Exit Sub
    RemoveBtnFilterStyles "btn302"
End Sub

Private Sub btnLastYear_Click()
    lastYearFilter = Not lastYearFilter
    ApplyFilterS
    
    If lastYearFilter Then AddBtnFilterStyles "btnLastYear": Exit Sub
    RemoveBtnFilterStyles "btnLastYear"
End Sub

Private Sub btnAuthoredDocs_Click()
    authoredFilter = Not authoredFilter
    ApplyFilterS
    
    If authoredFilter Then AddBtnFilterStyles "btnAuthoredDocs": Exit Sub
    RemoveBtnFilterStyles "btnAuthoredDocs"
End Sub

Private Sub btnClearSearchFilters_Click()
    ClearFilterS "searchClick"
End Sub

'----------------

Public Sub ApplyFilterS()
    Dim arr() As String
    Dim filterStr As String, filterItem As String, returnStr As String
    Dim pastDate As Date, userStr As String
    Dim i As Long
    
    filterStr = ""
    
    'drop down
    If Trim(LCase(SFilterStr)) <> "" And InStr(LCase(SFilterStr), "[display all ") = 0 Then
        filterItem = "[selector] = '" & SFilterStr & "'"
        filterStr = filterStr & filterItem & " AND "
    End If
    
    If attachmentFilter Then
        filterItem = "[docSubType] <> 'Attachment'"
        filterStr = filterStr & filterItem & " AND "
    End If
    
    If fd302Filter Then
        filterItem = "([docSubType] = 'FD302' OR [docSubType] = 'FD1057')"
        filterStr = filterStr & filterItem & " AND "
    End If
    
    If lastYearFilter Then
        pastDate = DateAdd("yyyy", -1, Date)
        
        filterItem = "[createdDate] >= #" & Format(pastDate, "yyyy-mm-dd") & "#"
        filterStr = filterStr & filterItem & " AND "
    End If
    
    If authoredFilter Then
        userStr = Trim(LCase(Environ("USERNAME")))
        
        filterItem = "[author] = '" & userStr & "'"
        filterStr = filterStr & filterItem & " AND "
    End If
    
    If Trim(filterStr) = "" Or Trim(filterStr) = "AND" Then
        Me.subSResults.Form.Filter = "'*'"
        Me.subSResults.Form.FilterOn = False
        Me.subSResults.SetFocus
        Exit Sub
    End If

    'remove trailing
    filterStr = Trim(Left(filterStr, Len(filterStr) - 5))

    Me.subSResults.Form.Filter = filterStr
    Me.subSResults.Form.FilterOn = True
    Me.subSResults.SetFocus
End Sub

'reset everything
Public Sub ClearFilterS(Optional clearType As String = "")
    SFilterStr = "[Display All S Hits]"
    Me.cboSHitsFilter = "[Display All S Hits]"
    
    attachmentFilter = False
    fd302Filter = False
    lastYearFilter = False
    authoredFilter = False
    
    RemoveBtnFilterStyles "btnNoAttachments"
    RemoveBtnFilterStyles "btn302"
    RemoveBtnFilterStyles "btnLastYear"
    RemoveBtnFilterStyles "btnAuthoredDocs"
    
    Me.subSResults.Form.Filter = "'*'"
    Me.subSResults.Form.FilterOn = False
    
    If clearType = "searchClick" Then Me.subSResults.SetFocus
End Sub

'++++++++++++++++

'GW Filter
Private Sub cboGWHitsFilter_AfterUpdate()
    
    If IsNull(Me.cboGWHitsFilter) Then Exit Sub
    
    gwFilterStr = Me.cboGWHitsFilter
    ApplyFilterGW
End Sub

'Buttons
Private Sub btnInS_Click()
    inSFilter = Not inSFilter
    ApplyFilterGW
    
    If inSFilter Then AddBtnFilterStyles "btnInS": Exit Sub
    RemoveBtnFilterStyles "btnInS"
End Sub

Private Sub btnNotInS_Click()
    notInSFilter = Not notInSFilter
    ApplyFilterGW
    
    If notInSFilter Then AddBtnFilterStyles "btnNotInS": Exit Sub
    RemoveBtnFilterStyles "btnNotInS"
End Sub

Private Sub btnGWTarget_Click()
    gwTargetFilter = Not gwTargetFilter
    ApplyFilterGW
    
    If gwTargetFilter Then AddBtnFilterStyles "btnGWTarget": Exit Sub
    RemoveBtnFilterStyles "btnGWTarget"
End Sub

Private Sub btnInGW_Click()
    gwSelectorFilter = Not gwSelectorFilter
    ApplyFilterGW
    
    If gwSelectorFilter Then AddBtnFilterStyles "btnInGW": Exit Sub
    RemoveBtnFilterStyles "btnInGW"
End Sub

Private Sub btnNotInGW_Click()
    notInGWFilter = Not notInGWFilter
    ApplyFilterGW
    
    If notInGWFilter Then AddBtnFilterStyles "btnNotInGW": Exit Sub
    RemoveBtnFilterStyles "btnNotInGW"
End Sub

Private Sub btnClearGWFilters_Click()
    ClearFilterGW "gwClick"
End Sub

'----------------

Public Sub ApplyFilterGW()
    Dim arr() As String
    Dim filterStr As String, filterItem As String, returnStr As String
    Dim pastDate As Date, userStr As String
    Dim i As Long
    
    filterStr = ""
    
    'drop down
    If Trim(LCase(gwFilterStr)) <> "" And InStr(LCase(gwFilterStr), "[display all ") = 0 Then
        filterItem = "([targetName] = '" & gwFilterStr & "' OR [targetId] = '" & gwFilterStr & "')"
        filterStr = filterStr & filterItem & " AND "
    End If
    
    If inSFilter Then
        filterItem = "[SHits] > 0"
        filterStr = filterStr & filterItem & " AND "
    End If
    
    If notInSFilter Then
        filterItem = "[SHits] = 0"
        filterStr = filterStr & filterItem & " AND "
    End If
    
    If gwTargetFilter Then
        filterItem = "([targetId] Is Not Null And [targetId] <> '')"
        filterStr = filterStr & filterItem & " AND "
    End If
    
    If gwSelectorFilter Then
        filterItem = "[inGrayWolfe] = 'YES'"
        filterStr = filterStr & filterItem & " AND "
    End If
    
    If notInGWFilter Then
        filterItem = "[inGrayWolfe] = 'NO'"
        filterStr = filterStr & filterItem & " AND "
    End If
    
    If Trim(filterStr) = "" Or Trim(filterStr) = "AND" Then
        Me.subGWResults.Form.Filter = "'*'"
        Me.subGWResults.Form.FilterOn = False
        Me.subGWResults.SetFocus
        Exit Sub
    End If

    'remove trailing
    filterStr = Trim(Left(filterStr, Len(filterStr) - 5))

    Me.subGWResults.Form.Filter = filterStr
    Me.subGWResults.Form.FilterOn = True
    Me.subGWResults.SetFocus
End Sub

'reset everything
Public Sub ClearFilterGW(Optional clearType As String = "")
    gwFilterStr = "[Display All GW Hits]"
    Me.cboGWHitsFilter = "[Display All GW Hits]"
    
    inSFilter = False
    notInSFilter = False
    gwTargetFilter = False
    gwSelectorFilter = False
    notInGWFilter = False
    
    RemoveBtnFilterStyles "btnInS"
    RemoveBtnFilterStyles "btnNotInS"
    RemoveBtnFilterStyles "btnGWTarget"
    RemoveBtnFilterStyles "btnInGW"
    RemoveBtnFilterStyles "btnNotInGW"
    
    Me.subGWResults.Form.Filter = "'*'"
    Me.subGWResults.Form.FilterOn = False
    
    If clearType = "gwClick" Then Me.subGWResults.SetFocus
End Sub




'---------

Public Sub AddBtnFilterStyles(str As String)
    Dim ctl As Control
    
    Set ctl = Me.Controls(Trim(str))
    
    ctl.BackColor = RGB(135, 206, 235)
    ctl.ForeColor = RGB(0, 0, 0)
    ctl.BorderColor = RGB(0, 191, 255)
    ctl.BorderWidth = 2
    ctl.FontBold = True
End Sub

Public Sub RemoveBtnFilterStyles(str As String)
    Dim ctl As Control
    
    Set ctl = Me.Controls(Trim(str))
    
    ctl.BackColor = RGB(240, 240, 240)
    ctl.ForeColor = RGB(64, 64, 64)
    ctl.BorderColor = RGB(172, 172, 172)
    ctl.BorderWidth = 1
    ctl.FontBold = False
End Sub

'+++++++++++++++++++++



Private Sub btnCloseSearch_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub btnCloseGW_Click()
    DoCmd.Close acForm, Me.Name
End Sub