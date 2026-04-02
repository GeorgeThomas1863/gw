*****************************************

Form_frmMainMenu

*****************************************

Option Compare Database
Option Explicit

Public localNorksCount As Long
Public localSelectorsCount As Long
Public localTargetsCount As Long
'Public localTargetsNameCount As Long 'WANT USER TO MANUALLY SET
Public localSelectorsTargetIdCount As Long

'++++++++++++++++++++++++++++

'SUBMIT SECTION

'Search Submit
Private Sub btnSearch_Click()
    ResetLocalTablesSP
    
    'JUST PASS IN AS RAW STRING AND DO ALL PROCESSING IN MOD
    RunSearch Trim(Me.txtSearchInput), Trim(Me.cboSearchTarget), Trim(Me.cboSearchDelimiter)
End Sub

'add data submit
Private Sub btnAdd_Click()
    ResetLocalTablesSP
    
    RunAddData Trim(Me.txtAddDataInput), Trim(Me.cboImportMode), Trim(Me.cboSelectorTypeAdd), Trim(Me.cboAddDataDelimiter)
End Sub

'edit targets Submit
Private Sub btnEditTargets_Click()
    'ResetLocalTablesSP

    'DoCmd.OpenForm "frmTargetDetails", acNormal, , , , , openArgsStr
    DoCmd.OpenForm "frmTargetDetails"
End Sub

'++++++++++++++++++++++++++

'RESPONSIVE SECTION

'on load
Private Sub Form_Load()
    
    Me.cboSearchDelimiter.RowSource = DefineFormDefaults("cboSQL")
    Me.cboAddDataDelimiter.RowSource = DefineFormDefaults("cboSQL")
    
    'SET SEARCH PLACEHOLDER
    Me.txtSearchInput = DefineFormDefaults("txtSearchInput")
    Me.txtSearchInput.ForeColor = &H666666
    Me.txtSearchInput.FontSize = 10
    Me.txtDisclaimerSearch = DefineFormDefaults("disclaimerSearch")
    
    'token, check to avoid deleting user input
    If Trim(Me.txtSToken) = "" Or IsNull(Trim(Me.txtSToken)) Then
        Me.txtSToken = DefineFormDefaults("txtToken")
    End If
    
    'token add
    If Trim(Me.txtSTokenAdd) = "" Or IsNull(Trim(Me.txtSTokenAdd)) Then
        Me.txtSTokenAdd = DefineFormDefaults("txtToken")
    End If
    
    'SET ADD DATA PLACEHOLDER
    Me.txtAddDataInput.TextFormat = acTextFormatHTMLRichText
    Me.txtAddDataInput = DefineFormDefaults("txtAddDataDefault")
    Me.txtAddDataInput.ForeColor = &H666666
    Me.txtAddDataInput.FontSize = 10
    Me.txtDisclaimerAdd = DefineFormDefaults("disclaimerAdd")
    
    

End Sub

'++++++++++++++++++++++
'txtbox search click
Private Sub txtSearchInput_Click()

    Me.lblSearchFontSize.Visible = True
    Me.cboSearchFontSize.Visible = True
    
    Me.lblSearchDelimiter.Visible = True
    Me.cboSearchDelimiter.Visible = True

    'clear ONLY placeholder
    If InStr(Me.txtSearchInput, "Search one or multiple") <> 1 Then Exit Sub
        
    Me.txtSearchInput = ""
    Me.txtSearchInput.ForeColor = vbBlack
    Me.txtSearchInput.FontSize = 12
    Me.txtSearchInput.SetFocus
End Sub

'txtbox add data click
Private Sub txtAddDataInput_Click()
    
    Me.lblAddDataFontSize.Visible = True
    Me.cboAddDataFontSize.Visible = True
    
    Me.lblAddDataDelimiter.Visible = True
    Me.cboAddDataDelimiter.Visible = True

    'clear ONLY placeholder
    If InStr(Me.txtAddDataInput, "Input / Paste ") <> 1 Then Exit Sub

    Me.txtAddDataInput = ""
    Me.txtAddDataInput.TextFormat = acTextFormatPlain
    Me.txtAddDataInput.ForeColor = vbBlack
    Me.txtAddDataInput.FontSize = 12
    Me.txtAddDataInput.SetFocus
End Sub

'btn clear search click
Private Sub btnClearSearch_Click()
 
    'set to default
    Me.txtSearchInput.Value = ""
    Me.txtSearchInput = DefineFormDefaults("txtSearchInput")
    Me.txtSearchInput.ForeColor = &H666666
    Me.txtSearchInput.FontSize = 10
    
    Me.lblSearchFontSize.Visible = False
    Me.cboSearchFontSize.Visible = False
    
    Me.lblSearchDelimiter.Visible = False
    Me.cboSearchDelimiter.Visible = False
    Me.cboSearchDelimiter.DefaultValue = """New Row"""
    Me.cboSearchFontSize = 12
    
End Sub

'btn clear add data click
Private Sub btnClearAdd_Click()
    Dim importType As String
    
    importType = Trim(LCase(Me.cboImportMode))
 
    'set to default
    Me.txtAddDataInput.Value = ""
    Me.txtAddDataInput.TextFormat = acTextFormatHTMLRichText
    Me.txtAddDataInput.ForeColor = &H666666
    Me.txtAddDataInput.FontSize = 10
    
    Me.lblAddDataFontSize.Visible = False
    Me.cboAddDataFontSize.Visible = False
    
    Me.lblAddDataDelimiter.Visible = False
    Me.cboAddDataDelimiter.Visible = False
    Me.cboAddDataDelimiter.DefaultValue = """New Row"""
    
    Me.cboAddDataFontSize = 12
    
    If InStr(importType, "default") > 0 Then Me.txtAddDataInput = DefineFormDefaults("txtAddDataDefault"): Exit Sub
    
    Me.txtAddDataInput = DefineFormDefaults("txtAddDataUnrelated")
End Sub

Private Sub txtSToken_Click()
    Dim defaultStr As String
    
    defaultStr = Trim(DefineFormDefaults("txtToken"))
    
    If Trim(Me.txtSToken) = defaultStr Then Me.txtSToken = ""
    Me.txtSToken.SetFocus

End Sub

Private Sub txtHelp_Click()
    Dim docURL As String
    
    'link to doc
    docURL = "https://adhoc.Fnet.F/sites/003089/GrayWolfeProd/Shared%20Documents/How_To_Get_YOUR_S_API_Token.docx?d=wae09888ab6bf485bbd678f90d60be2e2"
    
    Application.FollowHyperlink docURL
End Sub

Private Sub txtSTokenAdd_Click()
    Dim defaultStr As String
    
    defaultStr = Trim(DefineFormDefaults("txtToken"))
    
    If Trim(Me.txtSTokenAdd) = defaultStr Then Me.txtSTokenAdd = ""
    Me.txtSTokenAdd.SetFocus

End Sub

Private Sub txtHelpAdd_Click()
    Dim docURL As String
    
    'link to doc
    docURL = "https://adhoc.Fnet.F/sites/003089/GrayWolfeProd/Shared%20Documents/How_To_Get_YOUR_S_API_Token.docx?d=wae09888ab6bf485bbd678f90d60be2e2"
    
    Application.FollowHyperlink docURL
End Sub

'++++++++++++++++++++++

'change header color
Private Sub tabMain_Change()
    
    Select Case Me.tabMain.Value
    
    Case 0
        Me.FormHeader.BackColor = &H9A572B
        cboSearchTarget.SetFocus
        
        
    Case 1
        Me.FormHeader.BackColor = &H3B5E35
        cboImportMode.SetFocus
    
    End Select

End Sub

Private Sub cboSearchFontSize_AfterUpdate()

    'fail conditions
    If IsNull(Me.cboSearchFontSize) Or IsAllNumbers(Me.cboSearchFontSize) = False Or Trim(Me.txtSearchInput.Value) = "" Or _
    Me.cboSearchFontSize < 6 Or Me.cboSearchFontSize > 72 Then Me.cboSearchFontSize = "12": Exit Sub
    
    Me.txtSearchInput.FontSize = Me.cboSearchFontSize
End Sub

Private Sub cboAddDataFontSize_AfterUpdate()

    If IsNull(Me.cboAddDataFontSize) Or IsAllNumbers(Me.cboAddDataFontSize) = False Or Trim(Me.txtAddDataInput.Value) = "" Or _
    Me.cboAddDataFontSize < 6 Or Me.cboAddDataFontSize > 72 Then Me.cboAddDataFontSize = "12": Exit Sub
    
    Me.txtAddDataInput.FontSize = Me.cboAddDataFontSize
End Sub

Private Sub cboSearchDelimiter_Change()
    Dim str As String
    
    str = Trim(cboSearchDelimiter.Value)
    
    If InStr(LCase(str), "auto") > 0 Then
        cboSearchDelimiter.TextAlign = 1
    
    Else
        cboSearchDelimiter.TextAlign = 2
    End If
        
End Sub

Private Sub cboAddDataDelimiter_Change()
    Dim str As String
    
    str = Trim(cboAddDataDelimiter.Value)
    
    If InStr(LCase(str), "auto") > 0 Then
        cboAddDataDelimiter.TextAlign = 1
    
    Else
        cboAddDataDelimiter.TextAlign = 2
    End If
        
End Sub


'change import mode DISPLAY
Private Sub cboImportMode_Change()
    Dim str As String, currentStr As String, updateStr As String
    
    str = Trim(LCase(Left(cboImportMode.Value, InStr(cboImportMode.Value, " ") - 1)))
    'Debug.Print str
    
    currentStr = Trim(LCase(txtAddDataInput.Value))
    updateStr = ""
    
    Select Case str
    
    Case "default"
    lblSelectorTypeAddData.Visible = False
    cboSelectorTypeAdd.Visible = False
    
    updateStr = DefineFormDefaults("txtAddDataDefault")
    
    Case "unrelated"
    lblSelectorTypeAddData.Visible = True
    cboSelectorTypeAdd.Visible = True

    updateStr = DefineFormDefaults("txtAddDataUnrelated")
    
    End Select
    
    'dont overwrite user input
    If currentStr <> "" And InStr(currentStr, "input / paste") <> 1 Then Exit Sub
        
    'reset font sizes
    txtAddDataInput.Value = updateStr
    txtAddDataInput.ForeColor = &H666666
    txtAddDataInput.FontSize = 10
    txtAddDataInput.TextFormat = acTextFormatHTMLRichText
End Sub


Private Sub cboSAdd_Change()
    
    If Trim(LCase(cboSAdd)) = "yes" Then
        txtSTokenAdd.Visible = True
        txtHelpAdd.Visible = True
    
    Else
        txtSTokenAdd.Visible = False
        txtHelpAdd.Visible = False
    End If
End Sub

Private Sub cboSearchTarget_Change()

    If Trim(LCase(cboSearchTarget)) = "graywolfe only" Then
        txtSToken.Visible = False
        lblSToken.Visible = False
        txtHelp.Visible = False

    Else
        txtSToken.Visible = True
        txtSToken.Visible = True
        txtHelp.Visible = True
    End If
End Sub

'+++++++++++++++++++++

'debounce
Private Sub txtSearchInput_Change()
    Me.TimerInterval = 500
End Sub

Private Sub txtAddDataInput_Change()
    Me.TimerInterval = 500
End Sub

Private Sub Form_Timer()
    Dim txt As String, str As String
    Dim res As String
    
    Me.TimerInterval = 0
    txt = Screen.ActiveControl.Name
    
    Select Case txt
    
    Case "txtAddDataInput"
        str = DetectInputDelimiter(Me.txtAddDataInput.Text)
        Me.cboAddDataDelimiter.DefaultValue = """" & str & """"
        Me.cboAddDataDelimiter.TextAlign = 2
        
        If InStr(LCase(str), "auto") > 0 Then Me.cboAddDataDelimiter.TextAlign = 1
    
    Case "txtSearchInput"
        str = DetectInputDelimiter(Me.txtSearchInput.Text)
        Me.cboSearchDelimiter.DefaultValue = """" & str & """"
        Me.cboSearchDelimiter.TextAlign = 2
        
        If InStr(LCase(str), "auto") > 0 Then Me.cboSearchDelimiter.TextAlign = 1

    End Select
    
    'Debug.Print "RETURN STR: " & str
End Sub

'+++++++++++++++++++++++++

Private Sub txtHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.txtHelp.BackColor = RGB(200, 220, 255) Then Exit Sub
    Me.lblHelpPopup.Visible = True
    
    ' Light blue
    Me.txtHelp.BackColor = RGB(200, 220, 255)
End Sub

'reset help button
Private Sub rectSearchInput_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.txtHelp.BackColor = RGB(216, 216, 216) Then Exit Sub
    
    Me.lblHelpPopup.Visible = False
    Me.txtHelp.BackColor = RGB(216, 216, 216)
End Sub

'------------------------

Private Sub txtHelpAdd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.txtHelpAdd.BackColor = RGB(200, 220, 255) Then Exit Sub
    Me.lblHelpPopupAdd.Visible = True
    
    
    
    ' Light blue
    Me.txtHelpAdd.BackColor = RGB(200, 220, 255)
End Sub

Private Sub rectAddData_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.txtHelpAdd.BackColor = RGB(216, 216, 216) Then Exit Sub
    
    Me.lblHelpPopupAdd.Visible = False
    Me.txtHelpAdd.BackColor = RGB(216, 216, 216)
End Sub
