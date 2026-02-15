*****************************************

Form_frmTargetDetails

*****************************************

'DEAL WITH TYPE OTHER IN DISPLAY
'THROW TYPE CHECK POPUPS when addign selectors via targets display

Option Compare Database

Public targetIdCurrent As String
Public targetNameCurrent As String
Public caseNumberCurrent As String
Public personaNamesCurrent As String
Public addressesCurrent As String
Public emailsCurrent As String
Public phonesCurrent As String
Public ipsCurrent As String
Public otherCurrent As String
Public norksCurrent As String
Public laptopsCurrent As Long

Private Sub Form_Load()
    Debug.Print "FORM LOAD EVENT"
    
    'very slow
    'UpdateTotalSelectors
    
    'this fucks everything
    'ResetLocalTablesSP
    
    Me.subSelectors.Form.Filter = "[selectorCount] = DESC"
    Me.subSelectors.Form.OrderByOn = True
    
    If IsNull(Me.OpenArgs) Or Trim(Me.OpenArgs) = "" Then
        Me.txtTargetName.SetFocus
        Exit Sub
    End If
    
    Me.subSelectors.SetFocus
    DoCmd.FindRecord Me.OpenArgs
    Me.txtTargetName.SetFocus

End Sub

Private Sub btnClose_Click()
    
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub btnMerge_Click()

    DoCmd.OpenForm "frmMergeTargets", , , , , , Me.txtTargetId
End Sub

Private Sub btnReset_Click()
    Dim resetText As String, resetCheck As Integer
    
    'confirm reset
    resetText = DefineFormDefaults("resetTargets")
    resetCheck = MsgBox(resetText, vbYesNo, "ARE YOU SURE?!?")
    If resetCheck = vbNo Then Exit Sub
    
    ClearTableList "localTargets"
    
    SendDataLocalSP "Targets", "localTargets", "localTargets"
    UpdateLocalTableCount "localTargets"
    
    Me.subSelectors.Requery

End Sub

Private Sub btnSave_Click()
    Dim db As DAO.Database
    Dim strSQL As String
    
    MsgBox "Data sent to backend. Give it a couple seconds to update."
    
    'clears SP list and resets with local
    ClearTableList "Targets"
    ClearTableList "Selectors"
    
    SendDataLocalSP "localTargets", "Targets", "localTargets"
    SendDataLocalSP "localSelectors", "Selectors", "localSelectors"

    MsgBox ("Data saved!")
End Sub

'-------------------------------------

Private Sub txtTargetId_LostFocus()
    Dim inputStr As String, targetIdStr As String, ctlStr As String
    
    inputStr = Nz(Me.txtTargetId, "")
    If inputStr = targetIdCurrent Then Exit Sub
    
    'check if sure
    popupCheck = MsgBox(DefineFormDefaults("changeTargetId"), vbYesNo + vbDefaultButton2, "ARE YOU SURE!?!")
    If popupCheck = vbNo Then Me.txtTargetId = targetIdCurrent: Exit Sub
    
    'get from subform
    targetIdStr = Me.subSelectors.Form.txtTargetId
    ctl = Screen.ActiveControl.Name
    Debug.Print "CTL: " & ctl
        
    UpdateTargetStatsForm targetIdStr, inputStr, "targetId"
    targetIdCurrent = inputStr
    
    Me.subSelectors.Form.Requery
    Me.subSelectors.SetFocus
    DoCmd.FindRecord inputStr
    
    Me.Controls(ctl).SetFocus
    Me.Controls(ctl).SelStart = Len(inputStr) + 10 'extra for uncounted spaces
    Me.Controls(ctl).SelLength = 0
End Sub

Private Sub txtTargetName_LostFocus()
    Dim inputStr As String, targetIdStr As String
    
    inputStr = Nz(Me.txtTargetName, "")
    
    If inputStr = targetNameCurrent Then Exit Sub
    
    targetIdStr = Me.txtTargetId
    UpdateTargetStatsForm targetIdStr, inputStr, "targetName"
    targetNameCurrent = inputStr
    
    UpdateLastUpdated targetIdStr, "localTargets"
    
    Me.subSelectors.Form.Requery
    Me.subSelectors.SetFocus
    DoCmd.FindRecord targetIdStr
End Sub

Private Sub txtCaseNumber_LostFocus()
    Dim inputStr As String, targetIdStr As String
    
    inputStr = Nz(Me.txtCaseNumber, "")
    
    If inputStr = caseNumberCurrent Then Exit Sub
    
    targetIdStr = Me.txtTargetId
    UpdateTargetStatsForm targetIdStr, inputStr, "caseNumber"
    caseNumberCurrent = inputStr
    
    UpdateLastUpdated targetIdStr, "localTargets"
    
    Me.subSelectors.Form.Requery
    Me.subSelectors.SetFocus
    DoCmd.FindRecord targetIdStr

End Sub

Private Sub txtPersonaNames_LostFocus()
    Dim inputStr As String, targetIdStr As String
    
    inputStr = FixTargetDelimiters(Trim(Me.txtPersonaNames.Text))
    personaNamesCurrent = FixTargetDelimiters(Trim(personaNamesCurrent))
    
    If inputStr = personaNamesCurrent Then Exit Sub
    
    targetIdStr = Me.txtTargetId
    
    UpdateTargetSelectors targetIdStr, inputStr, personaNamesCurrent, "name"
    UpdateTargetsSelectorCountTbl targetIdStr
    personaNamesCurrent = inputStr
    
    UpdateLastUpdated targetIdStr, "localTargets"
    
    Me.subSelectors.Form.Requery
    Me.subSelectors.SetFocus
    DoCmd.FindRecord targetIdStr

End Sub

Private Sub txtAddresses_LostFocus()
    Dim inputStr As String, targetIdStr As String
    
    inputStr = FixTargetDelimiters(Trim(Me.txtAddresses.Text))
    addressesCurrent = FixTargetDelimiters(Trim(addressesCurrent))
    
    If inputStr = addressesCurrent Then Exit Sub
    
    targetIdStr = Me.txtTargetId
    
    UpdateTargetSelectors targetIdStr, inputStr, addressesCurrent, "address"
    UpdateTargetsSelectorCountTbl targetIdStr
    addressesCurrent = inputStr
    
    UpdateLastUpdated targetIdStr, "localTargets"
    
    Me.subSelectors.Form.Requery
    Me.subSelectors.SetFocus
    DoCmd.FindRecord targetIdStr
End Sub

Private Sub txtEmails_LostFocus()
    Dim inputStr As String, targetIdStr As String
    
    inputStr = FixTargetDelimiters(Trim(Me.txtEmails.Text))
    emailsCurrent = FixTargetDelimiters(Trim(emailsCurrent))
    
    If inputStr = emailsCurrent Then Exit Sub
    
    targetIdStr = Me.txtTargetId
    
    UpdateTargetSelectors targetIdStr, inputStr, emailsCurrent, "email"
    UpdateTargetsSelectorCountTbl targetIdStr
    emailsCurrent = inputStr
    
    UpdateLastUpdated targetIdStr, "localTargets"
    
    Me.subSelectors.Form.Requery
    Me.subSelectors.SetFocus
    DoCmd.FindRecord targetIdStr
End Sub

Private Sub txtPhones_LostFocus()
    Dim inputStr As String, targetIdStr As String
    
    inputStr = FixTargetDelimiters(Trim(Me.txtPhones.Text))
    phonesCurrent = FixTargetDelimiters(Trim(phonesCurrent))
    
    If inputStr = phonesCurrent Then Exit Sub
    
    targetIdStr = Me.txtTargetId
    
    UpdateTargetSelectors targetIdStr, inputStr, phonesCurrent, "phone"
    UpdateTargetsSelectorCountTbl targetIdStr
    phonesCurrent = inputStr
    
    UpdateLastUpdated targetIdStr, "localTargets"
    
    Me.subSelectors.Form.Requery
    Me.subSelectors.SetFocus
    DoCmd.FindRecord targetIdStr
End Sub

Private Sub txtIPs_LostFocus()
    Dim inputStr As String, targetIdStr As String
    
    inputStr = FixTargetDelimiters(Trim(Me.txtIPs.Text))
    ipsCurrent = FixTargetDelimiters(Trim(ipsCurrent))
    
    If inputStr = ipsCurrent Then Exit Sub
    
    targetIdStr = Me.txtTargetId
    
    UpdateTargetSelectors targetIdStr, inputStr, ipsCurrent, "ip"
    UpdateTargetsSelectorCountTbl targetIdStr
    ipsCurrent = inputStr
    
    UpdateLastUpdated targetIdStr, "localTargets"
    
    Me.subSelectors.Form.Requery
    Me.subSelectors.SetFocus
    DoCmd.FindRecord targetIdStr
End Sub

'DEAL WITH BELOW
Private Sub txtOther_LostFocus()
    Dim inputStr As String, targetIdStr As String
    
    inputStr = FixTargetDelimiters(Trim(Me.txtOther.Text))
    otherCurrent = FixTargetDelimiters(Trim(otherCurrent))
    
    If inputStr = otherCurrent Then Exit Sub
    
    targetIdStr = Me.txtTargetId
    
    UpdateTargetSelectors targetIdStr, inputStr, otherCurrent, "other"
    UpdateTargetsSelectorCountTbl targetIdStr
    otherCurrent = inputStr
    
    UpdateLastUpdated targetIdStr, "localTargets"
    
    Me.subSelectors.Form.Requery
    Me.subSelectors.SetFocus
    DoCmd.FindRecord targetIdStr
End Sub

Private Sub txtLaptopCount_LostFocus()
    Dim inputStr As String, targetIdStr As String
    
    inputStr = Nz(Me.txtLaptopCount, "")
    
    If inputStr = laptopsCurrent Then Exit Sub
    
    targetIdStr = Me.txtTargetId
    UpdateTargetStatsForm targetIdStr, inputStr, "laptopCount"
    laptopsCurrent = inputStr
    
    UpdateLastUpdated targetIdStr, "localTargets"
    
    Me.subSelectors.Form.Requery
    Me.subSelectors.SetFocus
    DoCmd.FindRecord targetIdStr
End Sub