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
    Dim popupCheck As Integer

    inputStr = Nz(Me.txtTargetId, "")
    If inputStr = targetIdCurrent Then Exit Sub

    'check if sure
    popupCheck = MsgBox(DefineFormDefaults("changeTargetId"), vbYesNo + vbDefaultButton2, "ARE YOU SURE!?!")
    If popupCheck = vbNo Then Me.txtTargetId = targetIdCurrent: Exit Sub

    'get from subform
    targetIdStr = Me.subSelectors.Form.txtTargetId
    ctlStr = Screen.ActiveControl.Name
        
    UpdateTargetStatsForm targetIdStr, inputStr, "targetId"
    targetIdCurrent = inputStr
    
    Me.subSelectors.Form.Requery
    Me.subSelectors.SetFocus
    DoCmd.FindRecord inputStr
    
    Me.Controls(ctlStr).SetFocus
    Me.Controls(ctlStr).SelStart = Len(inputStr) + 10 'extra for uncounted spaces
    Me.Controls(ctlStr).SelLength = 0
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

Private Sub SaveSelectorField(ctl As Control, ByVal currentVal As String, ByVal selectorType As String, ByRef newCurrentVal As String)
    Dim inputStr As String, targetIdStr As String, currentStr As String

    inputStr = FixTargetDelimiters(Trim(ctl.Text))
    currentStr = FixTargetDelimiters(Trim(currentVal))

    If inputStr = currentStr Then Exit Sub

    targetIdStr = Me.txtTargetId

    UpdateTargetSelectors targetIdStr, inputStr, currentStr, selectorType, BuildDataSource("Manual")
    UpdateTargetsSelectorCountTbl targetIdStr
    newCurrentVal = inputStr

    UpdateLastUpdated targetIdStr, "localTargets"

    Me.subSelectors.Form.Requery
    Me.subSelectors.SetFocus
    DoCmd.FindRecord targetIdStr
End Sub

Private Sub txtPersonaNames_LostFocus()
    SaveSelectorField Me.txtPersonaNames, personaNamesCurrent, "name", personaNamesCurrent
End Sub

Private Sub txtAddresses_LostFocus()
    SaveSelectorField Me.txtAddresses, addressesCurrent, "address", addressesCurrent
End Sub

Private Sub txtEmails_LostFocus()
    SaveSelectorField Me.txtEmails, emailsCurrent, "email", emailsCurrent
End Sub

Private Sub txtPhones_LostFocus()
    SaveSelectorField Me.txtPhones, phonesCurrent, "phone", phonesCurrent
End Sub

Private Sub txtIPs_LostFocus()
    SaveSelectorField Me.txtIPs, ipsCurrent, "ip", ipsCurrent
End Sub

Private Sub txtOther_LostFocus()
    SaveSelectorField Me.txtOther, otherCurrent, "other", otherCurrent
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