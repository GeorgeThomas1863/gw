Form_frmMergeTargets

*****************************************

*****************************************

Option Compare Database

Option Explicit

Private Sub Form_Load()
    'pre-fill keep targetId from OpenArgs (passed from frmTargetDetails)
    If Not IsNull(Me.OpenArgs) And Trim(Nz(Me.OpenArgs, "")) <> "" Then
        Me.txtKeepTargetId = Trim(Me.OpenArgs)
        UpdateKeepDisplay
    End If
End Sub

Private Sub txtKeepTargetId_AfterUpdate()
    UpdateKeepDisplay
End Sub

Private Sub txtAbsorbTargetId_AfterUpdate()
    UpdateAbsorbDisplay
End Sub

Private Sub btnSwap_Click()
    Dim tempStr As String
    tempStr = Nz(Me.txtKeepTargetId, "")
    Me.txtKeepTargetId = Nz(Me.txtAbsorbTargetId, "")
    Me.txtAbsorbTargetId = tempStr

    UpdateKeepDisplay
    UpdateAbsorbDisplay
End Sub

Private Sub btnMerge_Click()
    Dim keepStr As String, absorbStr As String
    Dim confirmText As String, confirmCheck As Integer
    Dim mergeResult As Boolean

    keepStr = Trim(Nz(Me.txtKeepTargetId, ""))
    absorbStr = Trim(Nz(Me.txtAbsorbTargetId, ""))

    If keepStr = "" Or absorbStr = "" Then
        MsgBox "Both target IDs are required.", , "#FAIL"
        Exit Sub
    End If

    If keepStr = absorbStr Then
        MsgBox "Cannot merge a target with itself.", , "#FAIL"
        Exit Sub
    End If

    confirmText = "MERGE targets?" & vbLf & vbLf & _
    "KEEP: " & keepStr & " (" & Nz(Me.lblKeepName.Caption, "?") & ")" & vbLf & _
    "ABSORB (will be DELETED): " & absorbStr & " (" & Nz(Me.lblAbsorbName.Caption, "?") & ")" & vbLf & vbLf & _
    "All selectors from the absorbed target will move to the kept target. This cannot be undone."

    confirmCheck = MsgBox(confirmText, vbYesNo + vbDefaultButton2, "ARE YOU SURE?!?")
    If confirmCheck = vbNo Then Exit Sub

    mergeResult = MergeTargets(keepStr, absorbStr)

    If mergeResult Then
        MsgBox "Targets merged successfully!" & vbLf & vbLf & _
        "All selectors now belong to " & keepStr & "." & vbLf & _
        "Target " & absorbStr & " has been deleted.", , "#SUCCESS"

        'refresh display
        Me.txtAbsorbTargetId = ""
        Me.lblAbsorbName.Caption = ""
        Me.lblAbsorbCount.Caption = ""
        UpdateKeepDisplay
    Else
        ThrowError 1972, "keep=" & keepStr & " absorb=" & absorbStr
    End If
End Sub

Private Sub btnClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

'helper - update keep side display
Private Sub UpdateKeepDisplay()
    Dim targetIdStr As String, nameStr As String
    Dim countVal As Long

    targetIdStr = Trim(Nz(Me.txtKeepTargetId, ""))
    If targetIdStr = "" Then
        Me.lblKeepName.Caption = ""
        Me.lblKeepCount.Caption = ""
        Exit Sub
    End If

    nameStr = SearchTargetNameTbl(targetIdStr)
    If nameStr = "" Then nameStr = "[not found]"
    Me.lblKeepName.Caption = nameStr

    Me.lblKeepCount.Caption = GetSelectorCountForTarget(targetIdStr) & " selectors"
End Sub

'helper - update absorb side display
Private Sub UpdateAbsorbDisplay()
    Dim targetIdStr As String, nameStr As String
    Dim countVal As Long

    targetIdStr = Trim(Nz(Me.txtAbsorbTargetId, ""))
    If targetIdStr = "" Then
        Me.lblAbsorbName.Caption = ""
        Me.lblAbsorbCount.Caption = ""
        Exit Sub
    End If

    nameStr = SearchTargetNameTbl(targetIdStr)
    If nameStr = "" Then nameStr = "[not found]"
    Me.lblAbsorbName.Caption = nameStr

    Me.lblAbsorbCount.Caption = GetSelectorCountForTarget(targetIdStr) & " selectors"
End Sub

'helper - get selector count for a target
Private Function GetSelectorCountForTarget(targetIdStr As String) As Long
    Dim rs As DAO.Recordset

    Set rs = OpenRS("SELECT COUNT(*) AS cnt FROM [localSelectors] WHERE [targetId] = ?", targetIdStr)

    If Not rs.EOF Then
        GetSelectorCountForTarget = Nz(rs!cnt, 0)
    Else
        GetSelectorCountForTarget = 0
    End If
    rs.Close
End Function