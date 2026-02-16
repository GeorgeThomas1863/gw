*****************************************

Form_frmResults_GWSubform

*****************************************

Option Compare Database

Private Sub Form_Current()

    'Debug.Print "TXT SELECOTR ID: " & txtSelectorId

    txtRowId = Nz(txtSelectorId, "")
End Sub

'sorting

'++++++++++++++++++++++++++++++++
Private Sub lblSelectorSearched_Click()
    If Me.OrderBy = "[selector]" Then
        Me.OrderBy = "[selector] DESC"
    
    Else
        Me.OrderBy = "[selector]"
    End If
    Me.OrderByOn = True
End Sub

Private Sub lblSHits_Click()
    If Me.OrderBy = "[SHits]" Then
        Me.OrderBy = "[SHits] DESC"
    
    Else
        Me.OrderBy = "[SHits]"
    End If
    Me.OrderByOn = True
End Sub

Private Sub lblInGW_Click()
    If Me.OrderBy = "[inGrayWolfe]" Then
        Me.OrderBy = "[inGrayWolfe] DESC"
    
    Else
        Me.OrderBy = "[inGrayWolfe]"
    End If
    Me.OrderByOn = True
End Sub

Private Sub lblAddedGW_Click()
    If Me.OrderBy = "[addedToGrayWolfe]" Then
        Me.OrderBy = "[addedToGrayWolfe] DESC"
    
    Else
        Me.OrderBy = "[addedToGrayWolfe]"
    End If
    Me.OrderByOn = True
End Sub


Private Sub lblTargetName_Click()
    If Me.OrderBy = "[targetName]" Then
        Me.OrderBy = "[targetName] DESC"
    
    Else
        Me.OrderBy = "[targetName]"
    End If
    Me.OrderByOn = True
End Sub

Private Sub lblTargetId_Click()
    If Me.OrderBy = "[targetId]" Then
        Me.OrderBy = "[targetId] DESC"
    
    Else
        Me.OrderBy = "[targetId]"
    End If
    Me.OrderByOn = True
End Sub

Private Sub lblLastUpdated_Click()
    If Me.OrderBy = "[lastUpdated]" Then
        Me.OrderBy = "[lastUpdated] DESC"
    
    Else
        Me.OrderBy = "[lastUpdated]"
    End If
    Me.OrderByOn = True
End Sub

Private Sub lblLastUpdatedBy_Click()
    If Me.OrderBy = "[lastUpdatedBy]" Then
        Me.OrderBy = "[lastUpdatedBy] DESC"
    
    Else
        Me.OrderBy = "[lastUpdatedBy]"
    End If
    Me.OrderByOn = True
End Sub

'++++++++++++++++++++++++++++

'open targets form
Private Sub OpenTargetFromRow()
    Dim targetIdStr As String

    targetIdStr = Nz(Trim(Me.targetId), "")
    If targetIdStr = "" Or LCase(targetIdStr) = "null" Then Exit Sub

    DoCmd.OpenForm "frmTargetDetails", acNormal, , , , , targetIdStr
End Sub

Private Sub txtSelector_DblClick(Cancel As Integer): OpenTargetFromRow: End Sub
Private Sub txtSHits_DblClick(Cancel As Integer): OpenTargetFromRow: End Sub
Private Sub txtInGW_DblClick(Cancel As Integer): OpenTargetFromRow: End Sub
Private Sub txtAddedGW_DblClick(Cancel As Integer): OpenTargetFromRow: End Sub
Private Sub txtTargetName_DblClick(Cancel As Integer): OpenTargetFromRow: End Sub
Private Sub txtTargetId_DblClick(Cancel As Integer): OpenTargetFromRow: End Sub
Private Sub txtLastUpdated_DblClick(Cancel As Integer): OpenTargetFromRow: End Sub
Private Sub txtLastUpdatedBy_DblClick(Cancel As Integer): OpenTargetFromRow: End Sub
