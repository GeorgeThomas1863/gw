*****************************************

Form_frmSearchResults

*****************************************

Option Compare Database

Private Sub Form_Load()
    Dim arr() As String
    
    If IsNull(Me.OpenArgs) Then Exit Sub
    
    arr = Split(Me.OpenArgs, "!!")
    
    'FIX / BUILD
    Me.lblSHitsCount.Caption = 0
    
    Me.lblSearchedCount.Caption = arr(0)
    Me.lblGrayWolfeHitsCount.Caption = arr(1)
    Me.lblNotFoundCount.Caption = arr(0) - arr(1)
    
    Me.lblGrayWolfeTargetsCount.Caption = GetLocalTableCount("knownTargets")

End Sub


'btn clear search click
Private Sub btnClose_Click()
    
    DoCmd.Close acForm, Me.Name
End Sub