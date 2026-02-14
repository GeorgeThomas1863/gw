*****************************************

Form_frmResults_SSubform

*****************************************

Option Compare Database

Private Sub Form_Current()

    'Debug.Print "TXT SELECOTR ID: " & txtSelectorId

    txtRowId = Nz(txtSId2, "")
End Sub

'sorting
Private Sub lblSelectorSearched_Click()
    If Me.OrderBy = "[selector]" Then
        Me.OrderBy = "[selector] DESC"
    
    Else
        Me.OrderBy = "[selector]"
    End If
    
    Me.OrderByOn = True
End Sub

Private Sub lblCaseSerial_Click()
    If Me.OrderBy = "[caseSerialFull]" Then
        Me.OrderBy = "[caseSerialFull] DESC"
    
    Else
        Me.OrderBy = "[caseSerialFull]"
    End If
    Me.OrderByOn = True
End Sub

Private Sub lblTitle_Click()
    If Me.OrderBy = "[docTitle]" Then
        Me.OrderBy = "[docTitle] DESC"
    
    Else
        Me.OrderBy = "[docTitle]"
    End If
    Me.OrderByOn = True
End Sub

Private Sub lblItemType_Click()
    If Me.OrderBy = "[docSubType]" Then
        Me.OrderBy = "[docSubType] DESC"
    
    Else
        Me.OrderBy = "[docSubType]"
    End If
    Me.OrderByOn = True
End Sub

Private Sub lblOffice_Click()
    If Me.OrderBy = "[office]" Then
        Me.OrderBy = "[office] DESC"
    
    Else
        Me.OrderBy = "[office]"
    End If
    Me.OrderByOn = True
End Sub

Private Sub lblAuthor_Click()
    If Me.OrderBy = "[author]" Then
        Me.OrderBy = "[author] DESC"
    
    Else
        Me.OrderBy = "[author]"
    End If
    Me.OrderByOn = True
End Sub

Private Sub lblCreatedDate_Click()
    If Me.OrderBy = "[createdDate]" Then
        Me.OrderBy = "[createdDate] DESC"
    
    Else
        Me.OrderBy = "[createdDate]"
    End If
    Me.OrderByOn = True
End Sub

'+++++++++++++++++++++++++++++++++++++

'links
Private Sub txtSelector_DblClick(Cancel As Integer)
    Dim uniqueIdStr As String, linkStr As String
    
    uniqueIdStr = Nz(Trim(Me.txtSId2), "")
    If uniqueIdStr = "" Or LCase(uniqueIdStr) = "null" Then Exit Sub
    
    'extract unique id
    linkStr = Trim(Mid(uniqueIdStr, 2, InStr(uniqueIdStr, "_") - 2))
    Debug.Print "UNIQUE ID STR: " & uniqueIdStr
    Debug.Print "LINK STR: " & linkStr
    
    OpenLink "https://S.Fnet.F/apps/desktop/#/main/serial/" & linkStr
    

End Sub

Private Sub txtPrimarySerial_DblClick(Cancel As Integer)
    Dim uniqueIdStr As String, linkStr As String
    
    uniqueIdStr = Nz(Trim(Me.txtSId2), "")
    If uniqueIdStr = "" Or LCase(uniqueIdStr) = "null" Then Exit Sub
    
    linkStr = Trim(Mid(uniqueIdStr, 2, InStr(uniqueIdStr, "_") - 2))
    OpenLink "https://S.Fnet.F/apps/desktop/#/main/serial/" & linkStr
End Sub


Private Sub txtDocTitle_DblClick(Cancel As Integer)
    Dim uniqueIdStr As String, linkStr As String
    
    uniqueIdStr = Nz(Trim(Me.txtSId2), "")
    If uniqueIdStr = "" Or LCase(uniqueIdStr) = "null" Then Exit Sub
    
    linkStr = Trim(Mid(uniqueIdStr, 2, InStr(uniqueIdStr, "_") - 2))
    OpenLink "https://S.Fnet.F/apps/desktop/#/main/serial/" & linkStr
End Sub

Private Sub txtDocSubType_DblClick(Cancel As Integer)
    Dim uniqueIdStr As String, linkStr As String
    
    uniqueIdStr = Nz(Trim(Me.txtSId2), "")
    If uniqueIdStr = "" Or LCase(uniqueIdStr) = "null" Then Exit Sub
    
    linkStr = Trim(Mid(uniqueIdStr, 2, InStr(uniqueIdStr, "_") - 2))
    OpenLink "https://S.Fnet.F/apps/desktop/#/main/serial/" & linkStr
End Sub


Private Sub txtOffice_DblClick(Cancel As Integer)
    Dim uniqueIdStr As String, linkStr As String
    
    uniqueIdStr = Nz(Trim(Me.txtSId2), "")
    If uniqueIdStr = "" Or LCase(uniqueIdStr) = "null" Then Exit Sub
    
    linkStr = Trim(Mid(uniqueIdStr, 2, InStr(uniqueIdStr, "_") - 2))
    OpenLink "https://S.Fnet.F/apps/desktop/#/main/serial/" & linkStr
End Sub


Private Sub txtAuthor_DblClick(Cancel As Integer)
    Dim uniqueIdStr As String, linkStr As String
    
    uniqueIdStr = Nz(Trim(Me.txtSId2), "")
    If uniqueIdStr = "" Or LCase(uniqueIdStr) = "null" Then Exit Sub
    
    linkStr = Trim(Mid(uniqueIdStr, 2, InStr(uniqueIdStr, "_") - 2))
    OpenLink "https://S.Fnet.F/apps/desktop/#/main/serial/" & linkStr
End Sub


Private Sub txtCreatedDate_DblClick(Cancel As Integer)
    Dim uniqueIdStr As String, linkStr As String
    
    uniqueIdStr = Nz(Trim(Me.txtSId2), "")
    If uniqueIdStr = "" Or LCase(uniqueIdStr) = "null" Then Exit Sub
    
    linkStr = Trim(Mid(uniqueIdStr, 2, InStr(uniqueIdStr, "_") - 2))
    OpenLink "https://S.Fnet.F/apps/desktop/#/main/serial/" & linkStr
End Sub
