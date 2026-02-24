*****************************************

Form_frmSchemaDetection

*****************************************

Option Compare Database

'Public dataStr As String

Private Sub Form_Load()
    Dim arr() As String, searchArr() As String
    Dim itemStr As String, captionStr As String, cboStr As String, percentStr As String
    Dim colMax As Long, i As Long, j As Long, k As Long
    
    If IsNull(Me.OpenArgs) Then Exit Sub
    
    'Debug.Print "OPEN ARGS: " & Me.OpenArgs
    
    arr = Split(Me.OpenArgs, "+++")
    colMax = UBound(arr) + 1
    
    'unhide everything
    For k = 1 To 6
        Me.Controls("lblColumn" & k).Visible = True
        Me.Controls("lblDetected" & k).Visible = True
        Me.Controls("lblSetTo" & k).Visible = True
        Me.Controls("cboColumn" & k).Visible = True
        Me.subFormSchema.Form.Controls("txtColumn" & k).ColumnHidden = True
        Me.subFormSchema.Form.Controls("lblColumn" & k).Visible = True
    Next
        
    'hide extra columns
    If colMax < 6 Then
        For k = colMax + 1 To 6
            
            Me.Controls("lblColumn" & k).Visible = False
            Me.Controls("lblDetected" & k).Visible = False
            Me.Controls("lblSetTo" & k).Visible = False
            Me.Controls("cboColumn" & k).Visible = False
            Me.subFormSchema.Form.Controls("txtColumn" & k).ColumnHidden = True
            Me.subFormSchema.Form.Controls("lblColumn" & k).Visible = False
        Next
    End If
    
    'fill in values
    For i = LBound(arr) To UBound(arr)
        itemStr = Trim(arr(i))
        percentStr = Trim(Left(itemStr, InStr(itemStr, "%") - 1))
        captionStr = FixPlural(Trim(Mid(itemStr, InStr(itemStr, "%") + 2, Len(itemStr))))
        
        'custom changes so items fit
        If Right(LCase(Trim(captionStr)), 3) = "ips" Then captionStr = "IPs"
        If InStr(LCase(Trim(captionStr)), "null") > 0 Then captionStr = "NULL"
        If InStr(LCase(captionStr), "persona") > 0 Then captionStr = "Names"
        If InStr(LCase(captionStr), "street") > 0 Then captionStr = "Addresses"
        
        Me.Controls("lblDetected" & i + 1).Caption = percentStr & "% " & captionStr
        Me.Controls("cboColumn" & i + 1).Value = captionStr
    Next
    
End Sub

Private Sub btnSubmit_Click()
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim typeStr As String, addStr As String, itemStr As String, rowStr As String
    Dim colMax As Long, recordCount As Long
    Dim i As Long
    
    colMax = DetectColMaxRS("tempSchema")
    'Debug.Print "!!!!COL MAX RS: " & colMax
    
    'gets data from schema for import (letting user change if needed)
    typeStr = BuildTypeStr(colMax)
    'Debug.Print "!!!! TYPE STR: " & typeStr

    DoCmd.Close acForm, Me.Name

    'TRIGGER ACTUAL ADD HERE
    RunAddSchemaData typeStr, colMax
End Sub

'btn clear search click
Private Sub btnCancel_Click()
    
    DoCmd.Close acForm, Me.Name
    
    'MsgBox "Upload Canceled", vbInformation, "CANCELED"
    
    'maybe clear stuff?
    'Exit Sub
    
End Sub

'gets data from combo box user input
Private Function BuildTypeStr(colMax As Long) As String
    Dim itemStr As String, returnStr As String
    Dim i As Long

    returnStr = ""
    For i = 1 To colMax
        itemStr = Trim(Nz(Me.Controls("cboColumn" & i).Value, ""))
        If itemStr = "" Or LCase(itemStr) = "null" Then
            returnStr = returnStr & "null" & "!!"
        Else
            returnStr = returnStr & CleanSelectorInput(itemStr) & "!!"
        End If
    Next

    If Len(returnStr) > 2 Then BuildTypeStr = Left(returnStr, Len(returnStr) - 2)
End Function
