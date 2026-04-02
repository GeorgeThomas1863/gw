*****************************************

Form_frmTargetDetails_Subform

*****************************************

Option Compare Database

Private Sub Form_Current()
    Dim arr() As String, itemArr() As String
    Dim dataStr As String, itemStr As String, selectorStr As String, typeStr As String, currentStr As String
    Dim i As Long
    'hidden txtbox
    txtRowId = Nz(txtTargetId, "")
    
    Debug.Print "TXT TARGET ID: " & txtTargetId
    Debug.Print "TXT ROW ID: " & txtRowId
    
    If txtRowId = "" Then Exit Sub
    
    'fill from subform
    Form_frmTargetDetails.txtTargetId = Me!txtTargetId
    Form_frmTargetDetails.txtTargetName = Nz(Me!txtTargetName, "")
    Form_frmTargetDetails.txtCaseNumber = Me!txtCaseNumber
    
    'clear all other items
    Form_frmTargetDetails.txtPersonaNames = ""
    Form_frmTargetDetails.txtAddresses = ""
    Form_frmTargetDetails.txtEmails = ""
    Form_frmTargetDetails.txtPhones = ""
    Form_frmTargetDetails.txtIPs = ""
    Form_frmTargetDetails.txtOther = ""
    Form_frmTargetDetails.txtNorks = ""
    Form_frmTargetDetails.txtLaptopCount = ""
    
    'FILL calculated DATA
    FillTargetDisplayForm txtRowId
    FillTargetLaptopCount txtRowId
    
    'track current data
    Form_frmTargetDetails.targetIdCurrent = Trim(Nz(Form_frmTargetDetails.txtTargetId, ""))
    Form_frmTargetDetails.targetNameCurrent = Trim(Nz(Form_frmTargetDetails.txtTargetName, ""))
    Form_frmTargetDetails.caseNumberCurrent = Trim(Nz(Form_frmTargetDetails.txtCaseNumber, ""))
    Form_frmTargetDetails.personaNamesCurrent = Trim(Nz(Form_frmTargetDetails.txtPersonaNames, ""))
    Form_frmTargetDetails.addressesCurrent = Trim(Nz(Form_frmTargetDetails.txtAddresses, ""))
    Form_frmTargetDetails.emailsCurrent = Trim(Nz(Form_frmTargetDetails.txtEmails, ""))
    Form_frmTargetDetails.phonesCurrent = Trim(Nz(Form_frmTargetDetails.txtPhones, ""))
    Form_frmTargetDetails.ipsCurrent = Trim(Nz(Form_frmTargetDetails.txtIPs, ""))
    Form_frmTargetDetails.otherCurrent = Trim(Nz(Form_frmTargetDetails.txtOther, ""))
    Form_frmTargetDetails.norksCurrent = Trim(Nz(Form_frmTargetDetails.txtNorks, ""))
    '[laptop tracking handled in query]

End Sub

'Private Sub Form_Load()
'
'End Sub