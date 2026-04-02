'*****************************************

' MOD 5_UTIL_EXPORT

'*****************************************

Option Compare Database

Option Explicit

Public Sub ExportAllVBACode()
    ' PREREQUISITE: File > Options > Trust Center > Trust Center Settings >
    '   Macro Settings > enable "Trust access to the VBA project object model"

    ' Guard: check VBE access before proceeding
    On Error Resume Next
    Dim testVBE As Object
    Set testVBE = Application.VBE.ActiveVBProject
    If Err.Number <> 0 Then
        MsgBox "VBA project access is blocked." & vbCrLf & vbCrLf & _
               "Go to: File > Options > Trust Center > Trust Center Settings" & vbCrLf & _
               "   > Macro Settings > enable 'Trust access to the VBA project object model'", _
               vbCritical, "ExportAllVBACode"
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    Dim dbPath As String
    Dim modulesDir As String
    Dim formsDir As String
    Dim comp As Object
    Dim exportedCount As Long
    Dim skippedCount As Long
    Dim failedCount As Long
    Dim outPath As String
    Dim failedNames As String

    dbPath = CurrentProject.Path
    modulesDir = dbPath & "\modules\"
    formsDir = dbPath & "\forms\"

    EnsureFolderExists modulesDir
    EnsureFolderExists formsDir

    exportedCount = 0
    skippedCount = 0
    failedCount = 0
    failedNames = ""

    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        On Error Resume Next  ' isolate per-component failures
        Dim compErr As Long
        compErr = 0
        Select Case comp.Type
            Case 1  ' StdModule (.bas)
                outPath = modulesDir & comp.Name & ".bas"
                comp.Export outPath
                compErr = Err.Number
            Case 2  ' ClassModule (.cls)
                outPath = formsDir & comp.Name & ".cls"
                comp.Export outPath
                compErr = Err.Number
            Case 100  ' Document (Access form/report code)
                If comp.CodeModule.CountOfLines > 0 Then
                    outPath = formsDir & comp.Name & ".cls"
                    comp.Export outPath
                    compErr = Err.Number
                Else
                    skippedCount = skippedCount + 1
                    GoTo NextComp
                End If
            Case Else
                skippedCount = skippedCount + 1
                GoTo NextComp
        End Select
        On Error GoTo ErrorHandler
        If compErr <> 0 Then
            failedCount = failedCount + 1
            failedNames = failedNames & vbCrLf & "  " & comp.Name
        Else
            exportedCount = exportedCount + 1
        End If
NextComp:
    Next comp

    Dim msg As String
    msg = "Export complete." & vbCrLf & _
          "Exported: " & exportedCount & vbCrLf & _
          "Skipped:  " & skippedCount
    If failedCount > 0 Then
        msg = msg & vbCrLf & "Failed:   " & failedCount & failedNames
    End If
    MsgBox msg, vbInformation, "ExportAllVBACode"
    Exit Sub

ErrorHandler:
    MsgBox "Export failed: " & Err.Description & " (Error " & Err.Number & ")", _
           vbCritical, "ExportAllVBACode"
End Sub

Private Sub EnsureFolderExists(folderPath As String)
    Dim cleanPath As String
    cleanPath = folderPath
    If Right(cleanPath, 1) = "\" Then cleanPath = Left(cleanPath, Len(cleanPath) - 1)
    If Dir(cleanPath, vbDirectory) = "" Then
        MkDir cleanPath
    End If
End Sub
