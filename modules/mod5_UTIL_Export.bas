*****************************************

MOD 5_UTIL_EXPORT

*****************************************

Option Compare Database

Option Explicit

Public Sub ExportAllVBACode()
    Dim dbPath As String
    Dim modulesDir As String
    Dim formsDir As String
    Dim comp As Object
    Dim exportedCount As Long
    Dim skippedCount As Long
    Dim outPath As String

    dbPath = CurrentProject.Path
    modulesDir = dbPath & "\modules\"
    formsDir = dbPath & "\forms\"

    EnsureFolderExists modulesDir
    EnsureFolderExists formsDir

    exportedCount = 0
    skippedCount = 0

    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        Select Case comp.Type
            Case 1  ' StdModule (.bas)
                outPath = modulesDir & comp.Name & ".bas"
                comp.Export outPath
                exportedCount = exportedCount + 1
            Case 2  ' ClassModule (.cls)
                outPath = formsDir & comp.Name & ".cls"
                comp.Export outPath
                exportedCount = exportedCount + 1
            Case 100  ' Document (Access form/report code)
                If comp.CodeModule.CountOfLines > 0 Then
                    outPath = formsDir & comp.Name & ".cls"
                    comp.Export outPath
                    exportedCount = exportedCount + 1
                Else
                    skippedCount = skippedCount + 1
                End If
            Case Else
                skippedCount = skippedCount + 1
        End Select
    Next comp

    MsgBox "Export complete." & vbCrLf & _
           "Exported: " & exportedCount & vbCrLf & _
           "Skipped:  " & skippedCount, vbInformation, "ExportAllVBACode"
End Sub

Private Sub EnsureFolderExists(folderPath As String)
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
End Sub
