*****************************************

MOD 5_UTIL_DELETE

*****************************************

Option Compare Database

Option Explicit

Sub DeleteAllListItemsSP()
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim tbl As String, selectorCount As Long, targetCount As Long
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    'SELECTORS SECTION
    tbl = "Selectors"

    Set db = CurrentDb
    Set rs = db.OpenRecordset(tbl, dbOpenDynaset)
    
    ' Count records
    rs.MoveLast
    rs.MoveFirst
    selectorCount = rs.recordCount
    
    ' Confirm deletion
    If MsgBox("This will delete " & selectorCount & " records from " & tbl & ". Continue?", _
              vbYesNo + vbQuestion, "Confirm Deletion") = vbNo Then
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        Exit Sub
    End If
    
    ' Delete all records using For loop
    rs.MoveFirst
    For i = 1 To selectorCount
        rs.Delete
        rs.MoveNext
    Next
    
    rs.Close
    
    'TARGETS SECTION
    tbl = "Targets"

    Set rs = db.OpenRecordset(tbl, dbOpenDynaset)
    
    ' Count records
    rs.MoveLast
    rs.MoveFirst
    targetCount = rs.recordCount
    
    ' Confirm deletion
    If MsgBox("This will delete " & targetCount & " records from " & tbl & ". Continue?", _
              vbYesNo + vbQuestion, "Confirm Deletion") = vbNo Then
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        Exit Sub
    End If
    
    ' Delete all records using For loop
    rs.MoveFirst
    For i = 1 To targetCount
        rs.Delete
        rs.MoveNext

    Next i
    
    rs.Close
    
    Set rs = Nothing
    Set db = Nothing
    
    MsgBox "Successfully deleted " & selectorCount + targetCount & " records.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

'++++++++++++

Sub ClearLocalData()
    Dim tblArr As Variant, tblName As String
    Dim i As Long
    
    Debug.Print "CLEARING LOCAL DATA" & vbLf & "--------------------"
    
    tblArr = Array("localNorks", "localSelectors", "localTargets", "tempSchema", "tempGWSearchResults", "tempSSearchResults")
    
    'clear everything
    For i = 0 To 5
        tblName = tblArr(i)
        ClearTableList tblName
    Next
    
    Debug.Print "LOCAL DATA CLEARED"
End Sub

'+++++++++++++++++++++

Sub PrintFieldNames(rs As DAO.Recordset)
    Dim fld As DAO.Field
    
    Debug.Print "PRINT FIELD NAMES:"
    For Each fld In rs.Fields
        Debug.Print fld.Name
    Next
    
End Sub