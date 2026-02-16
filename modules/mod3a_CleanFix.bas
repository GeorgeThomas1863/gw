*****************************************

MOD 3a_CleanFix

*****************************************

Option Compare Database

Option Explicit

Function CleanDelimInput(str As String, delim As String) As String
    Dim delimArr As Variant, inputStr As String, delimStr As String, delimItem As String
    Dim i As Long
    
    inputStr = Trim(LCase(str))
    delimStr = Trim(LCase(delim))
    
    If Trim(LCase(delimStr)) = "new row" Then delimStr = Chr(13)
    
'    Debug.Print InStr(inputStr, Chr(10))
'    Debug.Print InStr(inputStr, Chr(13))
    
    'if auto try lookup again
    If InStr(delimStr, "auto") > 0 Then delimStr = DetectInputDelimiter(inputStr)
    
    'check if one of allowed delims
    delimArr = Array(";", "|", vbTab, "+", ",", "[space]", "/", Chr(13))
    
    For i = LBound(delimArr) To UBound(delimArr)
        delimItem = delimArr(i)
        If delimItem = delimStr Then CleanDelimInput = delimStr: Exit Function
    Next
    
    'if get here, throw error
    ThrowError 1965, "DELIM STR: " & delim: Exit Function

End Function

Function CleanUserInput(str As String, delim As String) As String
    Dim rowArr() As String, itemArr() As String, inputStr As String, delimStr As String
    Dim rowStr As String, rowItem As String, rowCheck As String
    Dim itemStr As String, cleanStr As String, returnStr As String
    Dim i As Long, j As Long, k As Long
    

    delimStr = Trim(delim)
    'Debug.Print "CHR13: " & InStr(delimStr, Chr(13))
    'Debug.Print "CHR10: " & InStr(delimStr, Chr(10))
    
    inputStr = Trim(Replace(Trim(str), Chr(13), ""))
    inputStr = Trim(Replace(inputStr, delimStr, "!!"))
    
    'split by row (remove extra return)
    rowStr = Trim(Replace(inputStr, Chr(13), ""))
    rowStr = Trim(Replace(rowStr, Chr(10), "+++"))
    rowArr = Split(rowStr, "+++")
    
    'loop each row individually, check, reassemble
    returnStr = ""
    For i = LBound(rowArr) To UBound(rowArr)
        rowItem = Trim(rowArr(i))
        
        'check if row empty [aside from delims]
        If Trim(Replace(rowItem, "!!", "")) <> "" Then
        
            'break apart and remove spaces in items
            itemArr = Split(rowItem, "!!")
            
            cleanStr = ""
            For k = LBound(itemArr) To UBound(itemArr)
                itemStr = Trim(itemArr(k))
                
                'set blanks to null HERE
                If itemStr = "" Then itemStr = "NULL"
                
                'remove duplicates HERE
                If itemStr = "NULL" Or InStr(cleanStr, itemStr & "!!") = 0 And InStr(returnStr, itemStr & "!!") = 0 And InStr(returnStr, itemStr & "+++") = 0 Then
                
                    cleanStr = cleanStr & itemStr & "!!"
                End If
            Next
            
            If cleanStr <> "" Then
                'remove traling delim
                cleanStr = Trim(Left(cleanStr, Len(cleanStr) - 2))
                returnStr = returnStr & cleanStr & "+++"
            End If
        End If
    Next
    
    'trailing delim
    CleanUserInput = Trim(Left(returnStr, Len(returnStr) - 3))
End Function


Function CleanTargetInput(str As String) As String
    Dim searchTarget As String
     
    searchTarget = Trim(LCase(str))
    
    'null input
    If searchTarget = "" Then ThrowError 1951, "null input": Exit Function
    
    'pass conditions
    If InStr(searchTarget, "both") > 0 Then CleanTargetInput = "both": Exit Function
    If InStr(searchTarget, "graywolfe") > 0 Then CleanTargetInput = "graywolfe": Exit Function
    If InStr(searchTarget, "S") > 0 Then CleanTargetInput = "S": Exit Function
    
    'otherwise throw error
    ThrowError 1955, searchTarget: Exit Function

End Function

Function CleanSelectorInput(str As String) As String
    Dim selectorTypeArr As Variant, selectorType As String
    Dim i As Long
    
    selectorType = FixTypesInternal(Trim(LCase(str)))
    'Debug.Print "!!!SELECTOR TYPE: " & selectorType

    selectorTypeArr = DefineSelectorTypeArr
    
    For i = LBound(selectorTypeArr) To UBound(selectorTypeArr)
        If Trim(selectorTypeArr(i)) = selectorType Then CleanSelectorInput = selectorType: Exit Function
    Next
    
    'avoid error on nulls
    If LCase(selectorType) = "null" Then Exit Function
    
    'otherwise throw error
    ThrowError 1956, selectorType: Exit Function
End Function

Function CleanImportTypeInput(str As String) As String
    Dim importType As String
    
    importType = Trim(LCase(str))
    
    'pass conditions
    If InStr(importType, "default") > 0 Then CleanImportTypeInput = "default": Exit Function
    If InStr(importType, "unrelated") > 0 Then CleanImportTypeInput = "unrelated": Exit Function
    
    'otherwise throw error
    ThrowError 1957, importType: Exit Function
End Function

Function FixTargetDelimiters(str As String, Optional selectorType As String = "") As String
    Dim delimArr As Variant, cleanArr() As String, inputStr As String
    Dim delimItem As String, delimHit As Boolean, cleanItem As String, returnStr As String
    Dim i As Long
    
    inputStr = Trim(str)
    
    If inputStr = "" Then FixTargetDelimiters = "": Exit Function
    
    If selectorType = "address" Then
        delimArr = Array(";", "|", vbTab, "+")
    
    Else
        delimArr = Array(";", "|", vbTab, "+", ",")
    End If
    
    delimHit = False
    For i = LBound(delimArr) To UBound(delimArr)
        delimItem = delimArr(i)
        If InStr(inputStr, delimItem) > 0 Then
            inputStr = Replace(inputStr, delimItem, "!!")
            delimHit = True
        End If
    Next
    
    If delimHit = False Then FixTargetDelimiters = inputStr: Exit Function
    
    'break apart and reassemble to remove extra spaces
    cleanArr = Split(inputStr, "!!")
    
    returnStr = ""
    For i = LBound(cleanArr) To UBound(cleanArr)
        cleanItem = Trim(cleanArr(i))
        
        If cleanItem <> "" Then returnStr = returnStr & cleanItem & "!!"
    Next
    
    'remove trailing delim
    FixTargetDelimiters = Trim(Left(returnStr, Len(returnStr) - 2))
        
End Function

Function BuildSelectorClean(str As String, Optional selectorType As String = "") As String
    Dim inputStr As String, selectorTypeStr As String
    
    '"email", "phone", "ip", "address", "linkedin", "github", "telegram", "discord", "name"
    inputStr = Trim(str)
    selectorTypeStr = Trim(LCase(selectorType))
    
    'Debug.Print "%%% TYPE SELECT BEGIN: " & selectorTypeStr
    
    'detect type if not provided
    If selectorTypeStr = "" Or selectorTypeStr = "row" Then
        selectorTypeStr = DetectSelectorType(inputStr)
    End If
    
    'Debug.Print "*** TYPE SELECT: " & selectorTypeStr
    
    Select Case selectorTypeStr
    
    Case "phone"
        BuildSelectorClean = FixPhoneStr(inputStr)
        
    Case "email"
        BuildSelectorClean = FixEmailStr(inputStr)
    
    Case "address"
        BuildSelectorClean = FixAddressStr(inputStr)
    
    Case "ip"
        BuildSelectorClean = FixIPStr(inputStr)
    
    Case "name"
        BuildSelectorClean = FixNameStr(inputStr)
    
    Case "telegram", "linkedin", "github", "discord"
        BuildSelectorClean = FixOtherStr(inputStr)
    
    Case Else
        BuildSelectorClean = Trim(inputStr)
    
    End Select
    
    
    
End Function

'------------------------------

Function FixWrongSelectorType(str As String, selectorType As String, problemType As String) As String
    Dim wrongCheck As Variant, skipCheck As Variant
    Dim addStr As String, problemTypeStr As String, wrongText As String
    
    addStr = Trim(str)
    problemTypeStr = Trim(UCase(problemType))
    
    Select Case problemTypeStr
        
    Case "WRONG"
    
        wrongText = DefinePopupText(addStr, selectorType, "typeWrong")
        wrongCheck = MsgBox(wrongText, vbYesNo + vbDefaultButton1, "REALLY?!?")
        
        If wrongCheck = vbYes Then
            FixWrongSelectorType = selectorType: Exit Function
        
        Else
            ThrowError 1998, addStr: Exit Function
        End If
        
    Case "NULL"
    
        wrongText = DefinePopupText(addStr, selectorType, "typeNull")
        skipCheck = MsgBox(wrongText, vbYesNo + vbDefaultButton1, "#FAIL!!1!")
        
        If skipCheck = vbYes Then
            FixWrongSelectorType = "SKIP": Exit Function
        
        Else
            ThrowError 1998, addStr: Exit Function
        End If
        
    Case Else
        FixWrongSelectorType = selectorType
        
    End Select
End Function

'converts name / address into display defaults
Function FixTypesDisplay(str As String) As String
    Dim inputStr As String
    
    inputStr = Trim(LCase(str))
    
    Select Case inputStr
    
    Case "name"
        FixTypesDisplay = "Persona Name": Exit Function
    
    Case "address"
        FixTypesDisplay = "Street Address": Exit Function
        
    Case "ip"
        FixTypesDisplay = "IP": Exit Function
        
    Case ""
        FixTypesDisplay = "NULL": Exit Function
        
    Case Else
        FixTypesDisplay = StrConv(inputStr, vbProperCase): Exit Function
    
    End Select

End Function

'converts from display defaults to internal, removes plurals
Function FixTypesInternal(str As String) As String
    Dim inputStr As String
    
    inputStr = Trim(LCase(str))
    
    If InStr(inputStr, "auto detect") > 0 Or Trim(inputStr) = "row" Then inputStr = "row"
    If InStr(inputStr, "name") > 0 Then inputStr = "name"
    If InStr(inputStr, "address") > 0 Then inputStr = "address"
    If inputStr = "ips" Then inputStr = "ip"
    If inputStr = "phones" Then inputStr = "phone"
    If inputStr = "emails" Then inputStr = "email"
    
    FixTypesInternal = inputStr
End Function

'standardizes phone inputs
Function FixPhoneStr(str) As String
    Dim regex As Object, matches As Object
    Dim inputStr As String, returnStr As String
    Dim i As Long
    
    'null input
    inputStr = Trim(str)
    If inputStr = "" Then FixPhoneStr = "": Exit Function
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    
    ' Pattern for US/Canada numbers (with optional +1 or 1 prefix) per claude
    regex.Pattern = "^[\s\(]*(\+?1[\s\-\.]?)?[\s\(]*([2-9]\d{2})[\s\)\-\.]*([2-9]\d{2})[\s\-\.]*(\d{4})[\s\)]*$"
    
    If regex.Test(inputStr) Then
        Set matches = regex.Execute(inputStr)
        
        FixPhoneStr = matches(0).SubMatches(1) & "-" & matches(0).SubMatches(2) & "-" & matches(0).SubMatches(3)
        Exit Function
    End If
    
    'International numbers per claude; Matches: +44 20 7123 4567, +33 1 42 86 82 00, etc.
    regex.Pattern = "^\+(\d{1,3})[\s\-\.]?(\d[\d\s\-\.]{6,14})$"
    
    If regex.Test(inputStr) Then
        Set matches = regex.Execute(inputStr)
        
        'remove spaces / dashes from number portion
        returnStr = matches(0).SubMatches(1)
        
        returnStr = Replace(returnStr, " ", "")
        returnStr = Replace(returnStr, "-", "")
        returnStr = Replace(returnStr, ".", "")
        
        FixPhoneStr = "+" & matches(0).SubMatches(0) & " " & returnStr
        Exit Function
    End If
    
    'otherwise NO hits on regex, return input
    FixPhoneStr = inputStr
End Function

Function FixEmailStr(str As String) As String
    Dim regex As Object, matches As Object
    Dim inputStr As String, returnStr As String, userStr As String, domainStr As String
    Dim i As Long, x As Long

    'lcase all emails
    inputStr = Trim(LCase(str))
    If inputStr = "" Then FixEmailStr "": Exit Function

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    
    ' Email validation regex pattern per claude
    regex.Pattern = "^[a-zA-Z0-9]([a-zA-Z0-9._+-]*[a-zA-Z0-9])?@[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?)*\.[a-zA-Z]{2,}$"
    
    ' if not email just return input
    If Not regex.Test(inputStr) Then FixEmailStr = inputStr: Exit Function
    
    ' Split email into local and domain parts
    i = InStr(inputStr, "@")
    userStr = Left(inputStr, i - 1)
    domainStr = Mid(inputStr, i)
    
    ' Remove dots from Gmail addresses (gmail.com and googlemail.com)
    If domainStr = "@gmail.com" Or domainStr = "@googlemail.com" Then
        userStr = Replace(userStr, ".", "")
    End If
    
    ' Remove plus addressing (everything after +)
    x = InStr(userStr, "+")
    If x > 0 Then userStr = Left(userStr, x - 1)
    
    ' Return normalized email
    FixEmailStr = Trim(userStr & domainStr)
End Function

Function FixAddressStr(str As String) As String
    Dim regex As Object, matches As Object
    Dim inputStr As String, zipStr As String, stateStr As String, patternStr As String, returnStr As String
    
    inputStr = Trim(str)
    If inputStr = "" Then FixAddressStr = "": Exit Function
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = False
    
    patternStr = "^(.+?),?\s+([^,]+?),?\s+(" & BuildStatePattern() & ")\s*(\d{5}(?:-\d{4})?)?$"
    regex.Pattern = patternStr
    
    'no hits, just return
    If Not regex.Test(inputStr) Then
        FixAddressStr = inputStr: Exit Function
    End If
    
    Set matches = regex.Execute(Trim(inputStr))
    
    zipStr = ""
    
    If IsError(matches(0).SubMatches(3)) Then Exit Function
    
    zipStr = matches(0).SubMatches(3)
  
    
    'standardize to two letter
    stateStr = matches(0).SubMatches(2)
    If Len(stateStr) > 2 Then stateStr = StateMap(stateStr)
    
    'Build result: street city, state zip
    returnStr = Trim(matches(0).SubMatches(0) & " " & matches(0).SubMatches(1) & ", " & stateStr & " " & zipStr)
    
    FixAddressStr = returnStr

    Set regex = Nothing
End Function

Function FixIPStr(str As String) As String
    Dim regex As Object
    Dim inputStr As String, returnStr As String
    
    inputStr = Trim(str)
    If inputStr = "" Then FixIPStr = "": Exit Function
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = False
    
    'IPv4 per claude
    regex.Pattern = "^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
    
    If regex.Test(inputStr) Then FixIPStr = inputStr: Exit Function
    
    ' IPv6 per claude, matches: full format: 2001:0db8:85a3:0000:0000:8a2e:0370:7334, compressed: 2001:db8:85a3::8a2e:370:7334, localhost: ::1
    regex.Pattern = "^(([0-9a-fA-F]{1,4}:){7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|::)$"
    
    If regex.Test(inputStr) Then FixIPStr = inputStr: Exit Function

    'otherwise just return input
    FixIPStr = inputStr
End Function

' lcase everything, remove dashes
Function FixNameStr(str) As String
    Dim returnStr As String
    
    returnStr = Trim(LCase(str))
    If returnStr = "" Then FixNameStr = "": Exit Function
    
    returnStr = Replace(returnStr, "'", "")
    
    FixNameStr = returnStr
End Function

'for telegram, linkedin, github, discord
Function FixOtherStr(str) As String
    Dim naughtyArr As Variant, naughtyItem As String, returnStr As String
    Dim i As Long
    
    returnStr = Trim(LCase(str))
    
    naughtyArr = Array("+", "!", "@", "#")
    
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        naughtyItem = naughtyArr(i)
        returnStr = Trim(Replace(returnStr, naughtyItem, ""))
    Next
    
    FixOtherStr = returnStr
End Function


Function FixPlural(str As String) As String
    Dim wordStr As String, lastChar As String, lastTwo As String
    
    wordStr = Trim(str)
    
    lastChar = Right(wordStr, 1)
    lastTwo = Right(wordStr, 2)
    
    If lastChar = "s" Or lastChar = "x" Or lastChar = "z" Then FixPlural = Trim(wordStr & "es"): Exit Function
    
    If lastTwo = "ch" Or lastTwo = "sh" Then FixPlural = Trim(wordStr & "es"): Exit Function
    
    FixPlural = Trim(wordStr & "s")
End Function

'compiles multiple returns into single json str (adds returns to items array)
Function FixSBulkReturn(str As String) As String
    Dim arr() As String
    Dim itemStr As String, dataStr As String, returnStr As String
    Dim startPos As Long, stopPos As Long
    Dim i As Long
    
    If Trim(str) = "" Then FixSBulkReturn = str: Exit Function
    
    arr = Split(str, "!;")
    
    returnStr = ""
    For i = LBound(arr) To UBound(arr)
        itemStr = Trim(arr(i))
        
        If i = 0 Then
            stopPos = InStr(itemStr, "}],""facets")
            dataStr = Trim(Left(itemStr, stopPos))
        
        Else
            stopPos = InStr(itemStr, "}],""facets")
            startPos = InStr(itemStr, ",""items"":[") + 10
            dataStr = Trim(Mid(itemStr, startPos, stopPos - startPos + 1))
        End If
        
        returnStr = returnStr & dataStr & ","
    Next
    
    returnStr = Trim(Left(returnStr, Len(returnStr) - 1))
    
    FixSBulkReturn = returnStr & "]}"
    
End Function

'changes iso date so vba sees as date
Function FixDate(str As String) As Date
    Dim dateStr As String
    
    dateStr = Replace(Trim(str), "T", " ")
    dateStr = Trim(Replace(dateStr, "Z", ""))
    dateStr = Trim(Left(dateStr, Len(dateStr) - 4))
    'Debug.Print "DATE STR: " & dateStr
    
    FixDate = CDate(dateStr)
End Function

'uses some built in ref library might need to import System.Collections
Function AlphabetizeStr(str As String, delim As String) As String
    Dim obj As Object, arr() As String
    Dim i As Long
    
    arr = Split(Trim(str), delim)
    
    Set obj = CreateObject("System.Collections.ArrayList")
    
    For i = LBound(arr) To UBound(arr)
        obj.Add arr(i)
    Next
    
    obj.Sort
    
    AlphabetizeStr = Join(obj.ToArray, delim)
    
End Function
