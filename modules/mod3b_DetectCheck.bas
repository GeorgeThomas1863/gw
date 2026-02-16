*****************************************

MOD 3b_DetectCheck

*****************************************

Option Compare Database

Option Explicit

Function DetectInputDelimiter(str) As String
    Dim rowArr() As String, hitArr() As String, itemArr() As String, delimArr As Variant
    Dim inputStr As String, rowStr As String, rowItem As String, delimItem As String
    Dim hitStr As String, hitItem As String, rowCount As Long
    Dim i As Long, j As Long, t As Long, x As Long, y As Long
    
    inputStr = Trim(LCase(str))
    If inputStr = "" Then DetectInputDelimiter = "New Row": Exit Function

    'split by row (remove extra return)
    rowStr = Trim(Replace(inputStr, Chr(13), ""))
    rowStr = Trim(Replace(rowStr, Chr(10), "+++"))

    rowArr = Split(rowStr, "+++")
    
    rowCount = UBound(rowArr) + 1
    If rowCount = 0 Then ThrowError 1965, "INPUT STR: " & inputStr: Exit Function
    
    'limit to 20
    If rowCount > 20 Then rowCount = 20
    
    'define potential delims
    delimArr = Array(";", "|", vbTab, "+", ",", "/")
    
    'small inputs, hit on first thing
    If rowCount < 4 Then
        For i = LBound(delimArr) To UBound(delimArr)
            delimItem = delimArr(i)
            For j = LBound(rowArr) To UBound(rowArr)
                rowItem = rowArr(j)
                If InStr(rowItem, delimItem) > 0 Then
                    If delimItem = " " Then delimItem = "[space]"
                    
                    DetectInputDelimiter = delimItem
                    Exit Function
                End If
            Next
        Next
    End If
    
    'more complex thing that prob doesnt work
    hitStr = ""
    For i = LBound(delimArr) To UBound(delimArr)
        delimItem = delimArr(i)
        
        x = 0 'avg
        y = 0 'consistency
        For j = LBound(rowArr) To rowCount - 1
            rowItem = rowArr(j)
            t = CountDelimHits(rowItem, delimItem)
            x = x + t
            
            If t > 0 Then y = y + 1
        Next
        If x > 0 Then x = x / rowCount
        If y > 0 Then y = y / rowCount
        
        If delimItem = " " Then delimItem = "[space]"
        hitStr = hitStr & Trim(delimItem) & "$$" & Trim(x) & "$$" & Trim(y) & "!!"
    Next
    
    hitArr = Split(hitStr, "!!")
    
    For i = LBound(hitArr) To UBound(hitArr)
        hitItem = hitArr(i)
        If Trim(hitItem) <> "" Then
            itemArr = Split(hitItem, "$$")
            
            'at least 75%
            If Val(itemArr(2)) > 0.74 And Val(itemArr(1)) > 1 Then DetectInputDelimiter = itemArr(0): Exit Function
        End If
    Next
        
    'if get here just return auto
    DetectInputDelimiter = "New Row"
End Function

Function CountDelimHits(str As String, delim As String) As Long
    Dim inputStr As String
    Dim i As Long, x As Long
    
    inputStr = Trim(LCase(str))
    If inputStr = "" Then CountDelimHits = 0: Exit Function
    
    x = 0
    For i = 1 To Len(inputStr)
        If Mid(inputStr, i, 1) = delim Then x = x + 1
    Next
    
    CountDelimHits = x
End Function

Function DetectSelectorType(str As String, Optional selectorType As String = "") As String
    Dim detectOrderArr As Variant, inputStr As String, functionName As String
    Dim res As String, typeCheck As Boolean, typeDetect As Boolean
    Dim i As Long
    
    inputStr = Trim(str)
    
    If inputStr = "" Or LCase(inputStr) = "null" Then DetectSelectorType = "NULL": Exit Function
    
'    Debug.Print "DETECT SELECTOR TYPE"
'    Debug.Print "SELECTOR INPUT: " & inputStr
'    Debug.Print "SELECTOR TYPE INPUT: " & selectorType
    
    'type undefined
    If selectorType = "" Or InStr(LCase(selectorType), "row") > 0 Then
        detectOrderArr = DefineDetectOrderArr
        For i = LBound(detectOrderArr) To UBound(detectOrderArr)
            functionName = DetectFunctionMap(Trim(detectOrderArr(i)))

            typeDetect = Application.Run(functionName, inputStr)
            'Debug.Print "SELECTOR: " & inputStr & " | " & "FUNCTION: " & functionName & " | " & res
        
            'return hit
            If typeDetect = True Then DetectSelectorType = detectOrderArr(i): Exit Function
        Next
        
        'otherwise cant find
        DetectSelectorType = "NULL": Exit Function
    End If
    
    'handle type defined
    functionName = DetectFunctionMap(selectorType)
    If Trim(functionName) = "" Then DetectSelectorType = selectorType: Exit Function 'type other
        
    typeCheck = Application.Run(functionName, inputStr)
    If typeCheck = False Then DetectSelectorType = "WRONG": Exit Function
        
    'otherwise good
    DetectSelectorType = selectorType
End Function

Function DetectStrDiff(strA As String, strB As String, Optional delim As String = "!!") As String
    Dim arrA() As String, arrB() As String
    Dim itemA As String, itemB As String
    Dim arrMatch As Boolean, returnStr As String
    Dim i As Long, j As Long
    
    arrA = Split(strA, delim)
    arrB = Split(strB, delim)
    
    'If UBound(arrA) = -1 Or UBound(arrB) = -1 Then DetectStrDiff = "": Exit Function
    
    returnStr = ""
    If UBound(arrA) > -1 Then
        For i = LBound(arrA) To UBound(arrA)
            itemA = Trim(arrA(i))
            
            arrMatch = False
            If UBound(arrB) > -1 Then
                For j = LBound(arrB) To UBound(arrB)
                    itemB = Trim(arrB(j))
                    
                    If itemA = itemB Then arrMatch = True
                Next
            End If
            
            If arrMatch = False Then returnStr = returnStr & itemA & "!!"
        Next
    End If
    
    If Trim(returnStr) = "" Then DetectStrDiff = "": Exit Function
    
    'remove trailing delim
    DetectStrDiff = Trim(Left(returnStr, Len(returnStr) - 2))
End Function

'CHECK SELECTOR TYPES

'Check 1, email
Function CheckEmail(str As String) As Boolean
    Dim checkArr() As String, emailStr As String, tld As String
    Dim naughtyArr As Variant
    Dim i As Long
    
    emailStr = Trim(str)
    
    'no @ (and cant start with @)
    If InStr(emailStr, "@") < 2 Then CheckEmail = False: Exit Function
    
    'multiple @
    If CountChar(emailStr, "@") > 1 Then CheckEmail = False: Exit Function
    
    naughtyArr = Array(" ", "?", "/", "\", "#", "+", """")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        If InStr(emailStr, naughtyArr(i)) > 0 Then CheckEmail = False: Exit Function
    Next
    
    'BREAK at @ and search through each side
    checkArr = Split(emailStr, "@")
    
    'must have dot
    If InStr(Trim(checkArr(1)), ".") < 2 Then CheckEmail = False: Exit Function
    
    'check valid tld ending
    tld = Trim(Mid(checkArr(1), InStrRev(checkArr(1), ".") + 1, Len(checkArr(1))))
    If IsAllLetters(tld) = False Or Len(tld) < 2 Then CheckEmail = False: Exit Function
    
    'could add more checks on username portion of email
    
    'otherwise true
    CheckEmail = True
End Function

'Check 2, phone
Function CheckPhone(str As String) As Boolean
    Dim phoneStr As String, phoneStandardStr As String, naughtyArr As Variant
    Dim i As Long
    
    phoneStr = Trim(str)

    'if any letters
    If FindLetterChar(phoneStr) > 0 Then CheckPhone = False: Exit Function
    
    'any @ sign
    If InStr(phoneStr, "@") > 0 Then CheckPhone = False: Exit Function
    
    '+ only at beginning
    If InStr(phoneStr, "+") > 1 Then CheckPhone = False: Exit Function
    
    'only one set of ()
    If InStr(phoneStr, "(") <> 0 And InStr(phoneStr, ")") < 3 Then CheckPhone = False: Exit Function
    If CountChar(phoneStr, "(") > 1 Or CountChar(phoneStr, ")") > 1 Then CheckPhone = False: Exit Function
    
    naughtyArr = Array("@", ".", "?", "/", "\", "_", "%", """")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        If InStr(phoneStr, naughtyArr(i)) > 0 Then CheckPhone = False: Exit Function
    Next
    
    'Debug.Print "PHONE STR: " & phoneStr
    
    'standardize
    phoneStandardStr = FixPhoneStr(phoneStr)
        
    'if wrong length
    If Len(phoneStandardStr) < 5 Or Len(phoneStandardStr) > 17 Then CheckPhone = False: Exit Function
    
    'otherwise true
    CheckPhone = True
End Function

'Check 3, IP
Function CheckIP(str As String) As Boolean
    Dim ipStr As String, naughtyArr As Variant
    Dim ipv4Arr() As String, ipv6Arr() As String
    Dim i As Long, j As Long, k As Long
    
    ipStr = Trim(str)
    
    'must have . or :
    If InStr(ipStr, ".") = 0 And InStr(ipStr, ":") = 0 Then CheckIP = False: Exit Function
    
    If Len(ipStr) < 4 Or Len(ipStr) > 40 Then CheckIP = False: Exit Function
    
    'cannot have anything in naughtyArr
    naughtyArr = Array("@", " ", "?", "/", "\", "#", "+", "-", "_", """")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        If InStr(ipStr, naughtyArr(i)) > 0 Then CheckIP = False: Exit Function
    Next
    
    'ipv4
    If InStr(ipStr, ".") > 0 Then
        'must have exactly 3 dots
        If CountChar(ipStr, ".") <> 3 Then CheckIP = False: Exit Function
        
        'split and check each item
        ipv4Arr = Split(ipStr, ".")
        
        For j = LBound(ipv4Arr) To UBound(ipv4Arr)
            If Len(ipv4Arr(j)) < 1 Or Len(ipv4Arr(j)) > 4 Then CheckIP = False: Exit Function
        Next
        
        'otherwise is ipv4
        CheckIP = True
    End If
    
    'ipv6
    If InStr(ipStr, ":") > 0 Then
        'between 2 and 7 colons
        If CountChar(ipStr, ":") < 2 Or CountChar(ipStr, ":") > 7 Then CheckIP = False: Exit Function
        
        ':: can only appear once
        If CountChar(ipStr, "::") > 1 Then CheckIP = False: Exit Function
        
        'if :: replace to avoid fucking the split
        If CountChar(ipStr, "::") = 1 Then
            ipStr = Replace(ipStr, "::", ":")
        End If
        
        ipv6Arr = Split(ipStr, ":")
        
        For k = LBound(ipv6Arr) To UBound(ipv6Arr)
            If Len(ipv6Arr(k)) < 1 Or Len(ipv6Arr(k)) > 4 Then CheckIP = False: Exit Function
        Next
        
        'otherwise ipv6
        CheckIP = True
    End If
End Function

'Check 4, street address
Function CheckAddress(str As String) As Boolean
    Dim naughtyArr As Variant, wordArr As Variant
    Dim addressStr As String, breakStr As String, wordStr As String
    Dim stateIncluded As Boolean, wordCount As Long
    Dim i As Long, j As Long
    
    addressStr = Trim(str)
    
    'needs both numbers AND letters
    If FindNumericChar(addressStr) = 0 Or FindLetterChar(addressStr) = 0 Then CheckAddress = False: Exit Function
    
    'check for naughty words
    naughtyArr = Array("_", "\", "?", "!", "@")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        If InStr(addressStr, naughtyArr(i)) > 0 Then CheckAddress = False: Exit Function
    Next
    
    'Debug.Print "(((( CHECK STREET " & CheckStreet(addressStr)
    'must have common street names / state
    If CheckStreet(addressStr) = False Then CheckAddress = False: Exit Function
    
    '1 to 15 words
    wordArr = Split(addressStr, " ")
    wordCount = UBound(wordArr) - LBound(wordArr) + 1
    'Debug.Print "WORD COUNT: " & wordCount
    If wordCount < 1 Or wordCount > 15 Then CheckAddress = False: Exit Function
    
    'ONLY HIT ON US ADDRESSES FOR NOW, MAYBE ADD POPUP FOR FOREIGN?
    If CheckState(addressStr) = False Then CheckAddress = False: Exit Function
    
    CheckAddress = True
End Function

'-----------------------------

'ADDRESS UTIL FUNCTIONS
Function CheckStreet(str As String) As Boolean
    Dim streetArr As Variant, wordArr() As String
    Dim streetStr As String, wordStr As String, breakStr As String
    Dim i As Long, j As Long
    
    streetArr = DefineStreetArr()
    
    'break on spaces, commas, and periods
    breakStr = Trim(Replace(Trim(str), " ", "!!"))
    breakStr = Trim(Replace(breakStr, ",", "!!"))
    breakStr = Trim(Replace(breakStr, ".", "!!"))
    wordArr = Split(breakStr, "!!")
    
    'loop through each word
    For i = LBound(wordArr) To UBound(wordArr)
        wordStr = Trim(LCase(wordArr(i)))
        If Trim(wordStr) <> "" Then
        
            For j = LBound(streetArr) To UBound(streetArr)
                streetStr = Trim(LCase(streetArr(j)))
            
                If wordStr = streetStr Then CheckStreet = True: Exit Function
            Next
        End If
    Next
    
    CheckStreet = False
End Function

Function CheckState(str As String) As Boolean
    Dim regex As Object
    Dim inputStr As String, patternStr As String

    inputStr = Trim(str)

    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = False

    patternStr = "^(.+?),?\s+([^,]+?),?\s+(" & BuildStatePattern() & ")\s*(\d{5}(?:-\d{4})?)?$"
    regex.Pattern = patternStr

    If regex.Test(inputStr) Then CheckState = True: Exit Function

    CheckState = False
End Function

'---------------

'Check 5, LinkedIn
Function CheckLinkedIn(str As String) As Boolean
    Dim linkedinStr As String, naughtyArr As Variant
    Dim i As Long
    
    linkedinStr = Trim(UCase(str))
    
    If Len(linkedinStr) < 5 Or Len(linkedinStr) > 40 Then CheckLinkedIn = False: Exit Function
    
    naughtyArr = Array("!", ",", "?", " ", "\", "+")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        If InStr(linkedinStr, naughtyArr(i)) > 0 Then CheckLinkedIn = False: Exit Function
    Next
    
    If InStr(linkedinStr, "linkedin.com") > 0 Then CheckLinkedIn = True: Exit Function
    
    CheckLinkedIn = False
End Function

'Check 6, GitHub
Function CheckGitHub(str As String) As Boolean
    Dim githubStr As String, naughtyArr As Variant
    Dim i As Long
    
    githubStr = Trim(LCase(str))
    
    If Len(githubStr) < 6 Or Len(githubStr) > 40 Then CheckGitHub = False: Exit Function
    
    githubStr = Trim(UCase(str))
    
    naughtyArr = Array("!", ",", "?", " ", "\", "+", """")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        If InStr(githubStr, naughtyArr(i)) > 0 Then CheckGitHub = False: Exit Function
    Next
    
    'need to make less restrictive / stupid
    If InStr(githubStr, "github.com") > 0 Then CheckGitHub = True: Exit Function
    
    CheckGitHub = False
End Function

'Check 7, TG
'starts with an @ 'BUILD WAY TO DETECT TG ID
Function CheckTelegram(str As String) As Boolean
    Dim tgStr As String, naughtyArr As Variant
    Dim i As Long
    
    tgStr = Trim(LCase(str))
    
    If Len(tgStr) < 5 Or Len(tgStr) > 40 Then CheckTelegram = False: Exit Function
    
    'must have @, build way to detect id
    If InStr(tgStr, "@") <> 1 Then CheckTelegram = False: Exit Function
    
    naughtyArr = Array("!", "#", ",", ".", "?", " ", "/", "\", "+", """")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        If InStr(tgStr, naughtyArr(i)) > 0 Then CheckTelegram = False: Exit Function
    Next
    
    CheckTelegram = True
End Function

'Check 8, Discord
Function CheckDiscord(str As String) As Boolean
    Dim discordStr As String, naughtyArr As Variant
    Dim i As Long
    
    discordStr = Trim(LCase(str))
    
    If Len(discordStr) < 5 Or Len(discordStr) > 40 Then CheckDiscord = False: Exit Function
    
    'look for hash or @
    If InStr(discordStr, "@") <> 1 And InStr(discordStr, "#") < 2 Then CheckDiscord = False: Exit Function
    
    naughtyArr = Array("!", ",", ".", "?", " ", "%", "$", "/", "\", "+", """")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        If InStr(discordStr, naughtyArr(i)) > 0 Then CheckDiscord = False: Exit Function
    Next
    
    CheckDiscord = True
End Function

'Check 9, Name (persona name)
Function CheckName(str As String) As Boolean
    Dim nameStr As String, naughtyArr As Variant
    Dim wordArr() As String, wordCount As Long
    Dim i As Long, j As Long
    
    nameStr = Trim(LCase(str))
    'Debug.Print "NAME STR: " & nameStr
    
    'characters not allowed in name
    naughtyArr = Array("#", "_", "/", "\", "?", "[", "]", "@")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        If InStr(nameStr, naughtyArr(i)) > 0 Then CheckName = False: Exit Function
    Next
    
    'no numbers
    If FindNumericChar(nameStr) > 0 Then CheckName = False: Exit Function
    
    '1 to 5 words
    wordArr = Split(nameStr, " ")
    wordCount = UBound(wordArr) - LBound(wordArr) + 1
    'Debug.Print "WORD COUNT: " & wordCount
    If wordCount < 1 Or wordCount > 4 Then CheckName = False: Exit Function
    
    'name too long
    For j = LBound(wordArr) To UBound(wordArr)
        If Len(wordArr(j)) > 50 Then CheckName = False: Exit Function
    Next
    
    CheckName = True
End Function



'+++++++++++++++++++++++++


'---------------------------------

'UTIL check functions

'Looks for first numeric character in a string (mostly to see if str has any)
'@returns: placement of first numeric char; 0 if none exist
Function FindNumericChar(str As String) As Long
    Dim char As String, i As Long
    
    'loop through all characters
    For i = 1 To Len(Trim(str))
        'current character
        char = Mid(Trim(str), i, 1)
        
        'if numeric exit
        If IsNumeric(char) Then FindNumericChar = i: Exit Function
    Next
    
    'otherwise return 0
    FindNumericChar = 0
End Function

'Looks for first numeric character in a string (mostly to see if str has any)
'@returns: placement of first numeric char; 0 if none exist
Function FindLetterChar(str As String) As Long
    Dim char As String, ascii As Long, i As Long
    
    'loop through all characters
    For i = 1 To Len(Trim(str))
        
        'current character
        char = Mid(Trim(str), i, 1)
        ascii = AscW(char)
        
        If ascii > 64 And ascii < 91 Then FindLetterChar = i: Exit Function
        If ascii > 96 And ascii < 123 Then FindLetterChar = i: Exit Function
    Next
    
    'otherwise return 0
    FindLetterChar = 0
    
End Function

Function IsAllLetters(str As String) As Boolean
    Dim char As String, ascii As Long, i As Long
    
    'null input (unnecessary)
    'If Trim(str) = "" Then ThrowError 1959, str: Exit Function
    
    For i = 1 To Len(Trim(str))
        char = Mid(Trim(str), i, 1)
        ascii = AscW(char)
        
        'fail conditions
        If ascii < 65 Or ascii > 122 Then IsAllLetters = False: Exit Function
        If ascii > 90 And ascii < 97 Then IsAllLetters = False: Exit Function
    Next
    
    'otherwise true
    IsAllLetters = True
End Function


Function IsAllNumbers(str As String) As Boolean
    Dim char As String, ascii As Long, i As Long
    
    
    
    'null input (unnecessary)
    'If Trim(str) = "" Then ThrowError 1959, str: Exit Function
    
    For i = 1 To Len(Trim(str))
        char = Mid(Trim(str), i, 1)
        ascii = AscW(char)
        
        'fail conditions
        If ascii < 48 Or ascii > 57 Then IsAllNumbers = False: Exit Function
    Next
    
    'otherwise true
    IsAllNumbers = True

End Function

Function CountChar(str As String, char As String) As Long
    Dim i As Long, k As Long
    
    str = Trim(str)
    char = Trim(char)
    
    'old length
    i = Len(str)
    
    'new length
    k = Len(Replace(str, char, ""))
    
    CountChar = i - k
End Function

Function CheckKeyExists(obj As Object, k As String) As Boolean
    Dim v As Variant
    On Error GoTo FAIL
     
    Select Case TypeName(obj(k))

        Case "Null", "Empty", "Nothing"
            CheckKeyExists = False: Exit Function
        
        Case "String"
        
            'check for empty string
            If Trim(obj(k)) = "" Then
                CheckKeyExists = False: Exit Function
            Else
                CheckKeyExists = True: Exit Function
            End If
        
        Case Else
            CheckKeyExists = True: Exit Function
    
    End Select
    Exit Function
    
FAIL:
    CheckKeyExists = False
End Function


'+++++++++++++++++++++++++++++++++++++++++



Function DetectColMax(str As String) As Long
    Dim arr() As String, rowArr() As String, rowStr As String
    Dim colNum As Long, colMax As Long
    Dim i As Long
    
    arr = Split(Trim(str), "+++")
    
    colMax = 0
    For i = LBound(arr) To UBound(arr)
        rowStr = Trim(arr(i))

        If rowStr <> "" Then
            rowArr = Split(rowStr, "!!")

            'detect max columns
            colNum = UBound(rowArr) + 1
            If colNum > colMax Then colMax = colNum
        End If
    Next
    
    If colMax = 0 Then ThrowError 1962, "PROBLEM DETECTING COL MAX": Exit Function
    
    DetectColMax = colMax
End Function

Function DetectColMaxRS(tbl As String) As Long
    Dim db As DAO.Database, rs As DAO.Recordset, fld As DAO.Field
    Dim recordCount As Long, colMax As Long
    Dim i As Long, x As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(tbl, dbOpenDynaset)
    
    If (rs.EOF And rs.BOF) Then ThrowError 1962, "SCHEMA TABLE EMPTY, CANT GET COLMAX": Exit Function
    
    'shouldnt need if
    rs.MoveFirst
    rs.MoveLast
    recordCount = rs.recordCount
    rs.MoveFirst
    
    'loop through by row
    colMax = 0
    For i = 1 To recordCount
        x = 0
        For Each fld In rs.Fields
            If Not IsNull(fld.Value) And Trim(fld.Value) <> "" Then
                If InStr(fld.Name, "ColumnStr") > 0 Then
                    x = x + 1
                End If
            End If
        Next
        If x > colMax Then colMax = x
        
        If i < recordCount Then rs.MoveNext
    Next
    
    If colMax = 0 Then ThrowError 1962, "PROBLEM DETECTING COL MAX": Exit Function
    
    DetectColMaxRS = colMax
End Function

Function DetectSchemaTypeStr(colMax As Long) As String
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim returnStr As String, colTypeStr As String, colItem As String, typeItem As String
    Dim i As Long, j As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tempSchema", dbOpenDynaset)
    
    If rs.EOF Then DetectSchemaTypeStr = "": Exit Function
    
    rs.MoveFirst
    rs.MoveLast
    rs.MoveFirst

    returnStr = ""
    For i = 1 To colMax
        colTypeStr = ""
        colItem = "ColumnType" & i

        rs.MoveFirst
        For j = 1 To rs.recordCount
            typeItem = Nz(Trim(rs.Fields(colItem).Value), "NULL")
            If typeItem <> "" Then
                typeItem = FixTypesDisplay(typeItem)
                colTypeStr = colTypeStr & typeItem & "!!"
            End If
            rs.MoveNext
        Next
        'remove trailing delim
        colTypeStr = Trim(Left(colTypeStr, Len(colTypeStr) - 2))

        'Debug.Print "COL TYPE STR: " & colTypeStr

        returnStr = returnStr & DetectSchemaTypePercentage(colTypeStr) & "+++"
    Next

    'Debug.Print "SCHEMA TYPE STR: " & returnStr

    'remove trailing delim
    DetectSchemaTypeStr = Trim(Left(returnStr, Len(returnStr) - 3))
End Function

'takes delimited string as input
Function DetectSchemaTypePercentage(str As String) As String
    Dim arr() As String, dict As Object
    Dim k As Variant, typeStr As String, mostCommonStr As String
    Dim maxCount As Long, totalCount As Long, mostCommonPercent As Double
    Dim i As Long
    
    'null input
    If Trim(str) = "" Then ThrowError 1962, "DETECT SCHEMA TYPE NULL INPUT": Exit Function
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    arr = Split(Trim(str), "!!")
    
    'build dict
    For i = LBound(arr) To UBound(arr)
        typeStr = arr(i)
        
        If dict.Exists(typeStr) Then
            dict(typeStr) = dict(typeStr) + 1
        
        Else
            dict.Add typeStr, 1
        End If
    Next
    
    maxCount = 0
    For Each k In dict.Keys
        'Debug.Print "STRING PERCENTAGE VAL: " & Trim(LCase(k))
        If Trim(LCase(k)) <> "null" And Trim(LCase(k)) <> "" Then
            If dict(k) > maxCount Then
                maxCount = dict(k)
                mostCommonStr = k
            End If
        End If
     Next
     
     totalCount = UBound(arr) + 1
     
     mostCommonPercent = Round((maxCount / totalCount) * 100, 2)
     
     'Debug.Print "MAX COUNT: " & maxCount
     'Debug.Print "TOTAL COUNT: " & totalCount
     
     DetectSchemaTypePercentage = Trim(mostCommonPercent & "% " & mostCommonStr)
End Function

'++++++++++++++++++++++++++++++

'check S token
Function CheckToken(str As String) As String
    Dim token As String, res As String
    Dim apiTest1 As String, apiTest2 As String, apiTest3 As String, apiTest4 As String
    
    'logic is test WinINet first
    ' if failed test GetXML should prmompt popup and fail;
    ' then try GetXML immediately again (3rd check, hope to use session from right password)
    ' if that fails try WinINet last time (4th again hoping to use session), NO CLUE how this really works

    token = Trim(str)
    
    'null input / 'wrong format
    If token = "" Or LCase(Trim(token)) = Trim(LCase(DefineFormDefaults("txtToken"))) Then ThrowError 1966, "Token Input: " & token: Exit Function
    If Len(token) < 20 Then ThrowError 1967, "Token Input: " & token: Exit Function 'make more complex if needed
    
    'token test api's
    apiTest1 = "https://S-api.Fnet.F/services/externalservice/api/Lookups/Countries/v1"
    apiTest2 = "https://S-api.Fnet.F/services/externalservice/api/Lookups/CountriesDetails/v1"
    apiTest3 = "https://S-api.Fnet.F/services/externalservice/api/Lookups/Divisions/v1"
    apiTest4 = "https://S-api.Fnet.F/services/externalservice/api/CaseClassifications/Divisions/v1"""

    'test1
    'TestHeaderMethods apiTest1, token
    res = WinINetReq(apiTest1, token)
    If UCase(Trim(Left(res, InStr(res, " ") - 1))) <> "FORBIDDEN" And InStr(UCase(res), "HTTPSENDREQUEST FAILED") = 0 Then CheckToken = token: Exit Function
    'Debug.Print res

    '2 [usually the one that passes, no clue why]
    res = WinINetReq(apiTest2, token)
    If UCase(Trim(Left(res, InStr(res, " ") - 1))) <> "FORBIDDEN" And InStr(UCase(res), "HTTPSENDREQUEST FAILED") = 0 Then CheckToken = token: Exit Function
    
    '3
    res = GetXML(apiTest3, token)
    If UCase(Trim(Left(res, InStr(res, " ") - 1))) <> "FORBIDDEN" And InStr(UCase(res), "HTTPSENDREQUEST FAILED") = 0 Then CheckToken = token: Exit Function
    
    '4
    res = WinINetReq(apiTest4, token)
    If UCase(Trim(Left(res, InStr(res, " ") - 1))) <> "FORBIDDEN" And InStr(UCase(res), "HTTPSENDREQUEST FAILED") = 0 Then CheckToken = token: Exit Function

    CheckToken = ThrowError(1968, token)
End Function
