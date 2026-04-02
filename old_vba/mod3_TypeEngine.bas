*****************************************

MOD 3_TypeEngine

*****************************************

Option Compare Database

Option Explicit

'=====[ MODULE-LEVEL CONSTANTS — shared by Check* and Fix* ]=====

Private Const EMAIL_PATTERN As String = "^[a-zA-Z0-9]([a-zA-Z0-9._+-]*[a-zA-Z0-9])?@[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?)*\.[a-zA-Z]{2,}$"

Private Const PHONE_US_PATTERN As String = "^[\s\(]*(\+?1[\s\-\.]?)?[\s\(]*([2-9]\d{2})[\s\)\-\.]*([2-9]\d{2})[\s\-\.]*(\d{4})[\s\)]*$"
Private Const PHONE_INTL_PATTERN As String = "^\+(\d{1,3})[\s\-\.]?(\d[\d\s\-\.]{6,14})$"

Private Const IP_PATTERN_V4 As String = "^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
Private Const IP_PATTERN_V6 As String = "^(([0-9a-fA-F]{1,4}:){7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|::)$"

Private Const NAME_PATTERN As String = "^[^\d#_/\\?\[\]@]{1,50}(\s[^\d#_/\\?\[\]@]{1,50}){0,3}$"

'Address state pattern — computed once on first use; module-level to avoid rebuilding per call
Private m_AddressStatePattern As String

Private Function GetAddressStatePattern() As String
    If m_AddressStatePattern = "" Then
        m_AddressStatePattern = "^(.+?),?\s+([^,]+?),?\s+(" & BuildStatePattern() & ")\s*(\d{5}(?:-\d{4})?)?$"
    End If
    GetAddressStatePattern = m_AddressStatePattern
End Function


'=====[ EMAIL ]=====

Function CheckEmail(str As String) As Boolean
    Dim regex As Object
    Dim emailStr As String

    emailStr = Trim(str)
    If emailStr = "" Then CheckEmail = False: Exit Function

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = EMAIL_PATTERN

    CheckEmail = regex.Test(emailStr)
End Function

Function FixEmailStr(str As String) As String
    Dim regex As Object
    Dim inputStr As String, userStr As String, domainStr As String
    Dim i As Long, x As Long

    inputStr = Trim(LCase(str))
    If inputStr = "" Then FixEmailStr = "": Exit Function

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = EMAIL_PATTERN

    'no match — return "" (not silent passthrough)
    If Not regex.Test(inputStr) Then FixEmailStr = "": Exit Function

    'split into local and domain parts
    i = InStr(inputStr, "@")
    userStr = Left(inputStr, i - 1)
    domainStr = Mid(inputStr, i)

    'remove dots from Gmail addresses
    If domainStr = "@gmail.com" Or domainStr = "@googlemail.com" Then
        userStr = Replace(userStr, ".", "")
    End If

    'remove plus addressing
    x = InStr(userStr, "+")
    If x > 0 Then userStr = Left(userStr, x - 1)

    FixEmailStr = Trim(userStr & domainStr)
End Function


'=====[ PHONE ]=====

Function CheckPhone(str As String) As Boolean
    Dim regex As Object
    Dim phoneStr As String

    phoneStr = Trim(str)
    If phoneStr = "" Then CheckPhone = False: Exit Function

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True

    'US/Canada
    regex.Pattern = PHONE_US_PATTERN
    If regex.Test(phoneStr) Then CheckPhone = True: Exit Function

    'International
    regex.Pattern = PHONE_INTL_PATTERN
    If regex.Test(phoneStr) Then CheckPhone = True: Exit Function

    CheckPhone = False
End Function

Function FixPhoneStr(str As String) As String
    Dim regex As Object, matches As Object
    Dim inputStr As String, returnStr As String

    inputStr = Trim(str)
    If inputStr = "" Then FixPhoneStr = "": Exit Function

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True

    'US/Canada: normalize to NXX-NXX-XXXX
    regex.Pattern = PHONE_US_PATTERN
    If regex.Test(inputStr) Then
        Set matches = regex.Execute(inputStr)
        FixPhoneStr = matches(0).SubMatches(1) & "-" & matches(0).SubMatches(2) & "-" & matches(0).SubMatches(3)
        Exit Function
    End If

    'International: normalize to +CC NNNNNNNN
    regex.Pattern = PHONE_INTL_PATTERN
    If regex.Test(inputStr) Then
        Set matches = regex.Execute(inputStr)
        returnStr = matches(0).SubMatches(1)
        returnStr = Replace(returnStr, " ", "")
        returnStr = Replace(returnStr, "-", "")
        returnStr = Replace(returnStr, ".", "")
        FixPhoneStr = "+" & matches(0).SubMatches(0) & " " & returnStr
        Exit Function
    End If

    'no match — return ""
    FixPhoneStr = ""
End Function


'=====[ IP ]=====

Function CheckIP(str As String) As Boolean
    Dim regex As Object
    Dim ipStr As String

    ipStr = Trim(str)
    If ipStr = "" Then CheckIP = False: Exit Function

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True

    regex.Pattern = IP_PATTERN_V4
    If regex.Test(ipStr) Then CheckIP = True: Exit Function

    regex.Pattern = IP_PATTERN_V6
    If regex.Test(ipStr) Then CheckIP = True: Exit Function

    CheckIP = False
End Function

Function FixIPStr(str As String) As String
    Dim regex As Object
    Dim inputStr As String

    inputStr = Trim(str)
    If inputStr = "" Then FixIPStr = "": Exit Function

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True

    'valid = canonical for IP; no normalization needed
    regex.Pattern = IP_PATTERN_V4
    If regex.Test(inputStr) Then FixIPStr = inputStr: Exit Function

    regex.Pattern = IP_PATTERN_V6
    If regex.Test(inputStr) Then FixIPStr = inputStr: Exit Function

    'no match — return ""
    FixIPStr = ""
End Function


'=====[ ADDRESS ]=====

Function CheckAddress(str As String) As Boolean
    Dim regex As Object
    Dim naughtyArr As Variant, wordArr As Variant
    Dim addressStr As String
    Dim wordCount As Long
    Dim i As Long

    addressStr = Trim(str)

    'needs both numbers AND letters
    If FindNumericChar(addressStr) = 0 Or FindLetterChar(addressStr) = 0 Then CheckAddress = False: Exit Function

    'naughty chars
    naughtyArr = Array("_", "\", "?", "!", "@")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        If InStr(addressStr, naughtyArr(i)) > 0 Then CheckAddress = False: Exit Function
    Next

    'must have a street term
    If CheckStreet(addressStr) = False Then CheckAddress = False: Exit Function

    '1 to 15 words
    wordArr = Split(addressStr, " ")
    wordCount = UBound(wordArr) - LBound(wordArr) + 1
    If wordCount < 1 Or wordCount > 15 Then CheckAddress = False: Exit Function

    'must match state pattern (US only)
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = False
    regex.Pattern = GetAddressStatePattern()

    If Not regex.Test(addressStr) Then CheckAddress = False: Exit Function

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

Function FixAddressStr(str As String) As String
    Dim regex As Object, matches As Object
    Dim inputStr As String, zipStr As String, stateStr As String, returnStr As String

    inputStr = Trim(str)
    If inputStr = "" Then FixAddressStr = "": Exit Function

    'consistent with CheckAddress: must have a street term
    If CheckStreet(inputStr) = False Then FixAddressStr = "": Exit Function

    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = False
    regex.Pattern = GetAddressStatePattern()

    'no match — return ""
    If Not regex.Test(inputStr) Then FixAddressStr = "": Exit Function

    Set matches = regex.Execute(Trim(inputStr))

    If IsError(matches(0).SubMatches(3)) Then FixAddressStr = "": Exit Function

    zipStr = matches(0).SubMatches(3)

    'standardize to two letter state
    stateStr = matches(0).SubMatches(2)
    If Len(stateStr) > 2 Then stateStr = StateMap(stateStr)

    'street city, state zip
    returnStr = Trim(matches(0).SubMatches(0) & " " & matches(0).SubMatches(1) & ", " & stateStr & " " & zipStr)

    FixAddressStr = returnStr
End Function


'=====[ LINKEDIN ]=====

Function CheckLinkedIn(str As String) As Boolean
    Dim linkedinStr As String, naughtyArr As Variant
    Dim i As Long

    linkedinStr = Trim(UCase(str))

    If Len(linkedinStr) < 5 Or Len(linkedinStr) > 40 Then CheckLinkedIn = False: Exit Function

    naughtyArr = Array("!", ",", "?", " ", "\", "+")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        If InStr(linkedinStr, naughtyArr(i)) > 0 Then CheckLinkedIn = False: Exit Function
    Next

    If InStr(linkedinStr, "LINKEDIN.COM") > 0 Then CheckLinkedIn = True: Exit Function

    CheckLinkedIn = False
End Function

Function FixLinkedInStr(str As String) As String
    Dim inputStr As String, handleStr As String, naughtyArr As Variant
    Dim inPos As Long
    Dim i As Long

    inputStr = Trim(LCase(str))
    If inputStr = "" Then FixLinkedInStr = "": Exit Function

    'strip forbidden chars (consistent with CheckLinkedIn)
    naughtyArr = Array("!", ",", "?", " ", "\", "+")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        inputStr = Replace(inputStr, naughtyArr(i), "")
    Next

    'must contain linkedin.com
    If InStr(inputStr, "linkedin.com") = 0 Then FixLinkedInStr = "": Exit Function

    'normalize URL to handle-only form: https://linkedin.com/in/johnsmith → linkedin.com/in/johnsmith
    inPos = InStr(inputStr, "linkedin.com/in/")
    If inPos > 0 Then
        handleStr = Mid(inputStr, inPos + 16)  '16 = Len("linkedin.com/in/")
        If InStr(handleStr, "/") > 0 Then handleStr = Left(handleStr, InStr(handleStr, "/") - 1)
        If InStr(handleStr, "?") > 0 Then handleStr = Left(handleStr, InStr(handleStr, "?") - 1)
        inputStr = "linkedin.com/in/" & handleStr
    End If

    FixLinkedInStr = inputStr
End Function


'=====[ GITHUB ]=====

Function CheckGitHub(str As String) As Boolean
    Dim githubStr As String, naughtyArr As Variant
    Dim i As Long

    'use LCase consistently (fixed: original had LCase then UCase overwrite)
    githubStr = Trim(LCase(str))

    If Len(githubStr) < 6 Or Len(githubStr) > 40 Then CheckGitHub = False: Exit Function

    naughtyArr = Array("!", ",", "?", " ", "\", "+", """")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        If InStr(githubStr, naughtyArr(i)) > 0 Then CheckGitHub = False: Exit Function
    Next

    If InStr(githubStr, "github.com") > 0 Then CheckGitHub = True: Exit Function

    CheckGitHub = False
End Function

Function FixGitHubStr(str As String) As String
    Dim inputStr As String, handleStr As String, naughtyArr As Variant
    Dim inPos As Long
    Dim i As Long

    inputStr = Trim(LCase(str))
    If inputStr = "" Then FixGitHubStr = "": Exit Function

    'strip forbidden chars (consistent with CheckGitHub)
    naughtyArr = Array("!", ",", "?", " ", "\", "+", """")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        inputStr = Replace(inputStr, naughtyArr(i), "")
    Next

    'must contain github.com
    If InStr(inputStr, "github.com") = 0 Then FixGitHubStr = "": Exit Function

    'normalize URL to handle-only form: https://github.com/user → github.com/user
    inPos = InStr(inputStr, "github.com/")
    If inPos > 0 Then
        handleStr = Mid(inputStr, inPos + 11)  '11 = Len("github.com/")
        If InStr(handleStr, "/") > 0 Then handleStr = Left(handleStr, InStr(handleStr, "/") - 1)
        If InStr(handleStr, "?") > 0 Then handleStr = Left(handleStr, InStr(handleStr, "?") - 1)
        inputStr = "github.com/" & handleStr
    End If

    FixGitHubStr = inputStr
End Function


'=====[ TELEGRAM ]=====

Function CheckTelegram(str As String) As Boolean
    Dim regex As Object
    Dim tgStr As String, usernameStr As String, naughtyArr As Variant
    Dim i As Long

    tgStr = Trim(LCase(str))

    If Len(tgStr) < 5 Or Len(tgStr) > 40 Then CheckTelegram = False: Exit Function

    'username pattern: @handle (letters, numbers, underscores; 5-32 chars after @)
    If Left(tgStr, 1) = "@" Then
        usernameStr = Mid(tgStr, 2)

        naughtyArr = Array("!", "#", ",", ".", "?", " ", "/", "\", "+", """")
        For i = LBound(naughtyArr) To UBound(naughtyArr)
            If InStr(usernameStr, naughtyArr(i)) > 0 Then CheckTelegram = False: Exit Function
        Next

        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = False
        regex.IgnoreCase = True
        regex.Pattern = "^[a-z0-9_]{5,32}$"

        CheckTelegram = regex.Test(usernameStr)
        Exit Function
    End If

    'numeric ID: pure digits, 5-15 chars
    If IsAllNumbers(tgStr) And Len(tgStr) >= 5 And Len(tgStr) <= 15 Then
        CheckTelegram = True
        Exit Function
    End If

    CheckTelegram = False
End Function

Function FixTelegramStr(str As String) As String
    Dim tgStr As String

    tgStr = Trim(LCase(str))
    If tgStr = "" Then FixTelegramStr = "": Exit Function

    'numeric ID: return as-is if valid length
    If IsAllNumbers(tgStr) Then
        If Len(tgStr) >= 5 And Len(tgStr) <= 15 Then
            FixTelegramStr = tgStr
        Else
            FixTelegramStr = ""
        End If
        Exit Function
    End If

    'username: ensure @ prefix then validate
    If Left(tgStr, 1) <> "@" Then tgStr = "@" & tgStr

    If Not CheckTelegram(tgStr) Then FixTelegramStr = "": Exit Function

    FixTelegramStr = tgStr
End Function


'=====[ DISCORD ]=====

Function CheckDiscord(str As String) As Boolean
    Dim discordStr As String, naughtyArr As Variant
    Dim i As Long

    discordStr = Trim(LCase(str))

    If Len(discordStr) < 5 Or Len(discordStr) > 40 Then CheckDiscord = False: Exit Function

    'look for @ at start or # not at start (handle#discriminator format)
    If InStr(discordStr, "@") <> 1 And InStr(discordStr, "#") < 2 Then CheckDiscord = False: Exit Function

    naughtyArr = Array("!", ",", ".", "?", " ", "%", "$", "/", "\", "+", """")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        If InStr(discordStr, naughtyArr(i)) > 0 Then CheckDiscord = False: Exit Function
    Next

    CheckDiscord = True
End Function

Function FixDiscordStr(str As String) As String
    Dim discordStr As String, prefixChar As String, naughtyArr As Variant
    Dim i As Long

    discordStr = Trim(LCase(str))
    If discordStr = "" Then FixDiscordStr = "": Exit Function

    'preserve @ or # prefix before stripping
    prefixChar = ""
    If Left(discordStr, 1) = "@" Or Left(discordStr, 1) = "#" Then
        prefixChar = Left(discordStr, 1)
        discordStr = Mid(discordStr, 2)
    End If

    'strip forbidden chars
    naughtyArr = Array("!", ",", ".", "?", " ", "%", "$", "/", "\", "+", """")
    For i = LBound(naughtyArr) To UBound(naughtyArr)
        discordStr = Replace(discordStr, naughtyArr(i), "")
    Next

    discordStr = prefixChar & discordStr

    If Not CheckDiscord(discordStr) Then FixDiscordStr = "": Exit Function

    FixDiscordStr = discordStr
End Function


'=====[ NAME ]=====

Function CheckName(str As String) As Boolean
    Dim regex As Object
    Dim nameStr As String

    nameStr = Trim(LCase(str))
    If nameStr = "" Then CheckName = False: Exit Function

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = NAME_PATTERN

    CheckName = regex.Test(nameStr)
End Function

Function FixNameStr(str As String) As String
    Dim regex As Object
    Dim nameStr As String

    nameStr = Trim(LCase(str))
    If nameStr = "" Then FixNameStr = "": Exit Function

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = NAME_PATTERN

    'no match — return ""
    If Not regex.Test(nameStr) Then FixNameStr = "": Exit Function

    'remove apostrophes
    nameStr = Replace(nameStr, "'", "")

    FixNameStr = nameStr
End Function


'=====[ OTHER ]=====

'No CheckOther — "other" is the catch-all fallback type
'for telegram, linkedin, github, discord: now have type-specific Fix functions above
Function FixOtherStr(str As String) As String
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


'=====[ DISPATCH FUNCTIONS ]=====

Function DetectSelectorType(str As String, Optional selectorType As String = "") As String
    Dim detectOrderArr As Variant, inputStr As String, functionName As String
    Dim typeDetect As Boolean
    Dim i As Long

    inputStr = Trim(str)

    If inputStr = "" Or LCase(inputStr) = "null" Then DetectSelectorType = "NULL": Exit Function

    'type undefined — auto-detect via priority order
    If selectorType = "" Or InStr(LCase(selectorType), "row") > 0 Then
        detectOrderArr = DefineDetectOrderArr
        For i = LBound(detectOrderArr) To UBound(detectOrderArr)
            functionName = DetectFunctionMap(Trim(detectOrderArr(i)))
            typeDetect = Application.Run(functionName, inputStr)
            If typeDetect = True Then DetectSelectorType = detectOrderArr(i): Exit Function
        Next

        DetectSelectorType = "NULL": Exit Function
    End If

    'type defined — validate it
    functionName = DetectFunctionMap(selectorType)
    If Trim(functionName) = "" Then DetectSelectorType = selectorType: Exit Function  'type "other" has no checker

    If Application.Run(functionName, inputStr) = False Then DetectSelectorType = "WRONG": Exit Function

    DetectSelectorType = selectorType
End Function

Function BuildSelectorClean(str As String, Optional selectorType As String = "") As String
    Dim inputStr As String, selectorTypeStr As String

    inputStr = Trim(str)
    selectorTypeStr = Trim(LCase(selectorType))

    'detect type if not provided
    If selectorTypeStr = "" Or selectorTypeStr = "row" Then
        selectorTypeStr = DetectSelectorType(inputStr)
    End If

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

    Case "linkedin"
        BuildSelectorClean = FixLinkedInStr(inputStr)

    Case "github"
        BuildSelectorClean = FixGitHubStr(inputStr)

    Case "telegram"
        BuildSelectorClean = FixTelegramStr(inputStr)

    Case "discord"
        BuildSelectorClean = FixDiscordStr(inputStr)

    Case Else
        BuildSelectorClean = Trim(inputStr)

    End Select

End Function


'=====[ INPUT PROCESSING (from mod3a) ]=====

Function CleanDelimInput(str As String, delim As String) As String
    Dim delimArr As Variant, inputStr As String, delimStr As String, delimItem As String
    Dim i As Long

    inputStr = Trim(LCase(str))
    delimStr = Trim(LCase(delim))

    If Trim(LCase(delimStr)) = "new row" Then delimStr = Chr(13)

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
    Dim rowStr As String, rowItem As String
    Dim itemStr As String, cleanStr As String, returnStr As String
    Dim i As Long, j As Long, k As Long

    delimStr = Trim(delim)

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
                'remove trailing delim
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
    Dim delimArr As Variant, cleanArr() As String
    Dim inputStr As String, delimItem As String, cleanItem As String, returnStr As String
    Dim delimHit As Boolean
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

'converts name / address into display defaults
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


'=====[ DETECTION UTILITIES (from mod3b) ]=====

Function DetectInputDelimiter(str As String) As String
    Dim rowArr() As String, hitArr() As String, itemArr() As String, delimArr As Variant
    Dim inputStr As String, rowStr As String, rowItem As String, delimItem As String
    Dim hitStr As String, hitItem As String
    Dim rowCount As Long
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

    'larger inputs: check frequency and consistency
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
            'at least 75% consistency and avg > 1
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

Function DetectStrDiff(strA As String, strB As String, Optional delim As String = "!!") As String
    Dim arrA() As String, arrB() As String
    Dim itemA As String, itemB As String
    Dim arrMatch As Boolean, returnStr As String
    Dim i As Long, j As Long

    arrA = Split(strA, delim)
    arrB = Split(strB, delim)

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

    rs.MoveLast
    rs.MoveFirst
    recordCount = rs.RecordCount

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

    rs.MoveLast
    rs.MoveFirst

    returnStr = ""
    For i = 1 To colMax
        colTypeStr = ""
        colItem = "ColumnType" & i

        rs.MoveFirst
        For j = 1 To rs.RecordCount
            typeItem = Nz(Trim(rs.Fields(colItem).Value), "NULL")
            If typeItem <> "" Then
                typeItem = FixTypesDisplay(typeItem)
                colTypeStr = colTypeStr & typeItem & "!!"
            End If
            rs.MoveNext
        Next
        'remove trailing delim
        colTypeStr = Trim(Left(colTypeStr, Len(colTypeStr) - 2))

        returnStr = returnStr & DetectSchemaTypePercentage(colTypeStr) & "+++"
    Next

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
        If Trim(LCase(k)) <> "null" And Trim(LCase(k)) <> "" Then
            If dict(k) > maxCount Then
                maxCount = dict(k)
                mostCommonStr = k
            End If
        End If
    Next

    totalCount = UBound(arr) + 1

    mostCommonPercent = Round((maxCount / totalCount) * 100, 2)

    DetectSchemaTypePercentage = Trim(mostCommonPercent & "% " & mostCommonStr)
End Function

'check S token
Function CheckToken(str As String) As String
    Dim token As String, res As String
    Dim apiTest1 As String, apiTest2 As String, apiTest3 As String, apiTest4 As String

    token = Trim(str)

    'null input / wrong format
    If token = "" Or LCase(Trim(token)) = Trim(LCase(DefineFormDefaults("txtToken"))) Then ThrowError 1966, "Token Input: " & token: Exit Function
    If Len(token) < 20 Then ThrowError 1967, "Token Input: " & token: Exit Function

    'token test apis
    apiTest1 = "https://S-api.Fnet.F/services/externalservice/api/Lookups/Countries/v1"
    apiTest2 = "https://S-api.Fnet.F/services/externalservice/api/Lookups/CountriesDetails/v1"
    apiTest3 = "https://S-api.Fnet.F/services/externalservice/api/Lookups/Divisions/v1"
    apiTest4 = "https://S-api.Fnet.F/services/externalservice/api/CaseClassifications/Divisions/v1"""

    'test 1
    res = WinINetReq(apiTest1, token)
    If UCase(Trim(Left(res, InStr(res, " ") - 1))) <> "FORBIDDEN" And InStr(UCase(res), "HTTPSENDREQUEST FAILED") = 0 Then CheckToken = token: Exit Function

    '2
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


'=====[ CHAR / STRING UTILITIES (from mod3b) ]=====

'Looks for first numeric character in a string
'@returns: placement of first numeric char; 0 if none exist
Function FindNumericChar(str As String) As Long
    Dim char As String, i As Long

    For i = 1 To Len(Trim(str))
        char = Mid(Trim(str), i, 1)
        If IsNumeric(char) Then FindNumericChar = i: Exit Function
    Next

    FindNumericChar = 0
End Function

'Looks for first letter character in a string
'@returns: placement of first letter char; 0 if none exist
Function FindLetterChar(str As String) As Long
    Dim char As String, ascii As Long, i As Long

    For i = 1 To Len(Trim(str))
        char = Mid(Trim(str), i, 1)
        ascii = AscW(char)

        If ascii > 64 And ascii < 91 Then FindLetterChar = i: Exit Function
        If ascii > 96 And ascii < 123 Then FindLetterChar = i: Exit Function
    Next

    FindLetterChar = 0

End Function

Function IsAllLetters(str As String) As Boolean
    Dim char As String, ascii As Long, i As Long

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
    On Error GoTo FAIL

    Select Case TypeName(obj(k))

        Case "Null", "Empty", "Nothing"
            CheckKeyExists = False: Exit Function

        Case "String"
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
