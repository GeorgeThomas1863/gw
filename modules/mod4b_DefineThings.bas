*****************************************

MOD 4b_DefineThings

*****************************************

Option Compare Database

Option Explicit


#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
'below should be fine
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'escapes single quotes for SQL string concatenation
Function SqlSafe(str As String) As String
    SqlSafe = Replace(str, "'", "''")
End Function

'builds state regex alternation from DefineStateArr (used by CheckState and FixAddressStr)
Function BuildStatePattern() As String
    Dim stateArr As Variant, stateStr As String
    Dim i As Long

    stateStr = ""
    stateArr = DefineStateArr
    For i = LBound(stateArr) To UBound(stateArr)
        stateStr = stateStr & stateArr(i) & "|"
    Next

    BuildStatePattern = Trim(Left(stateStr, Len(stateStr) - 1))
End Function

Function DefineFormDefaults(str As String) As String
    Dim returnStr As String
    
    Select Case str
    
    Case "txtSearchInput"
        returnStr = "Search one or multiple I selectors. Input / paste as many as you want HERE."
        
    Case "txtAddDataDefault"
        returnStr = "Input / Paste I selectors to add HERE. GrayWolfe auto detects data types. <b>ONE RECORD SET PER LINE</b> (see examples below)." & "<br><br>" & _
        "Data can be delimited (separated) by comma, semicolon, pipe, tab, or anything you can imagine." & "<br><br>" & _
        "Input example (pipe-delimited):" & "<br>" & _
        "<i>" & "Kim Jong Un | Residence No. 55 National Hwy 65 Pyongyang, DPRK | linkedin.com/in/superstar1984|@michaeljordanlover84" & "<br>" & _
        "Kim Ju Ae | Dorm 1A Wellington Square Oxford, UK | linkedin.com/in/AI-Dev | @zoomerPrincess2012" & "</i>" & "<br><br>" & _
        "Another example (comma-delimited):" & "<br>" & _
        "<i>" & "entrepreneur1950@gmail.com, 202-123-4567, Bob Smith" & "<br>" & _
        "faciliatorForForeignPower@hotmail.com, 1-701-987-6543, Tim Jones" & "<br>" & _
        "literallyBuyingNukes@yahoo.com, 3025439876, Susie Johnson" & "</i>" & "<br><br>" & _
        "Another example (semicolons):" & "<br>" & _
        "<i>" & "Kim Jong Il; nerd.loser@gmail.com; linkedin.com/in/source-of-the-problem007" & "<br>" & _
        "Robert Robertson; [leave blank]; linkedin.com/in/i-dont-want-more-examples" & "</i>"
        
        
    Case "txtAddDataUnrelated"
        returnStr = "Input / Paste <b>UNRELATED</b> I selectors to add HERE. For bulk lists where connections are unknown or nonexistent." & "<br><br>" & _
        "Input example:" & "<br>" & _
        "<i>" & "Kim Il Sung; CCP_1931@email.org; Robertson Roberts; 123 Fake St. Wilmington, DE 19809; etc." & "</i>" & "<br><br>" & _
        "[TLDR: This import stores each selector separately. If multiple selectors are owned by the same dude please use the OTHER import.]"
        
    Case "cboSQL"
        returnStr = "SELECT 'New Row' AS DisplayValue, 1 as SortOrder FROM MSysObjects " & _
        "UNION SELECT ';', 2 FROM MSysObjects UNION SELECT '|', 3 FROM MSysObjects " & _
        "UNION SELECT '+', 4 FROM MSysObjects UNION SELECT '/', 5 FROM MSysObjects " & _
        "UNION SELECT '\', 6 FROM MSysObjects UNION SELECT ',', 7 FROM MSysObjects " & _
        "UNION SELECT '[space]', 7 FROM MSysObjects ORDER BY SortOrder"

    Case "disclaimerSearch"
        returnStr = "<b>DISCLAIMER:</b> GrayWolfe is NOT an F data repository and does NOT have unique I data. GrayWolfe queries / displays data from the F's official data repository."
        
    Case "disclaimerAdd"
        returnStr = "<b>DISCLAIMER:</b> Only add selectors <b>ALREADY IN S</b>; any selector not <b>ALREADY IN S</b> is auto deleted and removed from GrayWolfe."
        
    Case "resetTargets"
        returnStr = "Are you sure you want to reset the target data?!?" & vbLf & vbLf & _
        "This deletes Target edits since your last save and CANNOT be undone." & vbLf & vbLf & _
        "Click Yes to proceed, No to cancel."
        
    Case "changeTargetId"
        returnStr = "Are you sure you want to change the targetId?!?!1!" & vbLf & vbLf & _
        "Hit Yes to proceed, NO to cancel (pls hit NO)." & vbLf & vbLf & _
        "[TLDR: This is a very bad idea, and will probably break everything, I strongly recommend against it.]" & vbLf & vbLf & _
        "(But I also believe in freedom, so I will let you do it, but please dont unless you really know what you're doing)."
        
    Case "txtToken"
        returnStr = "[Paste YOUR S API Token HERE]"
        

    End Select
    
    DefineFormDefaults = returnStr
End Function


Function DefinePopupText(str As String, selectorType As String, popupType As String) As String
    Dim arr() As String, vowelArr As Variant, textStr As String
    Dim char As String, article As String
    Dim i As Long
    
    Select Case popupType
    
    Case "typeWrong"
        
        'char to check
        char = Trim(Left(selectorType, 1))
    
        'very stupid to be doing in 2025, dont care
        article = "a"
        vowelArr = Array("a", "e", "i", "o", "u")
        For i = LBound(vowelArr) To UBound(vowelArr)
            If char = vowelArr(i) Then article = "an"
        Next
        
        textStr = "Are you sure """ & str & """ is " & article & " " & selectorType & "??" & vbLf & vbLf & _
        "Click YES to override, click NO to cancel and resubmit"

        
    Case "typeNull"
        textStr = "GrayWolfe FAILED to detect a type for """ & str & """" & vbLf & vbLf & _
            "GrayWolfe's *detection algorithm* [a billion if statements and regexs] cant figure it out (#sad_emoticon)." & vbLf & vbLf & _
            "Do you want to skip """ & str & """ and keep going?" & vbLf & vbLf & _
            "Click Yes to skip, click No to set the type with the drop down / try again."
            
    Case "searchDisplayUnrelated"
        'selectorType is number successfully uploaded
        'selectors searched, selectors found
        arr = Split(str, "!!")
        
        If arr(0) = 1 Then
            textStr = "1 unique selector submitted."
        Else
            textStr = arr(0) & " unique selectors submitted."
        End If
        textStr = textStr & vbLf & vbLf
        
        'no new selectors (all old)
        If selectorType = 0 Then
            If arr(1) = 1 Then
                textStr = textStr & "Selector searched is already in GrayWolfe!"
            Else
                textStr = textStr & "All " & arr(1) & " selectors are already in GrayWolfe!"
            End If
            
            textStr = textStr & vbLf & vbLf & "To see full search results, click OK."
            DefinePopupText = textStr: Exit Function
        End If
        
        If selectorType = 1 Then
            If selectorType = arr(0) Then
                textStr = textStr & "Searched selector is NEW and was successfully uploaded to GrayWolfe!"
            Else
                textStr = textStr & "1 selector is NEW and was successfully uploaded to GrayWolfe!"
            End If
                
        Else
            If selectorType = arr(0) Then
                textStr = textStr & "All " & selectorType & " selectors are NEW and were successfully uploaded to GrayWolfe!"
            Else
                textStr = textStr & selectorType & " selectors are NEW and were successfully uploaded to GrayWolfe!"
            End If
                
        End If
        
        textStr = textStr & vbLf & vbLf
                
        'no search hits (all new)
        If arr(1) = 0 Then DefinePopupText = textStr: Exit Function
        
        'hits (mixed new and old)
        If arr(1) = 1 Then
            textStr = textStr & "1 selector is already in GrayWolfe."
        Else
            textStr = textStr & arr(1) & " selectors are already in GrayWolfe."
        End If
        
        textStr = textStr & vbLf & vbLf & "To see the full search results, click OK."
    
    Case "searchDisplayDefault"
         'selectors searched, selectors found, targets found
        arr = Split(str, "!!")
        
        If arr(0) = 1 Then
            textStr = "1 unique selector submitted."
        Else
            textStr = arr(0) & " unique selectors submitted."
        End If
        textStr = textStr & vbLf & vbLf
        
        If arr(1) = 0 Then
            
            If arr(0) = 1 Then
                textStr = textStr & "Searched selctor is NEW and was succesfully uploaded to GrayWolfe!"
            
            Else
                textStr = textStr & "All " & arr(0) & " selectors are NEW and were successfully uploaded to GrayWolfe!"
            End If
            
            DefinePopupText = textStr: Exit Function
        End If
        
        'hits (mixed new and old)
        If arr(1) = 1 Then
            textStr = textStr & "1 selector is already in GrayWolfe."
        Else
            textStr = textStr & arr(1) & " selectors are already in GrayWolfe."
        End If

        textStr = textStr & vbLf & vbLf & "To see the full search results, click OK."
        
    'for rate limit [putting here bc of extra input params]
    Case "rateLimitText"
        textStr = "Searching item number " & Val(str) + 1 & " of " & Val(selectorType) + 1 & " items." & vbLf & vbLf & _
        "We appreciate your confidence running " & Val(selectorType) + 1 & " items through this garbage at once, it probably won't break (inshaAllah)." & vbLf & vbLf & _
        "But I would prefer not to get fired for ddos'ing S, so the tool will take a 10 second break." & vbLf & vbLf & _
        "Please click ok, and then spin around in your chair 5 times while whistling ""Yankee Doodle"" and the tool will resume your search."
        
    Case "searchCheck"
        textStr = "Your search of """ & Trim(selectorType) & """ produced " & Val(str) & " hits!!!" & vbLf & vbLf & _
        "Do you want to wait for this search to finish (it will prob take a while), or do you want the tool to SKIP this item, so everything else finishes faster?" & vbLf & vbLf & _
        "Click YES to continue searching """ & Trim(selectorType) & """, NO to SKIP it."
    
    End Select
               
                
    DefinePopupText = textStr
End Function



'+++++++++++++++++++++++++++++++++

'things have to be in an address
Function DefineStreetArr() As Variant
    Dim streetArr As Variant
    
    streetArr = Array("street", "way", "place", "road", "lane", "drive", "boulevard", "alley", "annex", "avenue", "bypass", "causeway", "center", "circle", _
    "corner", "court", "cove", "creek", "crescent", "crest", "crossing", "dale", "dam", "divide", "estate", "estates", "expressway", "extension", "ferry", _
    "field", "flat", "ford", "forest", "forge", "fork", "fort", "freeway", "garden", "gateway", "green", "glen", "harbor", "haven", "heights", "highway", _
    "hill", "hollow", "inlet", "island", "isle", "junction", "key", "knoll", "lake", "landing", "lodge", "loop", "mall", "manor", "meadows", "mill", "mission", _
    "motorway", "mount", "mountain", "neck", "number", "orchard", "oval", "overpass", "park", "parkway", "pass", "passage", "path", "pike", "plain", "plains", "plaza", "point", _
    "port", "prairie", "ranch", "rapid", "rest", "ridge", "river", "route", "row", "run", "shoal", "shoals", "shore", "shores", "skyway", "spring", "springs", _
    "spur", "square", "squares", "station", "stream", "summit", "terrace", "throughway", "trace", "throughway", "trace", "track", "trail", "tunnel", "turnpike", _
    "underpass", "valley", "viaduct", "view", "village", "villages", "vista", "walk", "wall", "well", _
    "st", "rd", "pl", "ln", "dr", "blvd", "ave", "cir", "ct", "crs", "est", "ext", "ft", "fwy", "gdn", "grn", "hbr", "hvn", "hts", "hwy", "inlt", "is", "jct", "knl", "lk", "lndg", _
    "ldg", "mnr", "mdw", "msn", "mtwy", "mt", "mtn", "no", "orch", "pkwy", "pln", "plz", "pt", "pr", "rnch", "rpd", "ridge", "rt", "rte", "skwy", "sq", "sta", "stn", "ter", _
    "trwy", "trl", "tpke", "vly", "via", "vlg", "vw")
    
    DefineStreetArr = streetArr
End Function

'also includes Canadian provinces
Function DefineStateArr() As Variant
    Dim stateArr As Variant
    
    stateArr = Array("AL", "AK", "AZ", "AR", "AS", "CA", "CO", "CT", "DE", "DC", "D.C.", "FL", "GA", "GU", "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", _
    "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "MP", "OH", "OK", "OR", "PA", "PR", "RI", "SC", "SD", _
    "TN", "TX", "UT", "VT", "VA", "VI", "WA", "WV", "WI", "WY", "AB", "BC", "MB", "NB", "NL", "NT", "NS", "NU", "ON", "PE", "QC", "SK", "YT", _
    "Alabama", "Alaska", "Arizona", "Arkansas", "American Samoa", "California", "Colorado", "Connecticut", "Delaware", "Washington DC", "Washington D.C.", "Florida", "Georgia", "Guam", "Hawaii", _
    "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", _
    "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Northern Marianas", "Ohio", "Oklahoma", "Oregon", _
    "Pennsylvania", "Puerto Rico", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Virgin Islands", "Washington", _
    "West Virginia", "Wisconsin", "Wyoming", "Alberta", "British Columbia", "Manitoba", "New Brunswick", "Newfoundland", "Northwest Territories", "Nova Scotia", "Nunavut", _
    "Ontario", "Prince Edward's Island", "Quebec", "Saskatchewan", "Yukon")
    
    DefineStateArr = stateArr
End Function

'DEFAULT TYPES
Function DefineSelectorTypeArr() As Variant
    Dim selectorTypeArr As Variant
    
    selectorTypeArr = Array("row", "name", "address", "email", "phone", "ip", "linkedin", "github", "discord", "telegram", "other")
    
    DefineSelectorTypeArr = selectorTypeArr
End Function

Function DefineFormTypeArr() As Variant
    Dim formTypeArr As Variant
    
    formTypeArr = Array("name", "address", "email", "phone", "ip", "other")
    
    DefineFormTypeArr = formTypeArr
End Function

Function DefineDetectOrderArr() As Variant
    Dim detectOrderArr As Variant
    
    detectOrderArr = Array("email", "phone", "ip", "address", "linkedin", "github", "telegram", "discord", "name")
    
    DefineDetectOrderArr = detectOrderArr
End Function

Function DefineDetectTypeArr() As Variant
    Dim detectTypeArr As Variant
    
    detectTypeArr = Array("", "CheckName", "CheckAddress", "CheckEmail", "CheckPhone", "CheckIP", "CheckLinkedIn", "CheckGitHub", "CheckDiscord", "CheckTelegram", "")
    
    DefineDetectTypeArr = detectTypeArr

End Function
    'order most obv to least

'using time bc normal things blocked
Function DefineUniqueId() As String
    Dim m As Long
    
    m = CLng((Timer - Int(Timer)) * 1000)
    DefineUniqueId = Format(Now, "YYMMDDHHNNSS") & Format(m, "000")
    
    'DefineUniqueId = Mid(CreateObject("Scriptlet.TypeLib").Guid, 2, 36) 'BLOCKED by #security for no reason
End Function
    
'----------------------

'MAP OBJS
Function DetectFunctionMap(str As String) As String
    Dim dict As Object
    Dim keyArr As Variant, valueArr As Variant, searchStr As String
    Dim i As Long, j As Long
    
   'build dictionary with loop
    keyArr = DefineSelectorTypeArr()
    valueArr = DefineDetectTypeArr()

    Set dict = CreateObject("Scripting.Dictionary")
    For i = LBound(keyArr) To UBound(keyArr)
        dict.Add keyArr(i), valueArr(i)
    Next
    
    'check and return input
    searchStr = Trim(str)
    If CheckKeyExists(dict, searchStr) = False Then DetectFunctionMap = "": Exit Function
        
    DetectFunctionMap = dict(searchStr)
    
    'otherwise throw error
    'ThrowError 1958, str: Exit Function
End Function

Function TargetFormDisplayMap(str As String) As String
    Dim dict As Object
    Dim keyArr As Variant, valueArr As Variant, searchStr As String
    Dim i As Long, j As Long
    
    searchStr = Trim(str)
    
    If searchStr = "" Then TargetFormDisplayMap = "": Exit Function
    
    'keyArr is ("persona name", "street address", "email", "phone", "ip", "linkedin", "github", "discord", "telegram", "other")
    
   'build dictionary with loop
    keyArr = DefineFormTypeArr()
'    valueArr = Array("Form_frmTargetDetails.txtPersonaNames", "Form_frmTargetDetails.txtAddresses", "Form_frmTargetDetails.txtEmails", "Form_frmTargetDetails.txtPhones", _
'    "Form_frmTargetDetails.txtIPs", "Form_frmTargetDetails.txtOther")
    
    valueArr = Array("txtPersonaNames", "txtAddresses", "txtEmails", "txtPhones", "txtIPs", "txtOther")

    Set dict = CreateObject("Scripting.Dictionary")
    For i = LBound(keyArr) To UBound(keyArr)
        dict.Add Trim(keyArr(i)), Trim(valueArr(i))
    Next
    
    'check and return input
    If CheckKeyExists(dict, searchStr) = False Then TargetFormDisplayMap = "": Exit Function
    
    TargetFormDisplayMap = dict(searchStr)
End Function

Function StateMap(str As String) As String
    Dim staObj As Object, atsObj As Object
    Dim stateArr As Variant
    Dim i As Long, x As Long
    
    stateArr = DefineStateArr()
    
    Set atsObj = CreateObject("Scripting.Dictionary")
    Set staObj = CreateObject("Scripting.Dictionary")
    
    
    'ABBREVIATION TO STATE
    x = 70
    For i = 0 To 69
        atsObj.Add stateArr(i), stateArr(x)
        x = x + 1
    Next
    
    'STATE TO ABBREVIATION
    x = 70
    For i = 0 To 69
        staObj.Add stateArr(x), stateArr(i)
        x = x + 1
    Next
    
    'PrintObj atsObj
    'PrintObj staObj
    
    If staObj.Exists(str) Then StateMap = staObj(str): Exit Function
    If atsObj.Exists(str) Then StateMap = atsObj(str): Exit Function
    
    'otherwise fail
    StateMap = "FAIL"
    
    'Debug.Print "STATE ARR: " & UBound(stateArr)

End Function

