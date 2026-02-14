*****************************************

MOD 4c_ErrorHandler

*****************************************

Option Compare Database

Option Explicit

Function ThrowError(errCode As Long, errMsg As String) As String
    Dim errorStr As String, displayMsg As String, searchTerm As String
    
    Select Case errCode
    
    'empty search
    Case 1950
        errorStr = "You forgot to input I selectors to search." & vbLf & vbLf & _
        "Please input I selector in the textbox (the pretty BLUE box) and hit the button again."
        
    'empty search target
    Case 1951
        errorStr = "Search target cannot be blank." & vbLf & vbLf & _
        "Please select what you want to search and try again."
    
    'empty selector type
    Case 1952
        errorStr = "Selector type cannot be blank." & vbLf & vbLf & _
        "Please pick a selector type, and try again"

    'empty add
    Case 1953
        errorStr = "You forgot to input I selectors to add." & vbLf & vbLf & _
        "Please input I selector in the textbox (the pretty GREEN box) and hit the button again."
        
    'empty import type
    Case 1954
        errorStr = "Import type cannot be blank." & vbLf & vbLf & _
        "Please select the import method you are using to add I selectors and try again"
        
    'invalid search type
    Case 1955
        errorStr = "The thing you want to search is unsupported." & vbLf & vbLf & _
        "They do not pay me enough (and I am not good enough) to build ways for this garbage to search other things." & vbLf & vbLf & _
        "Please try again and keep your search location to the options provided."
    
    'invalid selector type
    Case 1956
        errorStr = "You entered an unsupported selector type." & vbLf & vbLf & _
        "They do not pay me enough (and I am not good enough) to add MySpace accounts or other rando selectors to this garbage." & vbLf & vbLf & _
        "Please try again and keep your type selections to the options provided."

    'invalid import type
    Case 1957
        errorStr = "You entered an unsupported import type." & vbLf & vbLf & _
        "They do not pay me enough (and I am not good enough) to build other ways to add data to this garbage." & vbLf & vbLf & _
        "Please try again and keep your import selections to the options provided."
    
    'map error, lookup map failed
    Case 1958
        errorStr = "Problem when building map object." & vbLf & vbLf & _
        "This is a coding error. The details are boring and unimportant." & vbLf & vbLf & _
        "Please tell Remedio about this, he will beat the developer and fix it immediately (inshaAllah)."
        
    'character parsing error (CharIsNumber / CharIsLetter)
    Case 1959
        errorStr = "Problem when parsing characters." & vbLf & vbLf & _
        "This is a coding error. The details are boring and unimportant." & vbLf & vbLf & _
        "Please tell Remedio about this, he will beat the developer and fix it immediately (inshaAllah)."
        
    'cant detect type
    Case 1960
        errorStr = "GrayWolfe FAILED to detect a type for """ & errMsg & """" & vbLf & vbLf & _
        "GrayWolfe's *detection algorithm* [a billion if statements and regexs] cant figure it out. I guess you get what you pay for." & vbLf & vbLf & _
        "Please select the *Selector Type* in the dropdown then try again."
    
    'sharepoint search problem
    Case 1961
        errorStr = "Problem searching GrayWolfe" & vbLf & vbLf & _
        "This is a coding error. The details are boring and unimportant." & vbLf & vbLf & _
        "Please tell Remedio about this, he will beat the developer and fix it immediately (inshaAllah)."
        
    'sharepoint data add problem
    Case 1962
        errorStr = "Problem adding data to GrayWolfe" & vbLf & vbLf & _
        "This is a coding error. The details are boring and unimportant." & vbLf & vbLf & _
        "Please tell Remedio about this, he will beat the developer and fix it immediately (inshaAllah)."
        
    'different target names
    Case 1963
        errorStr = "That name is DIFFERENT from what's already in GrayWolfe for those selectors." & vbLf & vbLf & _
        "Tell Remedio to stop sucking and FINISH building a way to deal with this."
        
    'local tables
    Case 1964
        errorStr = "Problem with local tables." & vbLf & vbLf & _
        "This is a coding error. The details are boring and unimportant." & vbLf & vbLf & _
        "Please tell Remedio about this, he will beat the developer and fix it immediately (inshaAllah)."
    
    'delim str
    Case 1965
        errorStr = "GrayWolfe FAILED to detect a delimiter (the thing separating data items) in your data" & vbLf & vbLf & _
        "GrayWolfe's *detection algorithm* [a billion if statements and regexs] cant figure it out. I guess you get what you pay for." & vbLf & vbLf & _
        "Please select the *Delimiter* in the dropdown then try again."
    
    'token empty
    Case 1966
        errorStr = "Please input YOUR S API Token to auto search S." & vbLf & vbLf & _
        "[Click the question mark for a step by step guide on how to do this, or contact ME and I will show you (or do it for you).]" & vbLf & vbLf & _
        "This will take LESS THAN 30 seconds (I promise)."
        
    'token wrong format
    Case 1967
        errorStr = "Your S Token doesn't look right (it's in the wrong format)." & vbLf & vbLf & _
        "Either you pasted it wrong, or (more likely) the tool messed it up." & vbLf & vbLf & _
        "If the latter is more likely please tell Remedio. He will beat the developer and make him fix it immediately (inshaAllah)."
        
    'token fucked (S api not working)
    Case 1968
        errorStr = "TLDR: For the S search API to work you need to LOG OUT of your FNet account and LOG BACK IN." & vbLf & vbLf & _
        "This is a very dumb problem, but in short, your PKI card's security auth has expired because its been too long since you logged in." & vbLf & vbLf & _
        "I am very sorry but there is literally no other way to fix this. Windows knows this 'security' feature breaks everything. They do not care."
        
    'S search fucked
    Case 1969
        errorStr = "Something is wrong with the S search. Prob a coding problem." & vbLf & vbLf & _
        "Please tell Remedio, he will beat the developer and fix it immediately (inshaAllah)."
    
    'S search empty
    Case 1970
        errorStr = "S's API failed to return data, no clue what the problem is" & vbLf & vbLf & _
        "Please tell Remedio about this, he will beat the developer, and fix it immediately (inshaAllah)"
    
    'open args wrong
    Case 1971
        errorStr = "There was a problem loading data into the results form. This is a coding error (apologies)." & vbLf & vbLf & _
        "Please tell Remedio about this, he will beat the developer, and fix it immediately (inshaAllah)"

    'user cancel
    Case 1998
        GoTo Kill

    End Select
    
    Debug.Print "Error Code: " & errCode & "; Message: " & errMsg & "; ErrorStr (popup): " & errorStr
    
    'Display message
    displayMsg = MsgBox(errorStr, , "#FAIL")
    GoTo Kill
    
    Exit Function

'End All Execution (prob causes crashes)
Kill:
    End
    
End Function
