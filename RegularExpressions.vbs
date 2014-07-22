'Regex methods
'SAMPLE USAGE:
  'Dim regLookup
  ''regLookup finds the pattern of a space, a two character state code, and some form of zip code
  'regLookup = "[A-Za-z]{2}\d{4,5}" 
  '
  ''[A-Za-z]{2} = Match a two character code
  ''\d{4,5} = Match 4 to 5 digits
  '
  'If Not(RegExpMatch(objKeyword.Value, regLookup) > 0) Then
  ' routeNum = Left(tranNum,2) & routeNum
  ' tempRN = Left(tranNum,2) & tempRN
  ' objKeyword.Value = tempRN
  'End If


FUNCTION RegExpMatch(StringToSearch, PatternToMatch)
'**************************************************************************************
' Function to return number of RegExp matches in a given string
' 0 - Not Matching

    Dim regEx, CurrentMatches

    Set regEx = New RegExp
    regEx.Pattern = PatternToMatch
    regEx.IgnoreCase = True
    regEx.Global = True
    'regEx.MultiLine = True
    Set CurrentMatches = regEx.Execute(StringToSearch)

    RegExpMatch = CurrentMatches.Count
    Set regEx = Nothing
'**************************************************************************************
END FUNCTION

FUNCTION GetFirstRegExpSubMatch(StringToSearch, PatternToMatch)
'**************************************************************************************
' Function to return First RegExp SubMatch in a given string
' "" - Not Matching

    Dim regEx, CurrentMatches

    Set regEx = New RegExp
    regEx.Pattern = PatternToMatch
    regEx.IgnoreCase = True
    regEx.Global = True
    regEx.MultiLine = True
    Set CurrentMatches = regEx.Execute(StringToSearch)
    If CurrentMatches.Count > 0 Then
        'GetFirstRegExpSubMatch = CurrentMatches.FirstIndex 
        GetFirstRegExpSubMatch = CurrentMatches(0)
    Else
        GetFirstRegExpSubMatch = "" 
    End If
    Set regEx = Nothing
'**************************************************************************************
END FUNCTION

FUNCTION RegExpReplace(StringToSearch, PatternToMatch, StringToReplace)
'**************************************************************************************
' Function to replace a regexp pattern in a string with another string

    Dim regEx, CurrentMatches

    Set regEx = New RegExp
    regEx.Pattern = PatternToMatch
    regEx.IgnoreCase = True
    regEx.Global = True

    RegExpReplace = regEx.Replace(StringToSearch, StringToReplace)
    Set regEx = Nothing
'**************************************************************************************
END FUNCTION
