Attribute VB_Name = "PPathExceUserDefinedFunctions"
'@Folder PPath
Option Explicit

'https://excelmacromastery.com/vba-dictionary/

Private FunctionNameMappings As Scripting.dictionary

Public Sub ClearDownFunctionNameMappings()
  Set FunctionNameMappings = Nothing
End Sub

Public Sub InitialiseFunctionNameMappings()
  'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controltype-ids
  If FunctionNameMappings Is Nothing Then
    Set FunctionNameMappings = New Scripting.dictionary
    FunctionNameMappings.Add "ceiling_math", "ceiling.math"
    FunctionNameMappings.Add "floor_math", "floor.math"
    FunctionNameMappings.Add "xp:contains", "xp_contains"
    FunctionNameMappings.Add "xp:format-number", "xp_format_number"
    FunctionNameMappings.Add "xp:starts-with", "xp_starts_with"
    FunctionNameMappings.Add "xp:substring", "xp_substring"
    FunctionNameMappings.Add "xp:normalize-space", "xp_normalize_space"
    FunctionNameMappings.Add "xp:string-length", "xp_string_length"
    FunctionNameMappings.Add "xp:string-after", "xp_string_after"
    FunctionNameMappings.Add "xp:string-before", "xp_string_before"
    FunctionNameMappings.Add "xp:translate", "xp_translate"
    FunctionNameMappings.Add "udf:text-between", "udf_text_between"
  End If
End Sub

Public Sub AddFunctionNameMapping(strFrom As String, strTo As String)
  FunctionNameMappings.Add strFrom, strTo
End Sub

Public Function RenameExcelFunctions(strCurrentPredicateTest) As String
  InitialiseFunctionNameMappings
  Dim strReturn As String
  strReturn = strCurrentPredicateTest
  Dim key As Variant
  For Each key In FunctionNameMappings.Keys
     strReturn = VBA.Strings.Replace(strReturn, key & "(", FunctionNameMappings(key) & "(")
  Next key
  'Handle all user defined function's in 1 - no need to map them individually
  strReturn = VBA.Strings.Replace(strReturn, "udf:", "udf_")
  RenameExcelFunctions = strReturn
End Function

Public Function udf_reversetext(strText) As String
  Dim strReturn As String
  If strText = "" Then
    strReturn = ""
  Else
    Dim i As Integer
    Dim c As String
    For i = 1 To VBA.Strings.Len(strText)
      c = VBA.Strings.Mid(strText, i, 1)
      strReturn = c & strReturn
    Next i
  End If
  udf_reversetext = strReturn
End Function

Public Function xp_contains(str1 As String, str2 As String) As Boolean
  xp_contains = (VBA.Strings.InStr(1, str1, str2) > 0)
End Function

Public Function xp_starts_with(str1 As String, str2 As String) As Boolean
  xp_starts_with = (VBA.Strings.InStr(1, str1, str2) = 1)
End Function

Public Function xp_format_number(var1 As Variant, strFormat As String) As String
  xp_format_number = Application.WorksheetFunction.Text(var1, strFormat)
End Function

Public Function xp_normalize_space(str As String) As String
  Dim strReturn As String
  strReturn = VBA.Strings.Replace(str, vbCrLf, "")
  strReturn = VBA.Strings.Replace(strReturn, vbCr, "")
  strReturn = VBA.Strings.Replace(strReturn, vbLf, "")
  strReturn = VBA.Strings.Replace(strReturn, vbTab, "")
  strReturn = VBA.Strings.Trim(strReturn)
  strReturn = VBA.Strings.Replace(strReturn, vbTab, "")
  While VBA.Strings.InStr(1, strReturn, "  ") > 0
    strReturn = VBA.Strings.Replace(strReturn, "  ", " ")
  Wend
  xp_normalize_space = strReturn
End Function

Public Function xp_string_length(str As String) As Integer
  xp_string_length = VBA.Len(str)
End Function

Public Function xp_substring(str As String, intFrom As Integer, Optional intTo As Integer) As String
  If intTo = 0 Then
    xp_substring = VBA.Strings.Mid(str, intFrom)
  Else
    xp_substring = VBA.Strings.Mid(str, intFrom, intTo - intFrom + 1)
  End If
End Function

Public Function xp_string_after(str As String, strSearch As String) As String
  xp_string_after = str
  If strSearch <> "" Then
    Dim iStartOfSearchString As Integer
    iStartOfSearchString = VBA.Strings.InStr(1, str, strSearch)
    If iStartOfSearchString > 0 Then
      xp_string_after = VBA.Strings.Mid(str, iStartOfSearchString + VBA.Len(strSearch), VBA.Len(str))
    End If
  End If
End Function

Public Function xp_string_before(str As String, strSearch As String) As String
  xp_string_before = str
  If strSearch <> "" Then
    Dim iStartOfSearchString As Integer
    iStartOfSearchString = VBA.Strings.InStr(1, str, strSearch)
    If iStartOfSearchString > 0 Then
      xp_string_before = VBA.Strings.Mid(str, 1, iStartOfSearchString - 1)
    End If
  End If
End Function

Public Function xp_translate(strSource As String, strMap As String, strTranslate As String) As String
  Dim strReturn As String
  strReturn = ""
  Dim strChar As String
  Dim countStr, countMap, countTranslate As Integer
  countStr = VBA.Strings.Len(strSource)
  countMap = VBA.Strings.Len(strMap)
  countTranslate = VBA.Strings.Len(strTranslate)
  Dim i As Integer
  For i = 1 To countStr
    strChar = VBA.Strings.Mid(strSource, i, 1)
    Dim j As Integer
    Dim strTranslatedCharacter As String
    strTranslatedCharacter = ""
    Dim boolCharacterFound As Boolean
    boolCharacterFound = False
    For j = 1 To countMap
      If strChar = VBA.Strings.Mid(strMap, j, 1) Then
        boolCharacterFound = True
        If j > countTranslate Then
          strTranslatedCharacter = ""
        Else
          strReturn = strReturn & VBA.Strings.Mid(strTranslate, j, 1)
        End If
        Exit For
      End If
    Next j
    If Not boolCharacterFound Then
      strReturn = strReturn & strChar
    End If
  Next i
  xp_translate = strReturn
End Function

'https://wellsr.com/vba/2016/excel/easily-extract-text-between-two-strings-with-vba/?s=09
Public Function udf_text_between(ByVal strMain As String, str1 As String, str2 As String, Optional reverse As Boolean) As String
  'DESCRIPTION: Extract the portion of a string between the two substrings defined in str1 and str2.
  'DEVELOPER: Ryan Wells (wellsr.com)
  'HOW TO USE: - Pass the argument your main string and the 2 strings you want to find in the main string.
  ' - This function will extract the values between the end of your first string and the beginning of your next string.
  ' - If the optional boolean "reverse" is true, an InStrRev search will occur to find the last instance of the substrings in your main string.
  Dim i As Integer, j As Integer, temp As Variant
'  On Error GoTo errhandler:
  If reverse = True Then
    i = VBA.Strings.InStrRev(strMain, str1)
    j = VBA.Strings.InStrRev(strMain, str2)
    If VBA.Math.Abs(j - i) < VBA.Strings.Len(str1) Then j = VBA.Strings.InStrRev(strMain, str2, i)
      If i = j Then 'try to search 2nd half of string for unique match
          j = VBA.Strings.InStrRev(strMain, str2, i - 1)
      End If
    Else
      i = VBA.Strings.InStr(1, strMain, str1)
      j = VBA.Strings.InStr(1, strMain, str2)
      If VBA.Math.Abs(j - i) < VBA.Strings.Len(str1) Then j = VBA.Strings.InStr(i + VBA.Strings.Len(str1), strMain, str2)
      If i = j Then 'try to search 2nd half of string for unique match
          j = VBA.Strings.InStr(i + 1, strMain, str2)
      End If
    End If
    If i = 0 And j = 0 Then
      udf_text_between = strMain
      Exit Function
    End If
    If j = 0 Then j = VBA.Strings.Len(strMain) + VBA.Strings.Len(str2) 'just to make it arbitrarily large
    If i = 0 Then i = VBA.Strings.Len(strMain) + VBA.Strings.Len(str1) 'just to make it arbitrarily large
    If i > j And j <> 0 Then 'swap order
      temp = j
      j = i
      i = temp
      temp = str2
      str2 = str1
      str1 = temp
    End If
    i = i + VBA.Strings.Len(str1)
    udf_text_between = VBA.Strings.Mid(strMain, i, j - i)
    Exit Function
'errhandler:
'  MsgBox "Error extracting strings. Check your input" & vbNewLine & vbNewLine & "Aborting", , "Strings not found"
  End
End Function

Public Function udf_test_starts_with(str1 As String, str2 As String) As Boolean
  udf_test_starts_with = (VBA.Strings.InStr(1, str1, str2) = 1)
End Function

Public Function udf_test_ends_with(str1 As String, str2 As String) As Boolean
  udf_test_ends_with = (VBA.Strings.InStr(1, str1, str2) = (VBA.Strings.Len(str1) - VBA.Strings.Len(str2) + 1))
End Function

