Attribute VB_Name = "Utils"
'@Folder Utils
Option Explicit

Public Function CountOccurrences(Text As String, SearchTerm As String) As Integer
  If Len(SearchTerm) = 0 Then
    CountOccurrences = 0
  Else
    CountOccurrences = (VBA.Strings.Len(Text) - VBA.Strings.Len(VBA.Strings.Replace(Text, SearchTerm, "")))
  End If
End Function

