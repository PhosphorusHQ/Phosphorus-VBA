Attribute VB_Name = "Utils"
'@Folder Utils
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public Function CountOccurrences(Text As String, SearchTerm As String) As Integer
  If Len(SearchTerm) = 0 Then
    CountOccurrences = 0
  Else
    CountOccurrences = (VBA.Strings.Len(Text) - VBA.Strings.Len(VBA.Strings.Replace(Text, SearchTerm, "")))
  End If
End Function

