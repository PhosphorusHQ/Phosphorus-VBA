Attribute VB_Name = "UIAPatts"
'@Folder UIAutomation
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit


Public Enum UIAPatterns
  SelectionItemPattern = UIA_PatternIds.UIA_SelectionItemPatternId
  TogglePattern = UIA_PatternIds.UIA_TogglePatternId
End Enum

'Tools > References > OLE Automation needed for IUnknown type
Public Function GetPattern(Element As pElement, PatternId As Long, Optional RaiseError As Boolean) As IUnknown
  On Error Resume Next
  Set GetPattern = Element.UIAElement.GetCurrentPattern(PatternId)
  On Error GoTo 0
  If RaiseError Then
    If GetPattern Is Nothing Then
      ErrorLogging.LogError Errors.PatternFailedForElement, "Pattern failed for element: " & Element.GivenName
      Exit Function
    End If
  End If
End Function

Public Function HasPattern(Element As pElement, PatternId As Long, Optional RaiseError As Boolean) As Boolean
  On Error Resume Next
  Dim Pattern As IUnknown
  On Error Resume Next
  Set Pattern = Element.UIAElement.GetCurrentPattern(PatternId)
  On Error GoTo 0
  HasPattern = Not Pattern Is Nothing
End Function
