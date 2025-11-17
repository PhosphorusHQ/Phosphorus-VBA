Attribute VB_Name = "Utils"
'@Folder pPath
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public Function OutputElementDetails( _
  eleUIElement As UIAutomationClient.IUIAutomationElement, _
  strAbsoluteXPath, _
  ByRef DebugMode As Boolean, _
  AutomationDictionaries As UIAutomationDictionaries)
  
'  If DebugMode Then
'    Debug.Print strAbsoluteXPath
'    Debug.Print _
'      "ControlTypeID: " & eleUIElement.CurrentControlType, _
'      "ControlTypeName: " & AutomationDictionaries.ControlTypeIDs(eleUIElement.CurrentControlType), _
'      "Name: " & eleUIElement.CurrentName
''    Debug.Print ""
'  End If
End Function

Public Function GetValue(UIElement As UIAutomationClient.IUIAutomationElement) As Variant
  Dim varValue As Variant
  varValue = UIElement.GetCurrentPropertyValue(UIA_ValueValuePropertyId)
  If varValue = "" Then
    Dim ValuePattern As IUIAutomationValuePattern
    On Error Resume Next
    varValue = UIElement.GetCurrentPattern(UIA_ValuePatternId)
    On Error GoTo 0
  End If
  GetValue = varValue
End Function


