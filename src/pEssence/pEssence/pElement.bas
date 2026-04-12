VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder pEssence
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public GivenName As String
Public UIAElement As IUIAutomationElement

Private Sub Class_Terminate()
  Set UIAElement = Nothing
End Sub

Public Function IsAlive() As Boolean
  On Error Resume Next
  Dim pid As Long
  pid = UIAElement.CurrentProcessId  'any property access will fail if stale
  IsAlive = (Err.Number = 0) And (pid > 0)
  On Error GoTo 0
End Function

Public Function GetProperty(PropertyId As Long) As Variant
  On Error Resume Next
  GetProperty = UIAElement.GetCurrentPropertyValue(PropertyId)
  On Error GoTo 0
End Function

'Tools > References > OLE Automation needed for IUnknown type
Public Function GetPattern(PatternId As Long, Optional RaiseError As Boolean) As IUnknown
  On Error Resume Next
  Set GetPattern = UIAElement.GetCurrentPattern(PatternId)
  On Error GoTo 0
  If RaiseError Then
    If GetPattern Is Nothing Then
      ErrorLogging.LogError Errors.PatternFailedForElement, "Pattern failed for element: " & GivenName
      Exit Function
    End If
  End If
End Function

Public Function GetToggleState() As Integer
  If HasPattern(UIAPatterns.TogglePattern) Then
    Dim CurrentElementTogglePattern As IUIAutomationTogglePattern
    Set CurrentElementTogglePattern = GetPattern(UIAPatterns.TogglePattern, RaiseError:=True)
    GetToggleState = CurrentElementTogglePattern.CurrentToggleState
  End If
End Function

Public Function GetValue() As String
  If HasPattern(UIAPatterns.ValuePattern) Then
    Dim CurrentElementValuePattern As IUIAutomationValuePattern
    Set CurrentElementValuePattern = GetPattern(UIAPatterns.ValuePattern, RaiseError:=True)
    GetValue = CurrentElementValuePattern.CurrentValue
  End If
End Function

Public Function HasPattern(PatternId As Long, Optional RaiseError As Boolean) As Boolean
  On Error Resume Next
  Dim Pattern As IUnknown
  On Error Resume Next
  Set Pattern = UIAElement.GetCurrentPattern(PatternId)
  On Error GoTo 0
  HasPattern = Not Pattern Is Nothing
End Function

Public Function IsEnabled() As Boolean
  Window.HighlightElement Me.UIAElement
  IsEnabled = GetProperty(UIAProperties.IsEnabled)
  Window.ReleaseHighlighting
End Function

Public Function IsSelected() As Boolean
            
  Window.HighlightElement Me.UIAElement
            
  ' Method 1: Preferred - Use SelectionItemPattern
  If HasPattern(UIA_PatternIds.UIA_SelectionItemPatternId) Then
     Dim SelectionItemPattern As IUIAutomationSelectionItemPattern
     Set SelectionItemPattern = GetPattern(UIA_PatternIds.UIA_SelectionItemPatternId)
     IsSelected = SelectionItemPattern.CurrentIsSelected
     GoTo Cleanup
  End If
  
  ' Method 2: try TogglePattern
  If HasPattern(UIA_PatternIds.UIA_TogglePatternId) Then
     Dim TogglePattern As IUIAutomationTogglePattern
     Set TogglePattern = GetPattern(UIA_PatternIds.UIA_TogglePatternId)
     IsSelected = TogglePattern.CurrentToggleState
     GoTo Cleanup
  End If
  
  ' Method 2: Fallback - Direct property (works on many controls)
'  Dim varValue As Variant
  'varValue = el.GetProperty(UIA_IsSelectedPropertyId)   ' Property ID 30079
'varValue = Actions.GetValue(Me.GivenName, Me.UIAElement)
'  If IsBoolean(varValue) Then
'    IsRadioButtonSelected = CBool(varValue)
'  Else
'    IsRadioButtonSelected = False
'  End If
    
Cleanup:
  Window.ReleaseHighlighting
    
End Function

'Highlight only when setting the element state - this is an action, but not for gets, which may be part of another action!?
Public Sub SetValue(Value As String)
  Window.HighlightElement UIAElement
  If HasPattern(UIAPatterns.ValuePattern) Then
    Dim CurrentElementValuePattern As IUIAutomationValuePattern
    Set CurrentElementValuePattern = GetPattern(UIAPatterns.ValuePattern, RaiseError:=True)
    CurrentElementValuePattern.SetValue Value
  End If
  Window.ReleaseHighlighting
End Sub
