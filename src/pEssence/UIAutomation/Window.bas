VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder UIAutomation
'VB_PredeclaredId - see: https://www.vbaplanet.com/attributes.php
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================Option Explicit
Option Explicit

Public Sub WaitForInteractionState( _
  ElementName As String, _
  CurrentElement As IUIAutomationElement, _
  UIAPropertyValue As pEssence.UIAWindowInteractionStates, _
  Optional TimeoutInMilliseconds As Long)
  
  Actions.WaitForPropertyValue _
    ElementName, _
    CurrentElement, _
    UIAProperties.WindowWindowInteractionState, _
    UIAPropertyValue
  
End Sub

Public Sub CloseWindow(Name As String, Ele As IUIAutomationElement)
  Actions.IsElementReady Name, Ele
  If UIAProps.GetProperty(Ele, UIAProperties.ControlType) = UIAControlTypeIDs.Window Then
    If UIAProps.HasProperty(Name, Ele, UIAProperties.IsWindowPatternAvailable) Then
      Dim patt As IUIAutomationWindowPattern
      Set patt = UIAPatts.GetPattern(Name, Ele, UIA_PatternIds.UIA_WindowPatternId, RaiseError:=True)
      patt.Close
      Exit Sub
    End If
  End If
End Sub

