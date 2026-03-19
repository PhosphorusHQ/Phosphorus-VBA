VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Actions"
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

Public Sub Click(Name As String, Ele As IUIAutomationElement)
  
  IsElementReady Name, Ele
  Window.HighlightElement Ele
    
  'Try the Invoke pattern
  If UIAProps.HasProperty(Name, Ele, UIAProperties.IsInvokePatternAvailable) Then
    If UIAProps.GetProperty(Ele, UIAProperties.IsInvokePatternAvailable) Then
      Dim CurrentElementInvokePattern As IUIAutomationInvokePattern
      Set CurrentElementInvokePattern = UIAPatts.GetPattern(Name, Ele, UIA_PatternIds.UIA_InvokePatternId, RaiseError:=True)
      CurrentElementInvokePattern.Invoke
      GoTo Cleanup
    End If
  End If
    
  If UIAProps.HasProperty(Name, Ele, UIAProperties.ControlType) Then
    If UIAProps.GetProperty(Ele, UIAProperties.ControlType) = UIAControlTypeIDs.ListItem Then
      If UIAProps.HasProperty(Name, Ele, UIAProperties.IsSelectionItemPatternAvailable) Then
        Dim CurrentElementSelectionItemPattern As IUIAutomationSelectionItemPattern
        Set CurrentElementSelectionItemPattern = UIAPatts.GetPattern(Name, Ele, UIA_PatternIds.UIA_SelectionItemPatternId, RaiseError:=True)
        CurrentElementSelectionItemPattern.Select
      GoTo Cleanup
      End If
    End If
  End If
  
  Exit Sub

Cleanup:
  Window.ReleaseHighlighting

End Sub

Public Sub WaitForPropertyValue( _
  ElementName As String, _
  CurrentElement As IUIAutomationElement, _
  UIAProperty As UIAProperties, _
  UIAPropertyValue As Variant, _
  Optional TimeoutInMilliseconds As Long)
  
  Dim CurrentPropertyValue As Variant
  CurrentPropertyValue = pEssence.GetProperty(CurrentElement, UIAProperty)
  If CurrentPropertyValue = UIAPropertyValue Then
    'Success - exit here!
    Exit Sub
  Else
    MsgBox "Need to wait for the value before timeout here!"
  End If
  
End Sub

Public Function IsElementReady(Name As String, Ele As IUIAutomationElement) As Boolean
  Dim ret As Boolean
  ret = True
  ret = ret And IsElementAlive(Name, Ele)
  If ret Then
    TryToScrollItemIntoView Name, Ele
  Else
    MsgBox "Raise an error?"
  End If
  IsElementReady = ret
End Function

Private Function IsElementAlive(Name As String, Ele As IUIAutomationElement) As Boolean
  On Error Resume Next
  Dim pid As Long
  pid = Ele.CurrentProcessId  'any property access will fail if stale
  IsElementAlive = (Err.Number = 0) And (pid > 0)
  On Error GoTo 0
End Function

Public Sub TryToScrollItemIntoView(Name As String, Ele As IUIAutomationElement)
  If UIAProps.HasProperty(Name, Ele, UIAProperties.IsScrollItemPatternAvailable) Then
    If UIAProps.HasProperty(Name, Ele, UIAProperties.IsOffscreen) Then
      If UIAProps.GetProperty(Ele, UIAProperties.IsOffscreen) Then
        Dim patt As IUIAutomationScrollItemPattern
        Set patt = UIAPatts.GetPattern(Name, Ele, UIA_PatternIds.UIA_ScrollItemPatternId, RaiseError:=True)
        patt.ScrollIntoView
      End If
    End If
  End If
End Sub

