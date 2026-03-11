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
Option Explicit

Public Sub Click(ElementName As String, CurrentElement As IUIAutomationElement)
  'Try the Invoke pattern
  If UIAPatts.HasPattern(ElementName, CurrentElement, UIA_PatternIds.UIA_InvokePatternId) Then
    Dim CurrentElementInvokePattern As IUIAutomationInvokePattern
    Set CurrentElementInvokePattern = UIAPatts.GetPattern(ElementName, CurrentElement, UIA_PatternIds.UIA_InvokePatternId, RaiseError:=True)
    CurrentElementInvokePattern.Invoke
    Exit Sub
  End If
  
  If UIAProps.GetProperty(CurrentElement, UIA_PropertyIds.UIA_ControlTypePropertyId) = UIA_ControlTypeIds.UIA_ListItemControlTypeId Then
    If UIAPatts.HasPattern(ElementName, CurrentElement, UIA_PatternIds.UIA_SelectionItemPatternId) Then
      Dim CurrentElementSelectionItemPattern As IUIAutomationSelectionItemPattern
      Set CurrentElementSelectionItemPattern = UIAPatts.GetPattern(ElementName, CurrentElement, UIA_PatternIds.UIA_SelectionItemPatternId, RaiseError:=True)
      '?  CurrentElementScrollItemPattern.ScrollIntoView
      CurrentElementSelectionItemPattern.Select
      Exit Sub
    End If
  End If

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
