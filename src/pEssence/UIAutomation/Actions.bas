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

Public Function Click(ElementName As String, CurrentElement As IUIAutomationElement)
  Dim CurrentElementInvokePattern As IUIAutomationInvokePattern
  Set CurrentElementInvokePattern = UIACommon.GetPattern(ElementName, CurrentElement, UIA_PatternIds.UIA_InvokePatternId, RaiseError:=True)
  CurrentElementInvokePattern.Invoke
End Function

Public Function SelectListItem(ElementName As String, CurrentElement As IUIAutomationElement)
  Dim CurrentElementSelectionItemPattern As IUIAutomationSelectionItemPattern
  Set CurrentElementSelectionItemPattern = UIACommon.GetPattern(ElementName, CurrentElement, UIA_PatternIds.UIA_SelectionItemPatternId, RaiseError:=True)
Debug.Print "CurrentIsSelected: " & (CurrentElementSelectionItemPattern.CurrentIsSelected = 1)
  CurrentElementSelectionItemPattern.Select
Debug.Print "CurrentIsSelected: " & (CurrentElementSelectionItemPattern.CurrentIsSelected = 1)
  
'  Dim CurrentElementScrollItemPattern As IUIAutomationScrollItemPattern
'  Set CurrentElementScrollItemPattern = UIACommon.GetPattern(ElementName, CurrentElement, UIA_PatternIds.UIA_ScrollItemPatternId, RaiseError:=True)
'  CurrentElementScrollItemPattern.ScrollIntoView
End Function


