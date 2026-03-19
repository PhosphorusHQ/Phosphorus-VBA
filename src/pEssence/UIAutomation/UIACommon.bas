Attribute VB_Name = "UIACommon"
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

' Requires reference to UIAutomationClient - Tools > References > UIAutomationClient
Public UIA As New CUIAutomation

Public Enum By
  pConditions
  AutomationId
End Enum

Public Enum UIAPropertyComparisons
  Equals
End Enum

Public Enum TreeScope
  Ancestors = UIAutomationClient.TreeScope.TreeScope_Ancestors
  Children = UIAutomationClient.TreeScope.TreeScope_Children
  Descendants = UIAutomationClient.TreeScope.TreeScope_Descendants
  Element = UIAutomationClient.TreeScope.TreeScope_Element
  None = UIAutomationClient.TreeScope.TreeScope_None
  Parent = UIAutomationClient.TreeScope.TreeScope_Parent
  Subtree = UIAutomationClient.TreeScope.TreeScope_Subtree
End Enum

Public Function GetByName(NamedBy As By) As String
  Dim R As String
  Select Case NamedBy
    Case By.pConditions
      R = "pConditon"
    Case By.AutomationId
      R = "AutomationId"
    Case Else
     R = "Unhandled By Id (" & NamedBy & ")"
  End Select
  GetByName = R
End Function

Public Function IsArrayEmpty(arrInput As Variant) As Boolean
  
  Dim lngTemp As Long
  On Error GoTo HandleError
  lngTemp = UBound(arrInput)
  IsArrayEmpty = False
  Exit Function

HandleError:
  IsArrayEmpty = True

End Function
