Attribute VB_Name = "UIACommon"
'@Folder UIAutomation
Option Explicit

' Requires reference to UIAutomationClient - Tools > References > UIAutomationClient
Public uiAutomation As New CUIAutomation
Private RootDesktopUIAElement As IUIAutomationElement

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

Public Function GetRootDesktopElement() As IUIAutomationElement
  If RootDesktopUIAElement Is Nothing Then
    Set RootDesktopUIAElement = uiAutomation.GetRootElement
  Else
    Set GetRootDesktopElement = RootDesktopUIAElement
  End If
End Function

Public Function GetNewpElement() As pElement
  Dim Element As pElement
  Set Element = New pElement
  Set GetNewpElement = Element
End Function

Public Function GetByName(NamedBy As By) As String
  Dim r As String
  Select Case NamedBy
    Case By.pConditions
      r = "pConditon"
    Case By.AutomationId
      r = "AutomationId"
    Case Else
     r = "Unhandled By Id (" & NamedBy & ")"
  End Select
  GetByName = r
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
