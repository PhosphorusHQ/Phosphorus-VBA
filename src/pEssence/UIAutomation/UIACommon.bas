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
Public RootDesktopUIAElement As IUIAutomationElement

Public Enum By
  pConditions
  AriaRole
  AutomationId
  ClassName
  ControlType
  NameIs
End Enum

Public Enum UIAPropertyComparisons
  IsTheString
  IsLikeTheString
  EqualsNumber
  StartsWithTheString
  EndsWithTheString
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
      ErrorLogging.LogError Errors.UnhandledGetName, "Unhanded UIAPropertyComparisons: " & "(" & NamedBy & ")"
      Exit Function
  End Select
  GetByName = R
End Function

Public Function GetUIAPropertyComparisonsName(Comparison As UIAPropertyComparisons) As String
  Dim R As String
  Select Case Comparison
    Case UIAPropertyComparisons.IsTheString
      R = "IsTheString"
    Case UIAPropertyComparisons.IsLikeTheString
      R = "IsLike"
    Case UIAPropertyComparisons.EqualsNumber
      R = "Equals"
    Case UIAPropertyComparisons.StartsWithTheString
      R = "StartsWith"
    Case UIAPropertyComparisons.EndsWithTheString
      R = "EndsWith"
    Case Else
      ErrorLogging.LogError Errors.UnhandledGetName, "Unhanded UIAPropertyComparisons: " & "(" & Comparison & ")"
      Exit Function
  End Select
  GetUIAPropertyComparisonsName = R
End Function

Public Function GetElementRuntimeId(ByRef UIAElement As UIAutomationClient.IUIAutomationElement) As String
     
  Dim strRuntimeIDString As String
  strRuntimeIDString = ""
  
  'The runtime ID is a unique array of elements (SAFEARRAY)
  Dim varRuntimeIDArray As Variant
  varRuntimeIDArray = UIAElement.GetRuntimeId
  
  'Check if the RuntimeID was retrieved successfully
  If IsArray(varRuntimeIDArray) Then
    'Iterate through the array
    Dim i As Long
    
    For i = LBound(varRuntimeIDArray) To UBound(varRuntimeIDArray)
      If i > LBound(varRuntimeIDArray) Then
        strRuntimeIDString = strRuntimeIDString & " "
      End If
      strRuntimeIDString = strRuntimeIDString & varRuntimeIDArray(i)
    Next i
       
  Else
  
    MsgBox "Failed to get runtime id!"
       
  End If
  
  GetElementRuntimeId = strRuntimeIDString

End Function

