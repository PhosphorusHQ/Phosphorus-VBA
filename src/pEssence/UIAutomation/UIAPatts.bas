Attribute VB_Name = "UIAPatts"
'@Folder UIAutomation
Option Explicit

'Tools > References > OLE Automation needed for IUnknown type
Public Function GetPattern(ElementName As String, Element As IUIAutomationElement, PatternId As Long, Optional RaiseError As Boolean) As IUnknown
  On Error Resume Next
  Set GetPattern = Element.GetCurrentPattern(PatternId)
  On Error GoTo 0
  If RaiseError Then
    If GetPattern Is Nothing Then
      pEssence.ErrorLogging.LogError pEssence.Errors.PatternFailedForElement, "Expected to find one element but found " & ElementName
      Exit Function
    End If
  End If
End Function

Public Function HasPattern(ElementName As String, Element As IUIAutomationElement, PatternId As Long, Optional RaiseError As Boolean) As Boolean
  On Error Resume Next
  Dim pattern As IUnknown
  On Error Resume Next
  Set pattern = Element.GetCurrentPattern(PatternId)
  On Error GoTo 0
  HasPattern = Not pattern Is Nothing
End Function
