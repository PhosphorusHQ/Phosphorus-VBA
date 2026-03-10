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

'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-automation-element-propids
Public Enum UIAProperties
  AcceleratorKey = 30006
  AccessKey = 30007
  AnnotationObjects = 30156
  AnnotationTypes = 30155
'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-ariaspecification
  AriaProperties = 30102
  AriaRole = 30101
  AutomationId = 30011
  BoundingRectangle = 30001
  CenterPoint = 30165
  ClassName = 30012
  ClickablePoint = 30014
  ControllerFor = 30104
  ControlType = 30003
  Culture = 30015
  DescribedBy = 30105
  FillColor = 30160
  FillType = 30162
  FlowsFrom = 30148
  FlowsTo = 30106
  FrameworkId = 30024
  FullDescription = 30159
  HasKeyboardFocus = 30008
  HeadingLevel = 30173
  HelpText = 30013
  IsContentElement = 30017
  IsControlElement = 30016
  IsDataValidForForm = 30103
  IsDialog = 30174
  IsEnabled = 30010
  IsKeyboardFocusable = 30009
  IsOffscreen = 30022
  IsPassword = 30019
  IsPeripheral = 30150
  IsRequiredForForm = 30025
  ItemStatus = 30026
  ItemType = 30021
  LabeledBy = 30018
  LandmarkType = 30157
  Level = 30154
  LiveSetting = 30135
  LocalizedControlType = 30004
  LocalizedLandmarkType = 30158
  Name = 30005
  NativeWindowHandle = 30020
  OptimizeForVisualContent = 30111
  Orientation = 30023
  OutlineColor = 30161
  OutlineThickness = 30164
  PositionInSet = 30152
  ProcessId = 30002
  ProviderDescription = 30107
  Rotation = 30166
  RuntimeId = 30000
  Size = 30167
  SizeOfSet = 30153
  VisualEffects = 30163
End Enum

Public Enum UIAPropertyComparisons
  Equals
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

Public Function GetWindowInteractionStateName(WindowInteractionState As Long) As String
  Dim r As String
  Select Case WindowInteractionState
    Case UIAutomationClient.WindowInteractionState.WindowInteractionState_Running: r = "Running"
    Case UIAutomationClient.WindowInteractionState.WindowInteractionState_Closing: r = "Closing"
    Case UIAutomationClient.WindowInteractionState.WindowInteractionState_ReadyForUserInteraction: r = "ReadyForUserInteraction"
    Case UIAutomationClient.WindowInteractionState.WindowInteractionState_BlockedByModalWindow: r = "BlockedByModalWindow"
    Case UIAutomationClient.WindowInteractionState.WindowInteractionState_NotResponding: r = "NotResponding"
    Case Else: r = "Unknown Window Interaction State (" & WindowInteractionState & ")"
  End Select
  GetWindowInteractionStateName = r
End Function

Public Function GetProperty(Element As IUIAutomationElement, PropertyId As Long) As Variant
  On Error Resume Next
  GetProperty = Element.GetCurrentPropertyValue(PropertyId)
  On Error GoTo 0
End Function

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

Public Function IsArrayEmpty(arrInput As Variant) As Boolean
  
  Dim lngTemp As Long
  On Error GoTo HandleError
  lngTemp = UBound(arrInput)
  IsArrayEmpty = False
  Exit Function

HandleError:
  IsArrayEmpty = True

End Function
