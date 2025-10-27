Attribute VB_Name = "pWindowsDriverStatic"
'@Folder WindowsDriver
Option Explicit

Public gCUIAutomation As CUIAutomation
Public gUIADesktopUIElement As UIAutomationClient.IUIAutomationElement
Public DesktopWindowsDriver As Phosphorus.pWindowsDriver
  
Public Enum pWindowsDriverType
  PreLaunched = 0
  WebBrowser = 1
  WindowsApp = 2
'  Windows = 1
End Enum

Public Enum pWebBrowserType
  MicrosoftEdge = 1
  Chrome = 2
  Firefox = 3
  Opera = 4
  DuckDuckGo = 5
  Brave = 6
End Enum

Public Enum pInstanceType
  ReuseACurrentOpenInstance = 1
  Executable = 2
  NewWindow = 3
  AppMode = 4
  NewProfile = 5
  GuestModeNoSignIn = 6
  ApplicationUserModelID = 7
End Enum

Public Sub GetDesktopWindowsDriver()
  If DesktopWindowsDriver Is Nothing Then
    Set DesktopWindowsDriver = Phosphorus.Factory.GetNewPDriver(Phosphorus.pWindowsDriverType.PreLaunched)
  End If
End Sub

Public Function GetWindowInteractionStateDescription(State As UIAutomationClient.WindowInteractionState) As String

  Dim DescriptionOfDesiredState As String
  Select Case State
    Case UIAutomationClient.WindowInteractionState.WindowInteractionState_BlockedByModalWindow
      DescriptionOfDesiredState = "Blocked By ModalWindow"
    Case UIAutomationClient.WindowInteractionState.WindowInteractionState_Closing
      DescriptionOfDesiredState = "Closing"
    Case UIAutomationClient.WindowInteractionState.WindowInteractionState_NotResponding
      DescriptionOfDesiredState = "Not Responding"
    Case UIAutomationClient.WindowInteractionState.WindowInteractionState_ReadyForUserInteraction
      DescriptionOfDesiredState = "Ready For User Interaction"
    Case UIAutomationClient.WindowInteractionState.WindowInteractionState_Running
      DescriptionOfDesiredState = "Running"
    Case Else
     'Raise Error
     Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverUndefinedWindowInteractionState, State
  End Select
  GetWindowInteractionStateDescription = DescriptionOfDesiredState
  
End Function

