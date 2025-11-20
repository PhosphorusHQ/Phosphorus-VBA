Attribute VB_Name = "pWindowsDriverStatic"
'@Folder WindowsDriver
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public gCUIAutomation As CUIAutomation 'Requires a reference to UIAutomationClient
Public gUIADesktopUIElement As UIAutomationClient.IUIAutomationElement
Public DesktopWindowsDriver As pWinDriver.pWindowsDriver
  
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

Public Enum pWebBrowserPPathConfigurationItems
  WebAppTitle
  BrowserRootViewControlType
  BrowserRootViewClassName
  RootWebAreaControlType
  RootWebAreaAutomationID
  HeaderNodeAriaRole
  HeaderNodePPath
  TextNodeAriaRole
  TextNodePPath
  HyperlinkNodeAriaRole
  HyperlinkNodePPath
End Enum

Public Type pWebBrowserPPathConfiguration
  WebAppTitle As String
  BrowserRootViewControlType As String
  BrowserRootViewClassName As String
  RootWebAreaControlType As String
  RootWebAreaAutomationID As String
  HeaderNodeAriaRole As String
  HeaderNodePPath As String
  TextNodeAriaRole As String
  TextNodePPath As String
  HyperlinkNodeAriaRole As String
  HyperlinkNodePPath As String
End Type

Public Function GetNewPDriver(dType As pWinDriver.pWindowsDriverType) As pWinDriver.pWindowsDriver
  Set GetNewPDriver = New pWinDriver.pWindowsDriver
  GetNewPDriver.SetDriverType dType
  If pWinDriver.pWindowsDriverStatic.gCUIAutomation Is Nothing Then
    Set pWinDriver.pWindowsDriverStatic.gCUIAutomation = New CUIAutomation
    If gUIADesktopUIElement Is Nothing Then
      Set gUIADesktopUIElement = pWinDriver.pWindowsDriverStatic.gCUIAutomation.GetRootElement
    End If
  End If
End Function

Public Sub GetDesktopWindowsDriver()
  If DesktopWindowsDriver Is Nothing Then
    Set DesktopWindowsDriver = pWinDriver.pWindowsDriverStatic.GetNewPDriver(pWinDriver.pWindowsDriverType.PreLaunched)
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


