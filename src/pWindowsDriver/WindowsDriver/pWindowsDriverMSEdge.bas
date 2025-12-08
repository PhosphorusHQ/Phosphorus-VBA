VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pWindowsDriverMSEdge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Implements pWinDriver.IWindowsDriverWebBrowser

Dim This As Configuration
Private Type Configuration
  TempDirectory As String
  ParentWindowsDriver As pWinDriver.pWindowsDriver
  pPathConfiguration As pWinDriver.pWebBrowserPPathConfiguration
End Type

Private Sub Class_Initialize()
  This.pPathConfiguration.BrowserRootViewControlType = "Pane"
  This.pPathConfiguration.BrowserRootViewClassName = "BrowserRootView"
  This.pPathConfiguration.BrowserRootViewUseWebAppTitleAsName = False
  This.pPathConfiguration.RootWebAreaControlType = "Document"
  This.pPathConfiguration.RootWebAreaAutomationID = "RootWebArea"
  This.pPathConfiguration.HeaderNodeAriaRole = "heading"
  This.pPathConfiguration.TextNodeAriaRole = "description"
  This.pPathConfiguration.HyperlinkNodeAriaRole = "link"
End Sub

Public Function IWindowsDriverWebBrowser_GetParentWindowsDriver() As pWindowsDriver
  Set IWindowsDriverWebBrowser_GetParentWindowsDriver = This.ParentWindowsDriver
End Function
  
Public Sub IWindowsDriverWebBrowser_LaunchApp(ByRef ParentWindowsDriver As pWindowsDriver, WebAppName As String, WebAppTitle As String, Optional URL As String)
  
  This.pPathConfiguration.WebAppTitle = WebAppTitle
  
  Set This.ParentWindowsDriver = ParentWindowsDriver
  Dim InstanceType As pWinDriver.pInstanceType
  InstanceType = This.ParentWindowsDriver.GetWindowsDriverWebBrowserType.InstanceType
  If InstanceType = 0 Then
'TODO: Make new window the default?
    InstanceType = pWinDriver.pInstanceType.ReuseACurrentOpenInstance
  End If
  
  'Set the default PPath
  Dim CurrentPPathOfPageLoadedPPath As String
  Dim PageLoadedElementExpectedWindowInteractionState As UIAutomationClient.WindowInteractionState
  CurrentPPathOfPageLoadedPPath = "/Window[xp:starts-with(@Name,""" & WebAppTitle & """)]"
  PageLoadedElementExpectedWindowInteractionState = UIAutomationClient.WindowInteractionState.WindowInteractionState_ReadyForUserInteraction
  
  Select Case InstanceType
    
    Case pWinDriver.pInstanceType.ReuseACurrentOpenInstance
      'Launch Edge via protocol
      Phosphorus.WindowsProcesses.LaunchCommandByProtocol WebAppName, "microsoft-edge:", URL, Phosphorus.WindowShowStates.SW_SHOWMAXIMIZED

    Case pWinDriver.pInstanceType.Executable
      'Launch Edge via executable with no parameters other than the url, if any
      Phosphorus.WindowsProcesses.LaunchExecutable Phosphorus.WindowsExecutables.MicrosoftEdge, URL, Phosphorus.WindowShowStates.SW_SHOWMAXIMIZED
    
    Case pWinDriver.pInstanceType.NewWindow
      'Launch Edge via executable with new window command line argument & the url, if any
      Phosphorus.WindowsProcesses.LaunchExecutable Phosphorus.WindowsExecutables.MicrosoftEdge, "--new-window " & URL, Phosphorus.WindowShowStates.SW_SHOWMINIMIZED
    
    Case pWinDriver.pInstanceType.AppMode
      'Launch Edge via executable in App Mode (new window + simplified interface)
      Phosphorus.WindowsProcesses.LaunchExecutable Phosphorus.WindowsExecutables.MicrosoftEdge, "--app " & URL, Phosphorus.WindowShowStates.SW_SHOWMINIMIZED
    
    Case pWinDriver.pInstanceType.NewProfile
      This.TempDirectory = This.ParentWindowsDriver.CreateTempDirectory
      Phosphorus.WindowsProcesses.LaunchExecutable Phosphorus.WindowsExecutables.MicrosoftEdge, "--user-data-dir=""" & This.TempDirectory & """ --new-window " & URL, Phosphorus.WindowShowStates.SW_SHOWMAXIMIZED
      CurrentPPathOfPageLoadedPPath = "/Window[And(xp:starts-with(@Name,""Sign Up – Create a Free Account""),@ClassName=""Chrome_WidgetWin_1"")]"
      PageLoadedElementExpectedWindowInteractionState = UIAutomationClient.WindowInteractionState.WindowInteractionState_BlockedByModalWindow
    
    Case pWinDriver.pInstanceType.ApplicationUserModelID
      Phosphorus.WindowsProcesses.LaunchAppByAUMID Phosphorus.WindowsWindowsApps.MicrosoftEdge, URL, Phosphorus.WindowShowStates.SW_SHOWNORMAL
    
    Case Else
      Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverUnhandledAppConfiguration, "Microsoft Edge, Instance Type: #" & InstanceType
  
  End Select
  
  This.ParentWindowsDriver.SetPageLoadedElement CurrentPPathOfPageLoadedPPath, PageLoadedElementExpectedWindowInteractionState

End Sub

Public Function IWindowsDriverWebBrowser_GetPPathConfigurationItem(ItemType As pWinDriver.pWebBrowserPPathConfigurationItems) As Variant
  Select Case ItemType
    Case pWebBrowserPPathConfigurationItems.WebAppTitle
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.pPathConfiguration.WebAppTitle
    Case pWebBrowserPPathConfigurationItems.BrowserRootViewControlType
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.pPathConfiguration.BrowserRootViewControlType
    Case pWebBrowserPPathConfigurationItems.BrowserRootViewClassName
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.pPathConfiguration.BrowserRootViewClassName
    Case pWebBrowserPPathConfigurationItems.BrowserRootViewUseWebAppTitleAsName
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.pPathConfiguration.BrowserRootViewUseWebAppTitleAsName
    Case pWebBrowserPPathConfigurationItems.RootWebAreaControlType
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.pPathConfiguration.RootWebAreaControlType
    Case pWebBrowserPPathConfigurationItems.RootWebAreaAutomationID
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.pPathConfiguration.RootWebAreaAutomationID
    Case pWebBrowserPPathConfigurationItems.HeaderNodeAriaRole
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.pPathConfiguration.HeaderNodeAriaRole
    Case pWebBrowserPPathConfigurationItems.TextNodeAriaRole
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.pPathConfiguration.TextNodeAriaRole
    Case pWebBrowserPPathConfigurationItems.HyperlinkNodeAriaRole
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.pPathConfiguration.HyperlinkNodeAriaRole
  End Select
End Function

Public Sub IWindowsDriverWebBrowser_CloseAllOtherTabs()
  Phosphorus.pExceptions.Raise Phosphorus.Exceptions.MethodNotImplementedYet, "pWindowsDriverMSEdge.CloseAllOtherTabs"
End Sub

Public Sub IWindowsDriverWebBrowser_RefreshPage()
  Phosphorus.pExceptions.Raise Phosphorus.Exceptions.MethodNotImplementedYet, "pWindowsDriverMSEdge.RefreshPage"
End Sub

'        'TODO: How to open new window for edge?



