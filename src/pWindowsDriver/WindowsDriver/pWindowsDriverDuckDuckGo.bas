VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pWindowsDriverDuckDuckGo"
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
  RefreshPageElement As pWinDriver.pWindowsDriverElement
End Type

Private Sub Class_Initialize()
  This.pPathConfiguration.BrowserRootViewControlType = ""
  This.pPathConfiguration.BrowserRootViewClassName = ""
  This.pPathConfiguration.BrowserRootViewUseWebAppTitleAsName = False
  This.pPathConfiguration.RootWebAreaControlType = "Document"
  This.pPathConfiguration.RootWebAreaAutomationID = ""
  This.pPathConfiguration.HeaderNodeAriaRole = ""
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
    InstanceType = pWinDriver.pInstanceType.ApplicationUserModelID
  End If
  Select Case InstanceType
    Case pWinDriver.pInstanceType.ApplicationUserModelID
      Phosphorus.WindowsProcesses.LaunchAppByAUMID "DuckDuckGo", Phosphorus.WindowsWindowsApps.DuckDuckGo_OfficialName, URL, Phosphorus.WindowShowStates.Normal
    Case Else
      Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverUnhandledAppConfiguration, "Duck Duck Go, Instance Type: #" & InstanceType
  End Select
    
 'Set default path
  Dim CurrentpPath As String
'  CurrentPPath = "/Window[And(@Name=""NoAutomationPeer"",@AutomationId=""BrowserWindow"")]//Pane[@Name=""" & WebAppTitle & """]"
  CurrentpPath = "/Window[And(@Name=""NoAutomationPeer"",@AutomationId=""BrowserWindow"")]"
  This.ParentWindowsDriver.SetPageLoadedElement CurrentpPath, UIAutomationClient.WindowInteractionState.WindowInteractionState_ReadyForUserInteraction
End Sub

Public Sub IWindowsDriverWebBrowser_PostLaunchApp()
  'Nothing to do if empty!
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
  Phosphorus.pExceptions.Raise Phosphorus.Exceptions.MethodNotImplementedYet, "pWindowsDriverDuckDuckGo.CloseAllOtherTabs"
End Sub

Public Sub IWindowsDriverWebBrowser_RefreshPage()
  Phosphorus.pExceptions.Raise Phosphorus.Exceptions.MethodNotImplementedYet, "pWindowsDriverDuckDuckGo.RefreshPage"
End Sub
