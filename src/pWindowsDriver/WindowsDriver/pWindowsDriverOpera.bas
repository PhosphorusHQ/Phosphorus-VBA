VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pWindowsDriverOpera"
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
  'Opera doesn't have a BrowserRootView so we force it to be the same as the BrowserRootWebArea element
  This.pPathConfiguration.BrowserRootViewControlType = "Document" '
  This.pPathConfiguration.BrowserRootViewClassName = ""
  This.pPathConfiguration.BrowserRootViewUseWebAppTitleAsName = True
  This.pPathConfiguration.RootWebAreaControlType = ""
  This.pPathConfiguration.RootWebAreaAutomationID = ""
  This.pPathConfiguration.HeaderNodeAriaRole = ""
  This.pPathConfiguration.TextNodeAriaRole = ""
  This.pPathConfiguration.HyperlinkNodeAriaRole = ""
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
    InstanceType = pWinDriver.pInstanceType.Executable
  End If
  
  'Set default path
  Dim CurrentpPath As String
  CurrentpPath = "/Pane[@Name=""" & WebAppTitle & " - Opera""]"
  Select Case InstanceType
    Case pWinDriver.pInstanceType.Executable
      'Launch Opera via executable with no parameters other than the url, if any
      Phosphorus.WindowsProcesses.LaunchExecutable Phosphorus.WindowsExecutables.Opera, URL, Phosphorus.WindowShowStates.SW_SHOWMAXIMIZED
    Case Else
      Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverUnhandledAppConfiguration, "Opera, Instance Type: #" & InstanceType
  End Select
  This.ParentWindowsDriver.SetPageLoadedElement CurrentpPath, UIAutomationClient.WindowInteractionState.WindowInteractionState_ReadyForUserInteraction
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
  Dim ele As pWinDriver.pWindowsDriverElement
  Set ele = This.ParentWindowsDriver.FindElement("Last Tab", "/Pane//Tab[@Name=""Tab bar""]//TabItem[last()]")
  ele.Actions.RightClick
  Set ele = This.ParentWindowsDriver.FindElement("Close Tab Context Menu Item", "/Pane[1]//MenuBar//MenuItem[@Name=""Close other tabs""]", TimeoutInSeconds:=5)
  ele.Actions.Click
End Sub

Public Sub IWindowsDriverWebBrowser_RefreshPage()
  Dim ele As pWinDriver.pWindowsDriverElement
  Set ele = This.ParentWindowsDriver.FindElement("Last Tab", "/Pane//ToolBar[@Name=""Navigation""]//Button[@Name=""Reload""]")
  ele.Actions.Click
End Sub

