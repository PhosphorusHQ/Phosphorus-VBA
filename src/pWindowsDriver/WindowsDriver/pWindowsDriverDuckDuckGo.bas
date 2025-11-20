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
  PPathConfiguration As pWinDriver.pWebBrowserPPathConfiguration
End Type

Private Sub Class_Initialize()
  This.PPathConfiguration.BrowserRootViewControlType = ""
  This.PPathConfiguration.BrowserRootViewClassName = ""
  This.PPathConfiguration.RootWebAreaControlType = "Document"
  This.PPathConfiguration.RootWebAreaAutomationID = ""
  This.PPathConfiguration.HeaderNodeAriaRole = ""
  This.PPathConfiguration.TextNodeAriaRole = "description"
  This.PPathConfiguration.HyperlinkNodeAriaRole = "link"
End Sub

Public Function IWindowsDriverWebBrowser_GetParentWindowsDriver() As pWindowsDriver
  Set IWindowsDriverWebBrowser_GetParentWindowsDriver = This.ParentWindowsDriver
End Function

Public Sub IWindowsDriverWebBrowser_LaunchApp(ByRef ParentWindowsDriver As pWindowsDriver, WebAppName As String, WebAppTitle As String, Optional URL As String)
  
  This.PPathConfiguration.WebAppTitle = WebAppTitle
  
  Set ParentWindowsDriver = ParentWindowsDriver
  Dim InstanceType As pWinDriver.pInstanceType
  InstanceType = ParentWindowsDriver.GetWindowsDriverWebBrowserType.InstanceType
  If InstanceType = 0 Then
    InstanceType = pWinDriver.pInstanceType.ApplicationUserModelID
  End If
  Select Case InstanceType
    Case pWinDriver.pInstanceType.ApplicationUserModelID
      Phosphorus.WindowsProcesses.LaunchAppByAUMID Phosphorus.WindowsWindowsApps.DuckDuckGo, URL, Phosphorus.WindowShowStates.SW_SHOWNORMAL
    Case Else
      Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverUnhandledAppConfiguration, "Duck Duck Go, Instance Type: #" & InstanceType
  End Select
    
 'Set default path
  Dim CurrentPPath As String
'  CurrentPPath = "/Window[And(@Name=""NoAutomationPeer"",@AutomationId=""BrowserWindow"")]//Pane[@Name=""" & WebAppTitle & """]"
  CurrentPPath = "/Window[And(@Name=""NoAutomationPeer"",@AutomationId=""BrowserWindow"")]"
  ParentWindowsDriver.SetPageLoadedElement CurrentPPath, UIAutomationClient.WindowInteractionState.WindowInteractionState_ReadyForUserInteraction
End Sub

Public Function IWindowsDriverWebBrowser_GetPPathConfigurationItem(ItemType As pWinDriver.pWebBrowserPPathConfigurationItems) As Variant
  Select Case ItemType
    Case pWebBrowserPPathConfigurationItems.WebAppTitle
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.PPathConfiguration.WebAppTitle
    Case pWebBrowserPPathConfigurationItems.BrowserRootViewControlType
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.PPathConfiguration.BrowserRootViewControlType
    Case pWebBrowserPPathConfigurationItems.BrowserRootViewClassName
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.PPathConfiguration.BrowserRootViewClassName
    Case pWebBrowserPPathConfigurationItems.RootWebAreaControlType
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.PPathConfiguration.RootWebAreaControlType
    Case pWebBrowserPPathConfigurationItems.RootWebAreaAutomationID
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.PPathConfiguration.RootWebAreaAutomationID
    Case pWebBrowserPPathConfigurationItems.HeaderNodeAriaRole
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.PPathConfiguration.HeaderNodeAriaRole
    Case pWebBrowserPPathConfigurationItems.TextNodeAriaRole
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.PPathConfiguration.TextNodeAriaRole
    Case pWebBrowserPPathConfigurationItems.HyperlinkNodeAriaRole
      IWindowsDriverWebBrowser_GetPPathConfigurationItem = This.PPathConfiguration.HyperlinkNodeAriaRole
  End Select
End Function


