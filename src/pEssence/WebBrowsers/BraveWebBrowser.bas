VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BraveWebBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder WebBrowsers
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Private Type BrowserAttributes
  WebAppName As String
  URL As String
  WebAppPageTitle As String
  MasterWindow As pLocator
  BrowserRootView As pLocator
  NonClientView As pLocator
  BrowserFrameViewWin As pLocator
  BrowserView As pLocator
  TopContainerView As pLocator
  BraveHorizontalTabStripRegionView As pLocator
  BraveTabStrip As pLocator
  LastTab As pLocator
  CloseOtherTabs As pLocator
  ToolbarView As pLocator
  BackButton As pLocator
  LocationBarView As pLocator
  AddressAndSearchBar As pLocator
  BrowserViewSubView1 As pLocator
  RootWebArea As pLocator
End Type

Private This As BrowserAttributes

Private Sub Class_Initialize()
  GetAllLocators
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  This.MasterWindow.Element.CloseWindow
  On Error GoTo 0
  DestroyLocators
End Sub

Private Sub GetAllLocators()
  Set This.MasterWindow = Factory.GetNewLocator
  Set This.BrowserRootView = Factory.GetNewLocator
  Set This.NonClientView = Factory.GetNewLocator
  Set This.BrowserFrameViewWin = Factory.GetNewLocator
  Set This.BrowserView = Factory.GetNewLocator
  Set This.TopContainerView = Factory.GetNewLocator
  Set This.BraveHorizontalTabStripRegionView = Factory.GetNewLocator
  Set This.BraveTabStrip = Factory.GetNewLocator
  Set This.LastTab = Factory.GetNewLocator
  Set This.CloseOtherTabs = Factory.GetNewLocator
  Set This.ToolbarView = Factory.GetNewLocator
  Set This.BackButton = Factory.GetNewLocator
  Set This.LocationBarView = Factory.GetNewLocator
  Set This.AddressAndSearchBar = Factory.GetNewLocator
  Set This.BrowserViewSubView1 = Factory.GetNewLocator
  Set This.RootWebArea = Factory.GetNewLocator
End Sub

Private Sub DestroyLocators()
  Set This.MasterWindow = Nothing
  Set This.BrowserRootView = Nothing
  Set This.NonClientView = Nothing
  Set This.BrowserFrameViewWin = Nothing
  Set This.BrowserView = Nothing
  Set This.TopContainerView = Nothing
  Set This.BraveHorizontalTabStripRegionView = Nothing
  Set This.BraveTabStrip = Nothing
  Set This.LastTab = Nothing
  Set This.CloseOtherTabs = Nothing
  Set This.ToolbarView = Nothing
  Set This.BackButton = Nothing
  Set This.LocationBarView = Nothing
  Set This.AddressAndSearchBar = Nothing
  Set This.BrowserViewSubView1 = Nothing
  Set This.RootWebArea = Nothing
End Sub

'Download latest stable Chromium binaries (64-bit and 32-bit)
'https://chromium.woolyss.com/

'How to install Chromium for all users on Windows
'https://martinrotter.github.io/it-programming/2016/07/17/install-chromium-system-wide-windows/

Public Sub Start(WebAppName As String, URL As String, WebAppPageTitle As String, Optional BaseWaitTimeSeconds As Long = 10, Optional AbsoluteWaitTimeSeconds As Long)
  Toaster.Message "Starting " & WebAppName
  This.WebAppName = WebAppName
  This.URL = URL
  This.WebAppPageTitle = WebAppPageTitle
  LaunchExecutable Phosphorus.WindowsExecutables.Brave, "--force-renderer-accessibility " & URL, WindowShowStates.Maximized
  InitialiseAllLocators
  If AbsoluteWaitTimeSeconds > 0 Then
    This.RootWebArea.Find AbsoluteWaitTimeSeconds
  Else
    This.RootWebArea.Find BaseWaitTimeSeconds * (1000 / WebBrowserCommon.DownloadSpeedMbps)
  End If
End Sub

Private Sub InitialiseAllLocators()

  With This.MasterWindow
    .Initialise "MasterWindow", Nothing, Children, pConditions, "AND(NameIs, ControlType, ClassName, WindowInteractionState)"
    .NameIs This.WebAppPageTitle & " - Brave"
    .ControlType UIAControlTypeIDs.Window
    .ClassName "Chrome_WidgetWin_1"
    .WindowInteractionState ReadyForUserInteraction
  End With

  With This.BrowserRootView
    .Initialise "BrowserRootView", This.MasterWindow, Children, By.ClassName, "BraveBrowserRootView"
  End With

  This.NonClientView.Initialise "NonClientView", This.BrowserRootView, Children, By.ClassName, "NonClientView"
  This.BrowserFrameViewWin.Initialise "BrowserFrameViewWin", This.NonClientView, Children, By.ClassName, "BrowserFrameViewWin"
  This.BrowserView.Initialise "BrowserView", This.BrowserFrameViewWin, Children, By.ClassName, "BraveBrowserView"

    'First Pane Below Browser View
    This.TopContainerView.Initialise "TopContainerView", This.BrowserView, Children, By.ClassName, "TopContainerView"

      This.ToolbarView.Initialise "ToolbarView", This.TopContainerView, Children, By.ClassName, "BraveToolbarView"

        With This.BackButton
          .Initialise "BackButton", This.ToolbarView, Children, pConditions, "AND(ControlType, NameIs)": .ControlType UIAControlTypeIDs.Button: .NameIs "Back"
        End With
        
        With This.LocationBarView
          .Initialise "LocationBarView", This.ToolbarView, Children, pConditions, "AND(ControlType, ClassName)": .ControlType UIAControlTypeIDs.Group: .ClassName "BraveLocationBarView"
        End With
          
          With This.AddressAndSearchBar
            .Initialise "AddressAndSearchBar", This.LocationBarView, Children, pConditions, "AND(ControlType, NameIs)": .ControlType UIAControlTypeIDs.Edit: .NameIs "Address and search bar"
          End With

      This.BraveHorizontalTabStripRegionView.Initialise "BraveHorizontalTabStripRegionView", This.BrowserView, Children, By.ClassName, "BraveHorizontalTabStripRegionView"

      With This.BraveTabStrip
        .Initialise "BraveTabStrip", This.BraveHorizontalTabStripRegionView, Children, By.ClassName, "BraveTabStrip"
      End With

        With This.LastTab
          .Initialise "LastTab", This.BraveTabStrip, Descendants, By.pConditions, "AND(ClassName, NameIs)", FindFirst:=True
          .ClassName "BraveTab": .NameIs This.WebAppPageTitle: .PositionInMatchingSet -1
        End With

      'Right click & close all other tabs
      This.LastTab.Element.RightClick
      With This.CloseOtherTabs
        .Initialise "CloseOtherTabs", This.MasterWindow, Descendants, By.pConditions, "AND(ControlType, NameIs)", FindFirst:=True
        .ControlType MenuItem: .NameIs "Close other tabs"
        .Element.ClickIfEnabled This.MasterWindow.Element
      End With

    'Second Pane Below Browser View
    With This.BrowserViewSubView1
      .Initialise "BrowserViewSubView1", This.BrowserView, Children, By.ClassName, "View", FindFirst:=True
    End With

      This.RootWebArea.Initialise "RootWebArea", This.BrowserViewSubView1, Descendants, By.AutomationId, "RootWebArea", FindFirst:=True

End Sub

Public Function GetRootWebArea(Optional NewWebPage As Boolean) As pLocator
  If NewWebPage Then
    DestroyLocators
    GetAllLocators
    InitialiseAllLocators
    This.RootWebArea.Find 10
  End If
  Set GetRootWebArea = This.RootWebArea
End Function

Public Function GetCurrentURL() As String
  With This.AddressAndSearchBar
    .Find 10
    GetCurrentURL = .Element.GetValue()
  End With
End Function

Public Sub NavigateBack()
  With This.BackButton
    .Find 10
    .Element.Click
    This.RootWebArea.Find 10, FindElementAgain:=True
  End With
End Sub

