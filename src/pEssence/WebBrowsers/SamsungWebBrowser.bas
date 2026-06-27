VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SamsungWebBrowser"
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
  LockableBrowserView As pLocator
  MainContainerView As pLocator
  TopContainerView As pLocator
  ToolbarView As pLocator
  ToolbarViewContainerView As pLocator
  LocationBarContainer As pLocator
  LocationBarView As pLocator
  AddressBox As pLocator
  View As pLocator
  BackButton As pLocator
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
  Set This.LockableBrowserView = Factory.GetNewLocator
  Set This.MainContainerView = Factory.GetNewLocator
  Set This.TopContainerView = Factory.GetNewLocator
  Set This.ToolbarView = Factory.GetNewLocator
  Set This.ToolbarViewContainerView = Factory.GetNewLocator
  Set This.LocationBarContainer = Factory.GetNewLocator
  Set This.LocationBarView = Factory.GetNewLocator
  Set This.AddressBox = Factory.GetNewLocator
  Set This.View = Factory.GetNewLocator
  Set This.BackButton = Factory.GetNewLocator
  Set This.RootWebArea = Factory.GetNewLocator
End Sub

Private Sub DestroyLocators()
  Set This.MasterWindow = Nothing
  Set This.BrowserRootView = Nothing
  Set This.NonClientView = Nothing
  Set This.BrowserFrameViewWin = Nothing
  Set This.LockableBrowserView = Nothing
  Set This.MainContainerView = Nothing
  Set This.TopContainerView = Nothing
  Set This.ToolbarView = Nothing
  Set This.ToolbarViewContainerView = Nothing
  Set This.LocationBarContainer = Nothing
  Set This.LocationBarView = Nothing
  Set This.AddressBox = Nothing
  Set This.View = Nothing
  Set This.BackButton = Nothing
  Set This.RootWebArea = Nothing
End Sub

Public Sub Start(WebAppName As String, URL As String, WebAppPageTitle As String, Optional BaseWaitTimeSeconds As Long = 10, Optional AbsoluteWaitTimeSeconds As Long)
  Toaster.Message "Starting " & WebAppName
  This.WebAppName = WebAppName
  This.URL = URL
  This.WebAppPageTitle = WebAppPageTitle
  LaunchExecutable Phosphorus.WindowsExecutables.SamsungBrowser, "--force-renderer-accessibility " & URL, WindowShowStates.Maximized
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
    .NameIs This.WebAppPageTitle & " - Samsung Browser"
    .ControlType UIAControlTypeIDs.Window
    .ClassName "Chrome_WidgetWin_1"
    .WindowInteractionState ReadyForUserInteraction
  End With

  With This.BrowserRootView
    .Initialise "BrowserRootView", This.MasterWindow, Children, By.ClassName, "BrowserRootView"
  End With

  This.NonClientView.Initialise "NonClientView", This.BrowserRootView, Children, By.ClassName, "NonClientView"

  This.BrowserFrameViewWin.Initialise "BrowserFrameViewWin", This.NonClientView, Children, By.ClassName, "BrowserFrameViewWin"

  This.LockableBrowserView.Initialise "LockableBrowserView", This.BrowserFrameViewWin, Children, By.ClassName, "LockableBrowserView"

  This.MainContainerView.Initialise "MainContainerView", This.LockableBrowserView, Children, By.ClassName, "MainContainerView"

    This.TopContainerView.Initialise "TopContainerView", This.MainContainerView, Children, By.ClassName, "TopContainerView"
      
      This.ToolbarView.Initialise "ToolbarView", This.TopContainerView, Children, By.ClassName, "ToolbarView"

        This.ToolbarViewContainerView.Initialise "ToolbarViewContainerView", This.ToolbarView, Children, By.ClassName, "ToolbarView::ContainerView"

          This.LocationBarContainer.Initialise "LocationBarContainer", This.ToolbarViewContainerView, Children, By.ClassName, "LocationBarContainer"

            This.LocationBarView.Initialise "LocationBarView", This.LocationBarContainer, Children, By.ClassName, "LocationBarView"

              This.AddressBox.Initialise "AddressBox", This.LocationBarView, Children, By.ClassName, "OmniboxViewViews"
      
        This.BackButton.Initialise "BackButton", This.ToolbarView, Descendants, By.NameIs, "Back", True

    This.View.Initialise "View", This.MainContainerView, Children, By.ClassName, "View"

    This.RootWebArea.Initialise "RootWebArea", This.View, Descendants, By.AutomationId, "RootWebArea", True

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
  With This.AddressBox
    .Find 10
    GetCurrentURL = .Element.GetProperty(UIAProperties.Name)
    GetCurrentURL = Replace(GetCurrentURL, "Address and search bar, ", "")
  End With
End Function

Public Sub NavigateBack()
  With This.BackButton
    .Find 10
    .Element.Click
    This.RootWebArea.Find 10, FindElementAgain:=True
  End With
End Sub
