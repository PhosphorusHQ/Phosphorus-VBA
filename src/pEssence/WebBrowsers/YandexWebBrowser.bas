VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "YandexWebBrowser"
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
  BrowserRootView2 As pLocator
  CustoTopContainerViewPlaceholderView As pLocator
  TabsAccessiblePaneView As pLocator
  FirstTab As pLocator
  CloseOtherTabs As pLocator
  NonClientView As pLocator
  FrameView As pLocator
  BrowserView As pLocator
  View As pLocator
  RootWebArea As pLocator
  RootView4 As pLocator
  SuggestContents As pLocator
  AddressAndSearchBar As pLocator
  TitleBarView As pLocator
  BackButton As pLocator
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
  Set This.BrowserRootView2 = Factory.GetNewLocator
  Set This.NonClientView = Factory.GetNewLocator
  Set This.FrameView = Factory.GetNewLocator
  Set This.BrowserView = Factory.GetNewLocator
  Set This.View = Factory.GetNewLocator
  Set This.RootWebArea = Factory.GetNewLocator
  Set This.CustoTopContainerViewPlaceholderView = Factory.GetNewLocator
  Set This.TabsAccessiblePaneView = Factory.GetNewLocator
  Set This.FirstTab = Factory.GetNewLocator
  Set This.CloseOtherTabs = Factory.GetNewLocator
  Set This.RootView4 = Factory.GetNewLocator
  Set This.SuggestContents = Factory.GetNewLocator
  Set This.AddressAndSearchBar = Factory.GetNewLocator
  Set This.TitleBarView = Factory.GetNewLocator
  Set This.BackButton = Factory.GetNewLocator
End Sub

Private Sub DestroyLocators()
  Set This.MasterWindow = Nothing
  Set This.BrowserRootView = Nothing
  Set This.BrowserRootView2 = Nothing
  Set This.NonClientView = Nothing
  Set This.NonClientView = Nothing
  Set This.BrowserView = Nothing
  Set This.View = Nothing
  Set This.RootWebArea = Nothing
  Set This.CustoTopContainerViewPlaceholderView = Nothing
  Set This.TabsAccessiblePaneView = Nothing
  Set This.FirstTab = Nothing
  Set This.CloseOtherTabs = Nothing
  Set This.RootView4 = Nothing
  Set This.SuggestContents = Nothing
  Set This.AddressAndSearchBar = Nothing
  Set This.TitleBarView = Nothing
  Set This.BackButton = Nothing
End Sub

Public Sub Start(WebAppName As String, URL As String, WebAppPageTitle As String, Optional BaseWaitTimeSeconds As Long = 10, Optional AbsoluteWaitTimeSeconds As Long)
  Toaster.Message "Starting " & WebAppName
  This.WebAppName = WebAppName
  This.URL = URL
  This.WebAppPageTitle = WebAppPageTitle
  'We need to disable Background Running in settings (browser://settings/)
  LaunchExecutable Phosphorus.WindowsExecutables.YandexWebBrowser, " --force-renderer-accessibility=complete " & URL, WindowShowStates.Maximized
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
    .NameIs This.WebAppPageTitle & " — Yandex" & Chr(160) & "Browser" 'Uses the non-breaking space
    .ControlType UIAControlTypeIDs.Window
    .ClassName "Chrome_Yandex_WidgetWin_1"
    .WindowInteractionState ReadyForUserInteraction
  End With

  With This.BrowserRootView
    .Initialise "BrowserRootView", This.MasterWindow, Children, By.ClassName, "BrowserRootView"
  End With

    With This.BrowserRootView2
      .Initialise "BrowserRootView2", This.BrowserRootView, Children, By.ClassName, "BrowserRootView"
    End With

      With This.CustoTopContainerViewPlaceholderView
        .Initialise "CustoTopContainerView::PlaceholderView", This.BrowserRootView2, Children, By.ClassName, "CustoTopContainerView::PlaceholderView"
        .PositionInMatchingSet 1
      End With
        
        This.TabsAccessiblePaneView.Initialise "TabsAccessiblePaneView", This.CustoTopContainerViewPlaceholderView, Children, By.ClassName, "TabsAccessiblePaneView"
        
        With This.FirstTab
          .Initialise "FirstTab", This.TabsAccessiblePaneView, Children, By.pConditions, "ControlType"
          .ControlType UIAControlTypeIDs.TabItem
          .PositionInMatchingSet 1
        End With

        'Right click & close all other tabs
        This.FirstTab.Element.RightClick
        With This.CloseOtherTabs
          .Initialise "CloseOtherTabs", This.MasterWindow, Descendants, By.pConditions, "AND(ControlType, NameIs)", FindFirst:=True
          .ControlType MenuItem: .NameIs "Close other tabs"
          .Element.ClickIfEnabled This.MasterWindow.Element
        End With
            
      With This.TitleBarView
        .Initialise "TitleBarView", This.BrowserRootView2, Descendants, By.pConditions, "AND(AriaRoleToolbar, ClassName)", FindFirst:=True
        .AriaRoleToolbar: .ClassName "TitleBarView"
      End With

        This.BackButton.Initialise "BackButton", This.TitleBarView, Children, By.NameIs, "Back"

  This.NonClientView.Initialise "NonClientView", This.BrowserRootView, Children, By.ClassName, "NonClientView"

    This.FrameView.Initialise "FrameView", This.NonClientView, Children, By.ClassName, "FrameView"

      This.BrowserView.Initialise "BrowserView", This.FrameView, Children, By.ClassName, "BrowserView"

        This.View.Initialise "View", This.BrowserView, Children, By.ClassName, "View"

          This.RootWebArea.Initialise "RootWebArea", This.View, Descendants, By.AutomationId, "RootWebArea", True

  With This.RootView4
    .Initialise "RootView4", This.BrowserRootView, Children, By.ClassName, "RootView": .PositionInMatchingSet 4
  End With
  
    With This.SuggestContents
      .Initialise "SuggestContents", This.RootView4, Children, By.ClassName, "SuggestContents"
    End With

      With This.AddressAndSearchBar
        .Initialise "AddressAndSearchBar", This.SuggestContents, Descendants, By.pConditions, "AND(AriaRoleTextBox, NameIs)"
        .AriaRoleTextBox: .NameIs "Address and search bar"
      End With

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

