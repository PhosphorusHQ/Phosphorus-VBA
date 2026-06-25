VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EdgeWebBrowser"
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
  EdgeBrowserFrameViewWin As pLocator
  BrowserView As pLocator
  TopContainerView As pLocator
  EdgeToolbarView As pLocator
  BackButton As pLocator
  LocationBarView As pLocator
  AddressAndSearchBar As pLocator
  SidebarContentsSplitView As pLocator
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
  Set This.MasterWindow = Nothing
  Set This.BrowserRootView = Nothing
  Set This.NonClientView = Nothing
  Set This.EdgeBrowserFrameViewWin = Nothing
  Set This.BrowserView = Nothing
  Set This.TopContainerView = Nothing
  Set This.EdgeToolbarView = Nothing
  Set This.BackButton = Nothing
  Set This.LocationBarView = Nothing
  Set This.AddressAndSearchBar = Nothing
  Set This.SidebarContentsSplitView = Nothing
  Set This.RootWebArea = Nothing
End Sub

Private Sub GetAllLocators()
  DestroyLocators
End Sub

Private Sub DestroyLocators()
  Set This.MasterWindow = Factory.GetNewLocator
  Set This.BrowserRootView = Factory.GetNewLocator
  Set This.NonClientView = Factory.GetNewLocator
  Set This.EdgeBrowserFrameViewWin = Factory.GetNewLocator
  Set This.BrowserView = Factory.GetNewLocator
  Set This.TopContainerView = Factory.GetNewLocator
  Set This.EdgeToolbarView = Factory.GetNewLocator
  Set This.BackButton = Factory.GetNewLocator
  Set This.LocationBarView = Factory.GetNewLocator
  Set This.AddressAndSearchBar = Factory.GetNewLocator
  Set This.SidebarContentsSplitView = Factory.GetNewLocator
  Set This.RootWebArea = Factory.GetNewLocator
End Sub

Public Sub Start(WebAppName As String, URL As String, WebAppPageTitle As String)
  Toaster.Message "Starting " & WebAppName
  This.WebAppName = WebAppName
  This.URL = URL
  This.WebAppPageTitle = WebAppPageTitle
  LaunchCommandByProtocol This.WebAppName, "microsoft-edge:", This.URL, WindowShowStates.Maximized
  InitialiseAllLocators
'  This.AddressAndSearchBar.Find 10
  This.RootWebArea.Find 60
End Sub

Private Sub InitialiseAllLocators()

  'Use: AscW & ChrW to determine embedded Unicode characters
  With This.MasterWindow
    .Initialise "MasterWindow", Nothing, Children, pConditions, "AND(NameLike, ControlType, ClassName, WindowInteractionState)"
    .Condition "NameLike", Name, IsLikeTheString, This.WebAppPageTitle & "* - *" & " - Microsoft" & ChrW(8203) & " Edge"
    .ControlType UIAControlTypeIDs.Window
    .ClassName "Chrome_WidgetWin_1"
    .WindowInteractionState ReadyForUserInteraction
  End With

  This.BrowserRootView.Initialise "BrowserRootView", This.MasterWindow, Children, By.ClassName, "BrowserRootView"
  This.NonClientView.Initialise "NonClientView", This.BrowserRootView, Children, By.ClassName, "NonClientView"
  This.EdgeBrowserFrameViewWin.Initialise "EdgeBrowserFrameViewWin", This.NonClientView, Children, By.ClassName, "EdgeBrowserFrameViewWin"
  This.BrowserView.Initialise "BrowserView", This.EdgeBrowserFrameViewWin, Children, By.ClassName, "BrowserView"

    'First Pane Below Browser View
    This.TopContainerView.Initialise "TopContainerView", This.BrowserView, Children, By.ClassName, "TopContainerView"
      This.EdgeToolbarView.Initialise "EdgeToolbarView", This.TopContainerView, Children, By.ClassName, "EdgeToolbarView"

        With This.BackButton
          .Initialise "BackButton", This.EdgeToolbarView, Children, pConditions, "AND(ControlType, NameIs)": .ControlType UIAControlTypeIDs.Button: .NameIs "Back"
        End With
  
        With This.LocationBarView
          .Initialise "LocationBarView", This.EdgeToolbarView, Children, pConditions, "AND(ControlType, ClassName)": .ControlType UIAControlTypeIDs.Group: .ClassName "LocationBarView"
        End With
  
          With This.AddressAndSearchBar
            .Initialise "AddressAndSearchBar", This.LocationBarView, Children, pConditions, "AND(ControlType, NameIs)": .ControlType UIAControlTypeIDs.Edit: .NameIs "Address and search bar"
          End With
  
    'Second Pane Below Browser View
    With This.SidebarContentsSplitView
      .Initialise "SidebarContentsSplitView", This.BrowserView, Children, By.ClassName, "SidebarContentsSplitView"
    End With

      With This.RootWebArea
        .Initialise "RootWebArea", This.SidebarContentsSplitView, Descendants, By.AutomationId, "RootWebArea", FindFirst:=True
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

