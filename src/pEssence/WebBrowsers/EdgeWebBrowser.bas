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

Implements pWB

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
  EdgeToolbarView As pLocator
  BackButton As pLocator
  LocationBarView As pLocator
  AddressAndSearchBar As pLocator
  
  SidebarContentsSplitView As pLocator
  RootWebArea As pLocator
End Type

Private This As BrowserAttributes
  
Private Sub Class_Initialize()
  Set This.MasterWindow = Factory.GetNewLocator
  Set This.BrowserRootView = Factory.GetNewLocator
  Set This.RootWebArea = Factory.GetNewLocator
  Set This.NonClientView = Factory.GetNewLocator
  Set This.BrowserFrameViewWin = Factory.GetNewLocator
  Set This.BrowserView = Factory.GetNewLocator
  Set This.TopContainerView = Factory.GetNewLocator
  Set This.EdgeToolbarView = Factory.GetNewLocator
  Set This.BackButton = Factory.GetNewLocator
  Set This.LocationBarView = Factory.GetNewLocator
  Set This.AddressAndSearchBar = Factory.GetNewLocator
  Set This.SidebarContentsSplitView = Factory.GetNewLocator
  Set This.RootWebArea = Factory.GetNewLocator
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  This.MasterWindow.Element.CloseWindow
  On Error GoTo 0
  Set This.MasterWindow = Nothing
  Set This.BrowserRootView = Nothing
  Set This.NonClientView = Nothing
  Set This.BrowserFrameViewWin = Nothing
  Set This.BrowserView = Nothing
  Set This.TopContainerView = Nothing
  Set This.EdgeToolbarView = Nothing
  Set This.BackButton = Nothing
  Set This.LocationBarView = Nothing
  Set This.AddressAndSearchBar = Nothing
  Set This.SidebarContentsSplitView = Nothing
  Set This.RootWebArea = Nothing
End Sub

Public Sub pWB_StartNormal(WebAppName As String, URL As String, WebAppPageTitle As String)
  Toaster.Message "Starting " & WebAppName
  This.WebAppName = WebAppName
  This.URL = URL
  This.WebAppPageTitle = WebAppPageTitle
  LaunchCommandByProtocol This.WebAppName, "microsoft-edge:", This.URL, WindowStyle.Maximized
  InitialiseAllLocators
  This.AddressAndSearchBar.Find 10
  This.RootWebArea.Find 10
End Sub

Private Sub InitialiseAllLocators()

  'Use: AscW & ChrW to determine embedded Unicode characters
  With This.MasterWindow
    .Initialise "MasterWindow", Nothing, Children, pConditions, "AND(NameLike, ControlType, ClassName, WindowInteractionState  )"
    .Condition "NameLike", Name, IsLikeTheString, This.WebAppPageTitle & " - " & "*" & " - Microsoft" & ChrW(8203) & " Edge"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Window
    .Condition "ClassName", ClassName, IsTheString, "Chrome_WidgetWin_1"
    .Condition "WindowInteractionState", WindowWindowInteractionState, EqualsNumber, UIAWindowInteractionStates.ReadyForUserInteraction
  End With

  With This.BrowserRootView
    .Initialise "BrowserRootView", This.MasterWindow, Children, pConditions, "ClassName"
    .Condition "ClassName", ClassName, IsTheString, "BrowserRootView"
  End With

  With This.NonClientView
    .Initialise "NonClientView", This.BrowserRootView, Children, pConditions, "ClassName"
    .Condition "ClassName", ClassName, IsTheString, "NonClientView"
  End With

  With This.BrowserFrameViewWin
    .Initialise "BrowserFrameViewWin", This.NonClientView, Children, pConditions, "ClassName"
    .Condition "ClassName", ClassName, IsTheString, "BrowserFrameViewWin"
  End With

  With This.BrowserView
    .Initialise "BrowserView", This.BrowserFrameViewWin, Children, pConditions, "ClassName"
    .Condition "ClassName", ClassName, IsTheString, "BrowserView"
  End With

    'First Pane Below Browser View
    With This.TopContainerView
      .Initialise "TopContainerView", This.BrowserView, Children, pConditions, "ClassName"
      .Condition "ClassName", ClassName, IsTheString, "TopContainerView"
    End With
  
      With This.EdgeToolbarView
        .Initialise "EdgeToolbarView", This.TopContainerView, Children, pConditions, "ClassName"
        .Condition "ClassName", ClassName, IsTheString, "EdgeToolbarView"
      End With
  
        With This.BackButton
          .Initialise "EdgeToolbarView", This.EdgeToolbarView, Children, pConditions, "AND(ControlType, NameIs)"
          .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Button
          .Condition "NameIs", Name, IsTheString, "Back"
        End With
  
        With This.LocationBarView
          .Initialise "LocationBarView", This.EdgeToolbarView, Children, pConditions, "AND(ControlType, ClassName)"
          .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Group
          .Condition "ClassName", ClassName, IsTheString, "LocationBarView"
        End With
  
          With This.AddressAndSearchBar
            .Initialise "AddressAndSearchBar", This.LocationBarView, Children, pConditions, "AND(ControlType, NameIs)"
            .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Edit
            .Condition "NameIs", Name, IsTheString, "Address and search bar"
          End With
  
    'Second Pane Below Browser View
    With This.SidebarContentsSplitView
      .Initialise "SidebarContentsSplitView", This.BrowserView, Children, pConditions, "ClassName"
      .Condition "ClassName", ClassName, IsTheString, "SidebarContentsSplitView"
    End With

      With This.RootWebArea
        .Initialise "RootWebArea", This.SidebarContentsSplitView, Descendants, pConditions, "AutomationId", FindFirst:=True
        .Condition "AutomationId", UIAProperties.AutomationId, IsTheString, "RootWebArea"
      End With

End Sub

Public Function pWB_GetRootWebArea() As pLocator
  Set pWB_GetRootWebArea = This.RootWebArea
End Function

Public Function pWB_GetCurrentURL() As String
  With This.AddressAndSearchBar
    .Find 10
    pWB_GetCurrentURL = .Element.GetValue()
  End With
End Function

Public Sub pWB_NavigateBack()
  With This.BackButton
    .Find 10
    Actions.Click .Element
    This.RootWebArea.Find 10, FindElementAgain:=True
  End With
End Sub

