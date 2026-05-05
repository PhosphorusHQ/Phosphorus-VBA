VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OperaWebBrowser"
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
  RootView As pLocator
  NonClientView As pLocator
  BrowserNonClient As pLocator
  BrowserClientView As pLocator
  LiveBackgroundView As pLocator
  DefaultContentWrapper As pLocator
  SidebarItemContentViewDockerView As pLocator
  View As pLocator
  TopBarContainerView As pLocator
  LastTabView As pLocator
  CloseOtherTabs As pLocator
  ToolbarView As pLocator
  BackButton As pLocator
  PageContainerView As pLocator
  RootWebArea As pLocator
End Type

Private This As BrowserAttributes

Private Sub Class_Initialize()
  Set This.MasterWindow = Factory.GetNewLocator
  Set This.RootView = Factory.GetNewLocator
  Set This.NonClientView = Factory.GetNewLocator
  Set This.BrowserNonClient = Factory.GetNewLocator
  Set This.BrowserClientView = Factory.GetNewLocator
  Set This.LiveBackgroundView = Factory.GetNewLocator
  Set This.DefaultContentWrapper = Factory.GetNewLocator
  Set This.SidebarItemContentViewDockerView = Factory.GetNewLocator
  Set This.View = Factory.GetNewLocator
  Set This.TopBarContainerView = Factory.GetNewLocator
  Set This.LastTabView = Factory.GetNewLocator
  Set This.CloseOtherTabs = Factory.GetNewLocator
  Set This.ToolbarView = Factory.GetNewLocator
  Set This.BackButton = Factory.GetNewLocator
  Set This.PageContainerView = Factory.GetNewLocator
  Set This.RootWebArea = Factory.GetNewLocator
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  This.MasterWindow.Element.CloseWindow
  On Error GoTo 0
  Set This.MasterWindow = Nothing
  Set This.RootView = Nothing
  Set This.NonClientView = Nothing
  Set This.BrowserNonClient = Nothing
  Set This.BrowserClientView = Nothing
  Set This.LiveBackgroundView = Nothing
  Set This.DefaultContentWrapper = Nothing
  Set This.SidebarItemContentViewDockerView = Nothing
  Set This.View = Nothing
  Set This.TopBarContainerView = Nothing
  Set This.LastTabView = Nothing
  Set This.CloseOtherTabs = Nothing
  Set This.ToolbarView = Nothing
  Set This.BackButton = Nothing
  Set This.PageContainerView = Nothing
  Set This.RootWebArea = Nothing
End Sub

Public Sub Start(WebAppName As String, URL As String, WebAppPageTitle As String)
  Toaster.Message "Starting " & WebAppName
  This.WebAppName = WebAppName
  This.URL = URL
  This.WebAppPageTitle = WebAppPageTitle
  LaunchExecutable Phosphorus.WindowsExecutables.Opera, "--force-renderer-accessibility " & URL, WindowShowStates.Maximized
  InitialiseAllLocators
'  This.AddressAndSearchBar.Find 10
'  This.RootWebArea.Find 10
End Sub

Private Sub InitialiseAllLocators()

  With This.MasterWindow
    .Initialise "MasterWindow", Nothing, Children, pConditions, "AND(NameIs, ControlType, ClassName, WindowInteractionState)"
    .NameIs This.WebAppPageTitle & " - Opera"
    .ControlType UIAControlTypeIDs.Window
    .ClassName "Chrome_WidgetWin_1"
    .WindowInteractionState ReadyForUserInteraction
.Find 10
  End With

  With This.RootView
    .Initialise "RootView", This.MasterWindow, Children, By.ClassName, "RootView"
.Find 10
  End With

  This.NonClientView.Initialise "NonClientView", This.RootView, Children, By.ClassName, "NonClientView"
This.NonClientView.Find

  This.BrowserNonClient.Initialise "BrowserNonClient", This.NonClientView, Children, By.NameIs, "Browser non-client"
This.BrowserNonClient.Find

  This.BrowserClientView.Initialise "BrowserClientView", This.BrowserNonClient, Children, By.ClassName, "BrowserClientView"
This.BrowserClientView.Find

  This.LiveBackgroundView.Initialise "LiveBackgroundView", This.BrowserClientView, Children, By.ClassName, "LiveBackgroundView"
This.LiveBackgroundView.Find

  This.DefaultContentWrapper.Initialise "DefaultContentWrapper", This.LiveBackgroundView, Children, By.ClassName, "DefaultContentWrapper"
This.DefaultContentWrapper.Find

  This.SidebarItemContentViewDockerView.Initialise "SidebarItemContentViewDockerView", This.DefaultContentWrapper, Children, By.ClassName, "SidebarItemContentViewDockerView"
This.SidebarItemContentViewDockerView.Find

    This.View.Initialise "View", This.SidebarItemContentViewDockerView, Children, By.ClassName, "View"
This.View.Find

      This.TopBarContainerView.Initialise "TopBarContainerView", This.View, Descendants, By.ClassName, "TopBarContainerView", FindFirst:=True
 This.TopBarContainerView.Find

        With This.LastTabView
          .Initialise "BackButton", This.TopBarContainerView, Descendants, By.pConditions, "AND(ClassName, NameIs)", FindFirst:=True
          .ClassName "TabView": .NameIs This.WebAppPageTitle: .PositionInMatchingSet -1
.Find
        End With

        
'Right click & close all other tabs
This.LastTabView.Element.RightClick
'Last tab?
This.MasterWindow.ListAllChildren

'"/Pane[1]//MenuBar//MenuItem[@Name=""Close other tabs""]"

        With This.CloseOtherTabs
          .Initialise "CloseOtherTabs", This.MasterWindow, Descendants, By.pConditions, "AND(ControlType, NameIs)", FindFirst:=True
          .ControlType MenuItem: .NameIs "Close other tabs"
.Find 10
        End With
This.CloseOtherTabs.Element.Click

Stop
'Search for element by name, then use a ListAllAncestors method and ListClildren to get find the best navigation path for transient elements!?

      This.ToolbarView.Initialise "ToolbarView", This.View, Descendants, By.ClassName, "ToolbarView", FindFirst:=True
 This.ToolbarView.Find

        With This.BackButton
          .Initialise "BackButton", This.ToolbarView, Descendants, By.pConditions, "AND(AriaRoleButton, NameIs)", FindFirst:=True
          .AriaRoleButton: .NameIs "Back"
.Find
        End With
 
Stop
     With This.PageContainerView
        .Initialise "PageContainerView", This.View, Descendants, By.pConditions, "AND(ClassName, NameIs)", FindFirst:=True
        .ClassName "PageContainerView": .NameIs "Page container"
.Find
      End With
        
         This.RootWebArea.Initialise "RootWebArea", This.PageContainerView, Descendants, By.AutomationId, "RootWebArea", FindFirst:=True
This.RootWebArea.Find

End Sub
