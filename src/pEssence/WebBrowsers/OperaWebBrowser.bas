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
  AddressBar As pLocator
  AddressField As pLocator
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
  Set This.AddressBar = Factory.GetNewLocator
  Set This.AddressField = Factory.GetNewLocator
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
  Set This.AddressBar = Nothing
  Set This.AddressField = Nothing
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
  This.RootWebArea.Find 10
End Sub

Private Sub InitialiseAllLocators()

  With This.MasterWindow
    .Initialise "MasterWindow", Nothing, Children, pConditions, "AND(NameIs, ControlType, ClassName, WindowInteractionState)"
    .NameIs This.WebAppPageTitle & " - Opera"
    .ControlType UIAControlTypeIDs.Window
    .ClassName "Chrome_WidgetWin_1"
    .WindowInteractionState ReadyForUserInteraction
  End With

  With This.RootView
    .Initialise "RootView", This.MasterWindow, Children, By.ClassName, "RootView"
  End With

  This.NonClientView.Initialise "NonClientView", This.RootView, Children, By.ClassName, "NonClientView"

  This.BrowserNonClient.Initialise "BrowserNonClient", This.NonClientView, Children, By.NameIs, "Browser non-client"

  This.BrowserClientView.Initialise "BrowserClientView", This.BrowserNonClient, Children, By.ClassName, "BrowserClientView"

  This.LiveBackgroundView.Initialise "LiveBackgroundView", This.BrowserClientView, Children, By.ClassName, "LiveBackgroundView"

  This.DefaultContentWrapper.Initialise "DefaultContentWrapper", This.LiveBackgroundView, Children, By.ClassName, "DefaultContentWrapper"

  This.SidebarItemContentViewDockerView.Initialise "SidebarItemContentViewDockerView", This.DefaultContentWrapper, Children, By.ClassName, "SidebarItemContentViewDockerView"

    This.View.Initialise "View", This.SidebarItemContentViewDockerView, Children, By.ClassName, "View"

      This.TopBarContainerView.Initialise "TopBarContainerView", This.View, Descendants, By.ClassName, "TopBarContainerView", FindFirst:=True

        With This.LastTabView
          .Initialise "BackButton", This.TopBarContainerView, Descendants, By.pConditions, "AND(ClassName, NameIs)", FindFirst:=True
          .ClassName "TabView": .NameIs This.WebAppPageTitle: .PositionInMatchingSet -1
        End With

        'Right click & close all other tabs
        This.LastTabView.Element.RightClick
        With This.CloseOtherTabs
          .Initialise "CloseOtherTabs", This.MasterWindow, Descendants, By.pConditions, "AND(ControlType, NameIs)", FindFirst:=True
          .ControlType MenuItem: .NameIs "Close other tabs"
        End With
        This.CloseOtherTabs.Element.Click

      With This.ToolbarView
        .Initialise "ToolbarView", This.View, Descendants, By.pConditions, "AND(ClassName, NameIs)", FindFirst:=True
        .ClassName "ToolbarView": .NameIs "Navigation"
      End With
      
        With This.BackButton
          .Initialise "BackButton", This.ToolbarView, Descendants, By.pConditions, "AND(AriaRoleButton, NameIs)", FindFirst:=True
          .AriaRoleButton: .NameIs "Back"
        End With
 
        With This.AddressBar
          .Initialise "AddressBar", This.ToolbarView, Descendants, By.pConditions, "AND(AriaRoleTextBox, ClassName, NameIs)", FindFirst:=True
          .AriaRoleTextBox: .ClassName "AddressBarView": .NameIs "Address bar"
        End With

          With This.AddressField
            .Initialise "AddressField", This.AddressBar, Descendants, By.pConditions, "AND(AriaRoleTextBox, ClassName, NameIs)", FindFirst:=True
            .AriaRoleTextBox: .ClassName "AddressTextfieldView": .NameIs "Address field"
          End With

    With This.PageContainerView
      .Initialise "PageContainerView", This.View, Descendants, By.pConditions, "AND(ClassName, NameIs)", FindFirst:=True
      .ClassName "PageContainerView": .NameIs "Page container"
    End With
        
         This.RootWebArea.Initialise "RootWebArea", This.PageContainerView, Descendants, By.AutomationId, "RootWebArea", FindFirst:=True
   
End Sub

Public Function GetRootWebArea() As pLocator
  Set GetRootWebArea = This.RootWebArea
End Function

Public Function GetCurrentURL() As String
  With This.AddressField
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

