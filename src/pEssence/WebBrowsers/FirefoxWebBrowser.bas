VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FirefoxWebBrowser"
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
  NavigationToolbar As pLocator
  BackButton As pLocator
  URLbar As pLocator
  URLInputBox As pLocator
  TabBrowserPanels As pLocator
  TabBrowserPanel As pLocator
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
  Set This.NavigationToolbar = Factory.GetNewLocator
  Set This.BackButton = Factory.GetNewLocator
  Set This.URLbar = Factory.GetNewLocator
  Set This.URLInputBox = Factory.GetNewLocator
  Set This.TabBrowserPanels = Factory.GetNewLocator
  Set This.TabBrowserPanel = Factory.GetNewLocator
  Set This.RootWebArea = Factory.GetNewLocator
End Sub

Private Sub DestroyLocators()
  Set This.MasterWindow = Nothing
  Set This.NavigationToolbar = Nothing
  Set This.BackButton = Nothing
  Set This.URLbar = Nothing
  Set This.URLInputBox = Nothing
  Set This.TabBrowserPanels = Nothing
  Set This.TabBrowserPanel = Nothing
  Set This.RootWebArea = Nothing
End Sub

Public Sub Start(WebAppName As String, URL As String, WebAppPageTitle As String)
  Toaster.Message "Starting " & WebAppName
  This.WebAppName = WebAppName
  This.URL = URL
  This.WebAppPageTitle = WebAppPageTitle
  'Use a pre-defined profile
'  LaunchExecutable Phosphorus.WindowsExecutables.Firefox, " -P Phosphorus -url " & This.URL, Phosphorus.WindowShowStates.Maximized
  LaunchExecutable Phosphorus.WindowsExecutables.Firefox, " -url " & This.URL, Phosphorus.WindowShowStates.Maximized
  InitialiseAllLocators
'  This.BackButton.Find 60
  This.RootWebArea.Find 60
End Sub

Private Sub InitialiseAllLocators()

  With This.MasterWindow
    .Initialise "MasterWindow", Nothing, Children, pConditions, "AND(NameIs, ControlType, ClassName, WindowInteractionState)"
    .NameIs This.WebAppPageTitle & " — Mozilla Firefox"
    .ControlType UIAControlTypeIDs.Window: .ClassName "MozillaWindowClass": .WindowInteractionState ReadyForUserInteraction
  End With

  With This.NavigationToolbar
    .Initialise "NavigationToolbar", This.MasterWindow, Children, pConditions, "AND(AriaRoleToolbar, NameIs)"
    .AriaRoleToolbar: .NameIs "Navigation"
  End With
  
    With This.BackButton
     .Initialise "BackButton", This.NavigationToolbar, Children, pConditions, _
      "AND(AriaRoleButton, OR(NameIsBackwards, NameIsBack))": .AriaRoleButton: .NameIs_ "Backwards": .NameIs_ "Back" 'Need Backwards BEFORE Back!
    End With

    With This.URLbar
     .Initialise "URLbar", This.NavigationToolbar, Children, pConditions, _
      "AND(AriaRoleGroup, AutomationId)": .AriaRoleGroup: .AutomationId "urlbar"
    End With

      With This.URLInputBox
       .Initialise "URLInputBox", This.URLbar, Descendants, pConditions, _
        "AND(AriaRoleComboBox, OR(ClassName1, ClassName2))"
        .AriaRoleComboBox
        .Condition "ClassName1", UIAProperties.ClassName, UIAPropertyComparisons.IsTheString, "urlbar-input-box"
        .Condition "ClassName2", UIAProperties.ClassName, UIAPropertyComparisons.IsTheString, "urlbar-input textbox-input" 'Firefox v 151.0.2
      End With
 
 With This.TabBrowserPanels
    .Initialise "TabBrowserPanels", This.MasterWindow, Children, pConditions, "AND(AutomationId, ControlType)"
    .AutomationId "tabbrowser-tabpanels": .ControlType Pane
  End With

  With This.TabBrowserPanel
    .Initialise "TabBrowserPanel", This.TabBrowserPanels, Children, pConditions, "AND(AriaRoleTabPanel, AutomationId)"
    .AriaRoleTabPanel: .AutomationId "panel-1-1"
  End With

  With This.RootWebArea
    .Initialise "RootWebArea", This.TabBrowserPanel, Descendants, pConditions, "AND(AriaRoleDocument, NameIs)"
    .AriaRoleDocument: .NameIs This.WebAppPageTitle
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
  With This.URLInputBox
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
