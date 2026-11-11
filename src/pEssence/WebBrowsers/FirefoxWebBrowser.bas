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
  AllTabs As pLocator
  LastTab As pLocator
  CloseMultipleTabs As pLocator
  CloseOtherTabs As pLocator
  BackButton As pLocator
  RefreshButton As pLocator
  URLInputBox As pLocator
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
  Set This.AllTabs = Factory.GetNewLocator
  Set This.LastTab = Factory.GetNewLocator
  Set This.CloseMultipleTabs = Factory.GetNewLocator
  Set This.CloseOtherTabs = Factory.GetNewLocator
  Set This.BackButton = Factory.GetNewLocator
  Set This.RefreshButton = Factory.GetNewLocator
  Set This.URLInputBox = Factory.GetNewLocator
  Set This.RootWebArea = Factory.GetNewLocator
End Sub

Private Sub DestroyLocators()
  Set This.MasterWindow = Nothing
  Set This.AllTabs = Nothing
  Set This.LastTab = Nothing
  Set This.CloseMultipleTabs = Nothing
  Set This.CloseOtherTabs = Nothing
  Set This.BackButton = Nothing
  Set This.RefreshButton = Nothing
  Set This.URLInputBox = Nothing
  Set This.RootWebArea = Nothing
End Sub

Public Sub Start(WebAppName As String, URL As String, WebAppPageTitle As String, Optional BaseWaitTimeSeconds As Long = 10, Optional AbsoluteWaitTimeSeconds As Long)
  Toaster.Message "Starting " & WebAppName
  This.WebAppName = WebAppName
  This.URL = URL
  This.WebAppPageTitle = WebAppPageTitle
  'Use a pre-defined profile
'  LaunchExecutable Phosphorus.WindowsExecutables.Firefox, " -P Phosphorus -url " & This.URL, Phosphorus.WindowShowStates.Maximized
  LaunchExecutable Phosphorus.WindowsExecutables.Firefox, " -url " & This.URL, Phosphorus.WindowShowStates.Maximized
  InitialiseAllLocators
  If AbsoluteWaitTimeSeconds = 0 Then
    AbsoluteWaitTimeSeconds = BaseWaitTimeSeconds
  End If
  If AbsoluteWaitTimeSeconds >= 0 Then
    This.RootWebArea.Find AbsoluteWaitTimeSeconds
  Else
    This.RootWebArea.Find BaseWaitTimeSeconds * (1000 / WebBrowserCommon.DownloadSpeedMbps)
  End If
End Sub

Private Sub InitialiseAllLocators()

  With This.MasterWindow
    .Initialise "MasterWindow", Nothing, Children, pConditions, "AND(NameIs, ControlType, ClassName, WindowInteractionState)"
    .NameIs This.WebAppPageTitle & " — Mozilla Firefox"
    .ControlType UIAControlTypeIDs.Window: .ClassName "MozillaWindowClass": .WindowInteractionState ReadyForUserInteraction
    .PositionInMatchingSet -1
  End With
  
  With This.BackButton
   .Initialise "BackButton", This.MasterWindow, Descendants, pConditions, _
    "AND(AriaRoleButton, OR(NameIsBackwards, NameIsBack))": .AriaRoleButton: .NameIs_ "Backwards": .NameIs_ "Back" 'Need Backwards BEFORE Back!
  End With

  This.RefreshButton.Initialise "RefreshButton", This.MasterWindow, Descendants, By.NameIs, "Reload"

  With This.URLInputBox
   .Initialise "URLInputBox", This.MasterWindow, Descendants, pConditions, _
     "AND(AriaRoleComboBox, OR(ClassName1, ClassName2))"
   .AriaRoleComboBox
   .Condition "ClassName1", UIAProperties.ClassName, UIAPropertyComparisons.IsTheString, "urlbar-input-box"
   .Condition "ClassName2", UIAProperties.ClassName, UIAPropertyComparisons.IsTheString, "urlbar-input textbox-input" 'Firefox v 151.0.2
  End With

  This.MasterWindow.Find 10
  With This.AllTabs
    .Initialise "AllTabs", This.MasterWindow, Descendants, By.ControlType, UIAControlTypeIDs.TabItem
    .FindAll 10
  End With
  
  With This.LastTab
    .Initialise "LastTab", This.MasterWindow, Descendants, By.pConditions, "AND(ControlType, NameIs)", FindFirst:=True
    .ControlType UIAControlTypeIDs.TabItem: .NameIs This.WebAppPageTitle: .PositionInMatchingSet -1
  End With
    
  If This.AllTabs.NumberOfMatchingElements > 1 Then
  
    'Right click & close all other tabs
    This.LastTab.Element.RightClick

    With This.CloseMultipleTabs
      .Initialise "CloseMultipleTabs", This.MasterWindow, Descendants, By.pConditions, "AND(ControlType, NameIs)", FindFirst:=True
      .ControlType MenuItem: .NameIs "Close Multiple Tabs"
      .Find 10
      .Element.Click
    End With
       
    With This.CloseOtherTabs
      .Initialise "CloseOtherTabs", This.CloseMultipleTabs, Descendants, By.pConditions, "AND(ControlType, NameIs)", FindFirst:=True
      .ControlType MenuItem: .NameIs "Close Other Tabs"
      .Find 10
      .Element.Click
    End With
 
  End If
  
  With This.RootWebArea
    .Initialise "RootWebArea", This.MasterWindow, Descendants, pConditions, "AND(AriaRoleDocument, NameIs)"
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

Public Function GetRefreshButton() As pLocator
  If This.RefreshButton.Element.UIAElement Is Nothing Then
    This.RefreshButton.Find 0
  End If
  Set GetRefreshButton = This.RefreshButton
End Function

Public Function GetCurrentURL() As String
  With This.URLInputBox
    .Find 10
    GetCurrentURL = .Element.Value()
  End With
End Function

Public Sub NavigateBack()
  WebBrowserCommon.Navigate Me, This.BackButton, This.URLInputBox, This.RootWebArea
End Sub

Public Sub WaitForNewURL(TimeoutInSeconds As Integer)
  WebBrowserCommon.WaitForNewURL GetCurrentURL, This.URLInputBox, This.RootWebArea, TimeoutInSeconds
End Sub

