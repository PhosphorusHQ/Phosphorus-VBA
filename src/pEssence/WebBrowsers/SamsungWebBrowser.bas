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
  AddressBox As pLocator
  BackButton As pLocator
  RefreshButton As pLocator
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
  Set This.AddressBox = Factory.GetNewLocator
  Set This.BackButton = Factory.GetNewLocator
  Set This.RefreshButton = Factory.GetNewLocator
  Set This.RootWebArea = Factory.GetNewLocator
End Sub

Private Sub DestroyLocators()
  Set This.MasterWindow = Nothing
  Set This.AddressBox = Nothing
  Set This.BackButton = Nothing
  Set This.RefreshButton = Nothing
  Set This.RootWebArea = Nothing
End Sub

Public Sub Start(WebAppName As String, URL As String, WebAppPageTitle As String, Optional BaseWaitTimeSeconds As Long = 10, Optional AbsoluteWaitTimeSeconds As Long)
  Toaster.Message "Starting " & WebAppName
  This.WebAppName = WebAppName
  This.URL = URL
  This.WebAppPageTitle = WebAppPageTitle
  LaunchExecutable Phosphorus.WindowsExecutables.SamsungBrowser, "--force-renderer-accessibility " & URL, WindowShowStates.Maximized
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
    .NameIs This.WebAppPageTitle & " - Samsung Browser"
    .ControlType UIAControlTypeIDs.Window
    .ClassName "Chrome_WidgetWin_1"
    .WindowInteractionState ReadyForUserInteraction
  End With

  This.AddressBox.Initialise "AddressBox", This.MasterWindow, Descendants, By.ClassName, "OmniboxViewViews"
      
  This.BackButton.Initialise "BackButton", This.MasterWindow, Descendants, By.NameIs, "Back", True

  This.RefreshButton.Initialise "RefreshButton", This.MasterWindow, Descendants, By.NameIs, "Reload"
  
  This.RootWebArea.Initialise "RootWebArea", This.MasterWindow, Descendants, By.AutomationId, "RootWebArea", True

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
  With This.AddressBox
    .Find 10
    GetCurrentURL = .Element.GetProperty(UIAProperties.Name)
    GetCurrentURL = Replace(GetCurrentURL, "Address and search bar, ", "")
  End With
End Function

Public Sub NavigateBack()
  WebBrowserCommon.Navigate Me, This.BackButton, This.AddressBox, This.RootWebArea
End Sub

Public Sub WaitForNewURL(TimeoutInSeconds As Integer)
  WebBrowserCommon.WaitForNewURL GetCurrentURL, This.AddressBox, This.RootWebArea, TimeoutInSeconds
End Sub

