VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChromiumWebBrowser"
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
  BackButton As pLocator
  RefreshButton As pLocator
  LastTab As pLocator
  CloseOtherTabs As pLocator
  AddressAndSearchBar As pLocator
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
  Set This.BackButton = Factory.GetNewLocator
  Set This.RefreshButton = Factory.GetNewLocator
  Set This.AddressAndSearchBar = Factory.GetNewLocator
  Set This.LastTab = Factory.GetNewLocator
  Set This.CloseOtherTabs = Factory.GetNewLocator
  Set This.RootWebArea = Factory.GetNewLocator
End Sub

Private Sub DestroyLocators()
  Set This.MasterWindow = Nothing
  Set This.BackButton = Nothing
  Set This.RefreshButton = Nothing
  Set This.AddressAndSearchBar = Nothing
  Set This.LastTab = Nothing
  Set This.CloseOtherTabs = Nothing
  Set This.RootWebArea = Nothing
End Sub

'Download latest stable Chromium binaries (64-bit and 32-bit)
'XXXhttps://chromium.woolyss.com/

'https://www.chromium.org/getting-involved/download-chromium/
'https://commondatastorage.googleapis.com/chromium-browser-snapshots/index.html?prefix=win_rel/

'How to install Chromium for all users on Windows
'https://martinrotter.github.io/it-programming/2016/07/17/install-chromium-system-wide-windows/

Public Sub Start(WebAppName As String, URL As String, WebAppPageTitle As String, Optional BaseWaitTimeSeconds As Long = 10, Optional AbsoluteWaitTimeSeconds As Long)
  Toaster.Message "Starting " & WebAppName
  This.WebAppName = WebAppName
  This.URL = URL
  This.WebAppPageTitle = WebAppPageTitle
  'Chromium accessibility issues - see:
  'https://www.google.com/search?q=vba+ui+automation+chromium+doesn%27t+render+we+h1+elements&sca_esv=ad07602a4d316d67&sxsrf=APpeQnu7MTYaW6ZyRo-CaTaFuqQZTPbSUw%3A1792768719151&udm=50&fbs=ABfTbFVyMZGZf1hfvX9uKjN_-G8cxpBkeIeqYwoCbfNVc4vKE5XdvSlI_LORjvMqi_KlDwmEfG0_Tf_E2J0y8OiIQPeyixqgWEePseYtZq-i0Axub4ueg1hsfAR5TyoWmmO7pCsltueWXXPzzcq3vLuRYnmKaKuNEPRS1Ek2pM-ix_Igh3BZltasrW71LZpm-btxyae7p4aGOC2-XGHGnGwSozV3UgOovQ&vsint=&aep=1&ntc=1&cs=0&sa=X&ved=2ahUKEwjU_6z1t9CXAxXAZ0EAHWNSKDIQ2J8OegQIExAD&biw=1229&bih=556&dpr=1.56&mstk=AUtExfBDlQBOlfcs9JvR0rqo8RHTcY9lOUl1pzbEmcpPcwYLAjMY2k8nUdsERZfphYXxexLUFF2DT1M7eeTfCmdCMG9ZrLC2kr-0bFuq7uJTlpR_kv3CmooNYIhc58cxnOguz319xq0UbClYPKupek---iCfy2v-IQsjif8_LBdJNEeKYdRiToEluXldAK4Q2kJxqX60hHDfuC7MNYKo5TEE5WOkFaosG5kTFHHzE4WTv37MSQ2HQFiPzmS5mufeWZyl5jjc25wy1Y08dd0w2UZgsIiMx8xZR5RCwDjf05o1VEUMvwiriIAevQjM4BTpgGSKjNoPCBjNUxE_Kg&csuir=1&mtid=03rbauO_K7m1hbIPseeG6Ag
  'chrome://accessibility/
  LaunchExecutable Phosphorus.WindowsExecutables.Chromium, " --force-renderer-accessibility=complete --enable-features=UiaProvider " & URL, WindowShowStates.Maximized
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
    .NameIs This.WebAppPageTitle & " - Chromium"
    .ControlType UIAControlTypeIDs.Pane
    .ClassName "Chrome_WidgetWin_1"
    .WindowInteractionState ReadyForUserInteraction
  End With
    
  With This.BackButton
    .Initialise "BackButton", This.MasterWindow, Descendants, pConditions, "AND(ControlType, NameIs)": .ControlType UIAControlTypeIDs.Button: .NameIs "Back"
  End With
  
  This.RefreshButton.Initialise "RefreshButton", This.MasterWindow, Descendants, By.NameIs, "Reload"

  With This.AddressAndSearchBar
    .Initialise "AddressAndSearchBar", This.MasterWindow, Descendants, By.NameIs, "Address and search bar"
  End With


  With This.LastTab
    .Initialise "LastTab", This.MasterWindow, Descendants, By.ControlType, UIAControlTypeIDs.TabItem, FindFirst:=True
    .PositionInMatchingSet -1
  End With

  'Right click & close all other tabs
  This.LastTab.Element.RightClick
  
  With This.CloseOtherTabs
    .Initialise "CloseOtherTabs", This.MasterWindow, Descendants, By.pConditions, "AND(ControlType, NameIs)", FindFirst:=True
    .ControlType MenuItem: .NameIs "Close other tabs"
    .Element.ClickIfEnabled This.MasterWindow.Element
  End With

  This.RootWebArea.Initialise "RootWebArea", This.MasterWindow, Descendants, By.ControlType, UIAControlTypeIDs.Document, FindFirst:=True

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
  With This.AddressAndSearchBar
    GetCurrentURL = .Element.Value()
  End With
End Function

Public Sub NavigateBack()
  WebBrowserCommon.Navigate Me, This.BackButton, This.AddressAndSearchBar, This.RootWebArea
End Sub

Public Sub WaitForNewURL(TimeoutInSeconds As Integer)
  WebBrowserCommon.WaitForNewURL GetCurrentURL, This.AddressAndSearchBar, This.RootWebArea, TimeoutInSeconds
End Sub

Public Sub AcknowledgeChangeYourPasswordAlert()

  Application.Wait (Now + TimeValue("0:00:01"))
  
'  Dim RootView As pLocator
'  Set RootView = Factory.GetNewLocator
'  With RootView
'    .Initialise "RootView", This.BrowserRootView, Children, By.ClassName, "RootView"
'    .WaitForElementExists 10:  .Find 10
'  End With
   
  Dim PasswordAlertOkButton As pLocator
  Set PasswordAlertOkButton = Factory.GetNewLocator
  Dim Continue As Boolean
  With PasswordAlertOkButton
'    .Initialise "RootView", This.MasterWindow, Descendants, By.AriaRole, AriaRoles.Button
    .Initialise "RootView", This.MasterWindow, Descendants, By.pConditions, "AND(AriaRole, NameIs)": .PositionInMatchingSet 1
    .AriaRole AriaRoles.Button
    .NameIs "Close"
    If .ElementExists(3) Then
      .Find 10
      Continue = True
      While Continue
        .Element.Click
        Application.Wait (Now + TimeValue("0:00:01"))
        Continue = False
        On Error Resume Next
        Continue = .ElementExists(0)
        On Error GoTo 0
      Wend
      This.RootWebArea.Find 10, FindElementAgain:=True
    End If
  End With
  
'  Set RootView = Nothing
  Set PasswordAlertOkButton = Nothing
  
End Sub

