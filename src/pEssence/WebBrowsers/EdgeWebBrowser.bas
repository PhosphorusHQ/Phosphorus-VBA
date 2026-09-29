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
  BackButton As pLocator
  RefreshButton As pLocator
  AddressAndSearchBar As pLocator
  RootWebArea As pLocator
End Type

Private This As BrowserAttributes
  
Private Sub Class_Initialize()
  GetAllLocators
End Sub

Private Sub Class_Terminate()
  Dim ProcessId As Long
  ProcessId = This.MasterWindow.Element.UIAElement.CurrentProcessId
  On Error Resume Next
  This.MasterWindow.Element.CloseWindow
  On Error GoTo 0
  'MSEdge does not always close the session properly, so new instancs remain hidde
  If ProcessId <> 0 Then
    Phosphorus.WindowsProcesses.KillProcessByID ProcessId
  End If
  Set This.MasterWindow = Nothing
  Set This.BackButton = Nothing
  Set This.RefreshButton = Nothing
  Set This.AddressAndSearchBar = Nothing
  Set This.RootWebArea = Nothing
End Sub

Private Sub GetAllLocators()
  DestroyLocators
End Sub

Private Sub DestroyLocators()
  Set This.MasterWindow = Factory.GetNewLocator
  Set This.BackButton = Factory.GetNewLocator
  Set This.RefreshButton = Factory.GetNewLocator
  Set This.AddressAndSearchBar = Factory.GetNewLocator
  Set This.RootWebArea = Factory.GetNewLocator
End Sub

Public Sub Start(WebAppName As String, URL As String, WebAppPageTitle As String, Optional BaseWaitTimeSeconds As Long = 10, Optional AbsoluteWaitTimeSeconds As Long)
  Toaster.Message "Starting " & WebAppName
  This.WebAppName = WebAppName
  This.URL = URL
  This.WebAppPageTitle = WebAppPageTitle
  LaunchCommandByProtocol This.WebAppName, "microsoft-edge:", This.URL, WindowShowStates.Maximized
'  LaunchExecutable Phosphorus.WindowsExecutables.MicrosoftEdge, URL, WindowShowStates.Maximized
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

  'Use: AscW & ChrW to determine embedded Unicode characters
  With This.MasterWindow
    .Initialise "MasterWindow", Nothing, Children, pConditions, "AND(NameLike, ControlType, ClassName, WindowInteractionState)"
    .Condition "NameLike", Name, IsLikeTheString, This.WebAppPageTitle & "* - *" & " - Microsoft" & ChrW(8203) & " Edge"
    .ControlType UIAControlTypeIDs.Window
    .ClassName "Chrome_WidgetWin_1"
    .WindowInteractionState ReadyForUserInteraction
  End With

  With This.BackButton
    .Initialise "BackButton", This.MasterWindow, Descendants, pConditions, "AND(ControlType, NameIs)": .ControlType UIAControlTypeIDs.Button: .NameIs "Back"
  End With

  With This.RefreshButton
    .Initialise "RefreshButton", This.MasterWindow, Descendants, pConditions, "AND(NameIs,ClassName)"
    .NameIs "Refresh"
    .ClassName "ReloadButton"
  End With

  With This.AddressAndSearchBar
    .Initialise "AddressAndSearchBar", This.MasterWindow, Descendants, pConditions, "AND(ControlType, NameIs, ClassName)": .ControlType UIAControlTypeIDs.Edit: .NameIs "Address and search bar": .ClassName "OmniboxViewViews"
    'Seems we need to force finding this here fro Edge
    .Find 10
  End With
  
  With This.RootWebArea
    .Initialise "RootWebArea", This.MasterWindow, Descendants, By.AutomationId, "RootWebArea", FindFirst:=True
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
  With This.RefreshButton
    If .Element.UIAElement Is Nothing Then
      'This seems to help!
      This.MasterWindow.Element.Click
      .Find 0
    End If
  End With
  Set GetRefreshButton = This.RefreshButton
End Function

Public Function GetCurrentURL() As String
  With This.AddressAndSearchBar
    .Find 10
    GetCurrentURL = .Element.GetValue()
  End With
End Function

Public Sub NavigateBack()
  WebBrowserCommon.Navigate Me, This.BackButton, This.AddressAndSearchBar, This.RootWebArea
End Sub

Public Sub WaitForNewURL(TimeoutInSeconds As Integer)
  WebBrowserCommon.WaitForNewURL GetCurrentURL, This.AddressAndSearchBar, This.RootWebArea, TimeoutInSeconds
End Sub

