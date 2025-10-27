VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pWindowsDriverDuckDuckGo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder WindowsDriver
Option Explicit

Implements Phosphorus.IWindowsDriverWebBrowser
Dim TempDirectory As String

Private ParentWindowsDriver As pWindowsDriver

Public Function IWindowsDriverWebBrowser_GetParentWindowsDriver() As pWindowsDriver
  Set IWindowsDriverWebBrowser_GetParentWindowsDriver = ParentWindowsDriver
End Function

Public Sub IWindowsDriverWebBrowser_LaunchApp(ByRef ParentWindowsDriver As pWindowsDriver, WebAppName As String, WebAppTitle As String, Optional URL As String)
  Set ParentWindowsDriver = ParentWindowsDriver
  Dim InstanceType As Phosphorus.pInstanceType
  InstanceType = ParentWindowsDriver.GetWindowsDriverWebBrowserType.InstanceType
  If InstanceType = 0 Then
    InstanceType = Phosphorus.pInstanceType.ApplicationUserModelID
  End If
  Select Case InstanceType
    Case Phosphorus.pInstanceType.ApplicationUserModelID
      Phosphorus.WindowsProcesses.LaunchAppByAUMID Phosphorus.WindowsWindowsApps.DuckDuckGo, URL, Phosphorus.WindowShowStates.SW_SHOWNORMAL
    Case Else
      Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverUnhandledAppConfiguration, "Duck Duck Go, Instance Type: #" & InstanceType
  End Select
  Dim CurrentPPath As String
'  CurrentPPath = "/Window[And(@Name=""NoAutomationPeer"",@AutomationId=""BrowserWindow"")]//Pane[@Name=""" & WebAppTitle & """]"
  CurrentPPath = "/Window[And(@Name=""NoAutomationPeer"",@AutomationId=""BrowserWindow"")]"
  ParentWindowsDriver.SetPageLoadedElement CurrentPPath, UIAutomationClient.WindowInteractionState.WindowInteractionState_ReadyForUserInteraction
End Sub


