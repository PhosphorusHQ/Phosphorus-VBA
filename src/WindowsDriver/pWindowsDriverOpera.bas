VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pWindowsDriverOpera"
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
    InstanceType = Phosphorus.pInstanceType.Executable
  End If
  Dim CurrentPPath As String
  'Set default path
  CurrentPPath = "/Pane[@Name=""" & WebAppTitle & " - Opera""]"
  Select Case InstanceType
    Case Phosphorus.pInstanceType.Executable
      'Launch Opera via executable with no parameters other than the url, if any
      Phosphorus.WindowsProcesses.LaunchExecutable Phosphorus.WindowsExecutables.Opera, URL, Phosphorus.WindowShowStates.SW_SHOWMAXIMIZED
    Case Else
      Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverUnhandledAppConfiguration, "Opera, Instance Type: #" & InstanceType
  End Select
  ParentWindowsDriver.SetPageLoadedElement CurrentPPath, UIAutomationClient.WindowInteractionState.WindowInteractionState_ReadyForUserInteraction
End Sub

