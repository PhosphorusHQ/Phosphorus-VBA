VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pWindowsDriverChrome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder WindowsDriver
Option Explicit

'https://peter.sh/experiments/chromium-command-line-switches/

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
  CurrentPPath = "/Window[@Name=""" & WebAppTitle & " - Google Chrome""]"
  Select Case InstanceType
    Case Phosphorus.pInstanceType.Executable
      'Launch Chrome via executable with no parameters other than the url, if any
      Phosphorus.WindowsProcesses.LaunchExecutable Phosphorus.WindowsExecutables.Chrome, "--force-renderer-accessibility " & URL, Phosphorus.WindowShowStates.SW_SHOWMAXIMIZED
    Case Phosphorus.pInstanceType.NewProfile
      TempDirectory = ParentWindowsDriver.CreateTempDirectory
      Phosphorus.WindowsProcesses.LaunchExecutable Phosphorus.WindowsExecutables.Chrome, "--force-renderer-accessibility" & " --user-data-dir=""" & TempDirectory & """ " & URL, Phosphorus.WindowShowStates.SW_SHOWMAXIMIZED
      'Override the current path for the new profile sign-in screen
      CurrentPPath = "/Window[@Name=""Google Chrome""]"
    Case Phosphorus.pInstanceType.GuestModeNoSignIn
      TempDirectory = ParentWindowsDriver.CreateTempDirectory
      ' --bwsi Indicates that the browser is in "browse without sign-in" (Guest session) mode. Should completely disable extensions, sync and bookmarks.
      Phosphorus.WindowsProcesses.LaunchExecutable Phosphorus.WindowsExecutables.Chrome, "--force-renderer-accessibility" & " --bwsi" & " --user-data-dir=""" & TempDirectory & """ " & URL, Phosphorus.WindowShowStates.SW_SHOWMAXIMIZED
    Case Else
      Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverUnhandledAppConfiguration, "Chrome, Instance Type: #" & InstanceType
  End Select
  ParentWindowsDriver.SetPageLoadedElement CurrentPPath, UIAutomationClient.WindowInteractionState.WindowInteractionState_ReadyForUserInteraction
End Sub

