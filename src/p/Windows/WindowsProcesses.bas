Attribute VB_Name = "WindowsProcesses"
'@Folder Windows
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public Const DEFAULT_CAPPED_SLEEP_TIME_IN_SECONDS = 30

'https://x.com/i/grok?conversation=1970433932106539134
'ShellExecute Parameters:hwnd: Handle to the parent window (use 0 for no parent).
'  lpOperation: Use "open" to launch the app.
'  lpFile: The App ID with the shell:AppsFolder\ prefix.
'  lpParameters: Parameters to pass to the app (usually vbNullString for UWP apps).
'  lpDirectory: Default directory (use vbNullString for default).
' nShowCmd: Window state (e.g., 1 for SW_SHOWNORMAL to show the window normally).
#If VBA7 Then
  Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As LongPtr
#Else
  private Declare  Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
#End If

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow
Public Enum WindowShowStates
  Hide = 0 'SW_HIDE
  Normal = 1 'SW_SHOWNORMAL
   Minimized = 2 'SW_SHOWMINIMIZED
   Maximized = 3 'SW_SHOWMAXIMIZED
   NoActivate = 4 'SW_SHOWNOACTIVATE
End Enum

#If VBA7 Then
  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
#End If

Private Sub Tests_LaunchAppByAUMID()
  'Not Ok
'  LaunchAppByAUMID Phosphorus.WindowsWindowsApps.NotAValidApp, Phosphorus.WindowShowStates.SW_SHOWNORMAL
'  LaunchAppByAUMID Phosphorus.WindowsWindowsApps.MSTeams, Phosphorus.WindowShowStates.SW_SHOWNORMAL
  'Ok
  'No UI?
'  LaunchAppByAUMID Phosphorus.WindowsWindowsApps.MicrosoftEdgeDevToolsClient, Phosphorus.WindowShowStates.SW_SHOWNORMAL
End Sub

Public Sub LaunchAppByAUMID(myWindowsApp As Phosphorus.WindowsApp, Optional URL As String, Optional ByVal ShowCmd As Phosphorus.WindowShowStates)
  myWindowsApp.AppID = Phosphorus.WindowsWindowsApps.GetAppID(myWindowsApp.OfficialName)
  'TODO: Do all windows apps open in a normal screen state regardless of which value we pass
  RunShellExecuteToStartNewProcess myWindowsApp.FriendlyName, "open", "shell:appsFolder\" & myWindowsApp.AppID, URL, VBA.Constants.vbNullString, ShowCmd
End Sub

Public Sub LaunchCommandByProtocol(ByVal ApplicationName, Protocol As String, URL As String, ByVal ShowCmd As Phosphorus.WindowShowStates)
  RunShellExecuteToStartNewProcess ApplicationName, "open", Protocol & URL, VBA.Constants.vbNullString, VBA.Constants.vbNullString, ShowCmd
End Sub

Public Sub Test_LaunchExecutable()
'  LaunchExecutable Phosphorus.WindowsExecutables.WindowsExplorer, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.SW_SHOWMINIMIZED
'  LaunchExecutable Phosphorus.WindowsExecutables.MicrosoftWord, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.SW_SHOWMINIMIZED
  'https://peter.sh/experiments/chromium-command-line-switches/
'  LaunchExecutable Phosphorus.WindowsExecutables.PowerShell, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.SW_SHOWMINIMIZED
'  LaunchExecutable Phosphorus.WindowsExecutables.Opera, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.SW_SHOWMINIMIZED
'  LaunchExecutable Phosphorus.WindowsExecutables.WindowsNotepad, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.SW_SHOWMINIMIZED
'  LaunchExecutable Phosphorus.WindowsExecutables.WindowsPCHealthCheck, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.SW_SHOWMINIMIZED
'  LaunchExecutable Phosphorus.WindowsExecutables.AdobeAcrobat, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.SW_SHOWMINIMIZED
'  LaunchExecutable Phosphorus.WindowsExecutables.WindowsMediaPlayer, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.SW_SHOWMINIMIZED
'  LaunchExecutable Phosphorus.WindowsExecutables.NotepadPlusPlus, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.SW_SHOWMINIMIZED

' WINDOWS SYSTEM TOOLS
'  LaunchExecutable Phosphorus.WindowsExecutables.WindowsTaskManager, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.SW_SHOWMINIMIZED
'  LaunchExecutable Phosphorus.WindowsExecutables.WindowsCharacterMap, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.SW_SHOWMINIMIZED
'  LaunchExecutable Phosphorus.WindowsExecutables.WindowsDiskCleanUp, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.Minimized
'    LaunchExecutable Phosphorus.WindowsExecutables.Firefox, " -P Phosphorus -url https://www.example.com/", Phosphorus.WindowShowStates.Maximized
  LaunchExecutable Phosphorus.WindowsExecutables.Firefox, " -url https://www.example.com/", Phosphorus.WindowShowStates.Maximized
    
End Sub

'Public Sub OpenExecuteable(ByVal ShowCmd As Phosphorus.WindowShowStates)
Public Sub LaunchExecutable(Exe As Phosphorus.Executable, ByVal Parameters As String, ByVal ShowCmd As Phosphorus.WindowShowStates)
  If Exe.FullPath = "" Then
    Exe.FullPath = GetAppPathFromRegistry(Exe.ExeFile)
  End If
  RunShellExecuteToStartNewProcess Exe.Name, VBA.Constants.vbNullString, Exe.FullPath, Parameters, vbNullString, ShowCmd
End Sub

'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & Executable & "\Path
Private Function GetAppPathFromRegistry(ByVal Exe As String) As String
  Dim AppPath As String
  'Check for the 'Path' subkey
  AppPath = Phosphorus.WindowsRegistry.ReadRegistry("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & Exe & "\Path")
  If AppPath = VBA.Constants.vbNullString Then
    'Get the default sub key
    AppPath = Phosphorus.WindowsRegistry.ReadRegistry("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & Exe & "\")
  End If
  'Replace %ProgramFiles(x86)% environment variable
  Dim strProgramFilesx86EnvironmentVariableSetting As String
  strProgramFilesx86EnvironmentVariableSetting = Environ("ProgramFiles(x86)")
   AppPath = VBA.Strings.Replace(AppPath, "%ProgramFiles(x86)%", strProgramFilesx86EnvironmentVariableSetting)
  'Remove any leading and trailing quotes
  If VBA.Strings.Left(AppPath, 1) = VBA.Strings.Chr(34) Then
    AppPath = VBA.Strings.Mid(AppPath, 2, VBA.Strings.Len(AppPath) - 1)
  End If
  If VBA.Strings.Right(AppPath, 1) = VBA.Strings.Chr(34) Then
    AppPath = VBA.Strings.Left(AppPath, VBA.Strings.Len(AppPath) - 1)
  End If
  If VBA.Strings.Right(AppPath, VBA.Strings.Len(Exe) + 1) <> ("\" & Exe) Then
    If VBA.Strings.Right(AppPath, 1) <> "\" Then
      AppPath = AppPath & "\"
    End If
    AppPath = AppPath & Exe
  End If
  GetAppPathFromRegistry = AppPath
End Function

Private Sub Test_LaunchShellFolder()
'  LaunchShellFolder Phosphorus.WindowsWindowsShellFolders.TestMisconfiguredWindowShellFolder
'  LaunchShellFolder Phosphorus.WindowsWindowsShellFolders.Desktop
'  LaunchShellFolder Phosphorus.WindowsWindowsShellFolders.MyComputer
'  LaunchShellFolder Phosphorus.WindowsWindowsShellFolders.RecycleBin
'  LaunchShellFolder Phosphorus.WindowsWindowsShellFolders.UserProfile
  LaunchShellFolder Phosphorus.WindowsWindowsShellFolders.CommonAdministrativeTools
End Sub

Public Sub LaunchShellFolder(myShellFolder As Phosphorus.ShellFolder)
  Dim Protocol As String
  Dim URI As String
  If myShellFolder.Path <> "" Then
    URI = myShellFolder.Path
    Protocol = "shell:"
  ElseIf myShellFolder.CLSID <> "" Then
    URI = myShellFolder.CLSID
    Protocol = "shell:::"
  Else
    Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverMisconfiguredWindowShellFolder, myShellFolder.Name
  End If
  RunShellExecuteToStartNewProcess myShellFolder.Name, "open", Protocol & URI, VBA.Constants.vbNullString, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.Maximized
End Sub

Public Sub Test_LaunchWindowsServiceConsole()
'  LaunchWindowsServiceConsole Phosphorus.WindowsMMCSnapIns.ComputerManagement
'  LaunchWindowsServiceConsole Phosphorus.WindowsMMCSnapIns.EventViewer
'  LaunchWindowsServiceConsole Phosphorus.WindowsMMCSnapIns.PerformanceMonitor
'  LaunchWindowsServiceConsole Phosphorus.WindowsMMCSnapIns.WindowsManagementInstrumentation
  LaunchWindowsServiceConsole Phosphorus.WindowsMMCSnapIns.WindowsServices
End Sub

Public Sub LaunchWindowsServiceConsole(myMMC As Phosphorus.WindowsMMCSnapIn)
  RunShellExecuteToStartNewProcess myMMC.Name, "open", myMMC.FileName, VBA.Constants.vbNullString, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.Maximized
End Sub

Public Sub Test_LaunchWindowsSettings()
'  LaunchWindowsSetting Phosphorus.WindowsWindowsSettings.Accounts_AccessWorkOrSchool
'  LaunchWindowsSetting Phosphorus.WindowsWindowsSettings.Accounts_YourInfo
'  LaunchWindowsSetting Phosphorus.WindowsWindowsSettings.Apps_AppsAndFeatures
'Apps pass parameters?  LaunchWindowsSetting Phosphorus.WindowsWindowsSettings.Apps_AppsFeatures, "DuckDuckGo"
'  LaunchWindowsSetting Phosphorus.WindowsWindowsSettings.ControlCenter_ControlCenter
'  LaunchWindowsSetting Phosphorus.WindowsWindowsSettings.Devices_Autoplay
'  LaunchWindowsSetting Phosphorus.WindowsWindowsSettings.Devices_Bluetooth
'  LaunchWindowsSetting Phosphorus.WindowsWindowsSettings.Devices_YourPhone
  LaunchWindowsSetting Phosphorus.WindowsWindowsSettings.NetworkAndInternet_NetworkAndInternet
End Sub

Public Sub LaunchWindowsSetting(myWindowsSetting As Phosphorus.WindowsSetting, Optional parameter As String)
  If myWindowsSetting.SettingsPage = "App features" Then
    RunShellExecuteToStartNewProcess myWindowsSetting.SettingsPage, "open", "ms-settings:" & myWindowsSetting.URISuffix & "?" & parameter, VBA.Constants.vbNullString, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.Maximized
  Else
    RunShellExecuteToStartNewProcess myWindowsSetting.SettingsPage, "open", "ms-settings:" & myWindowsSetting.URISuffix, VBA.Constants.vbNullString, VBA.Constants.vbNullString, Phosphorus.WindowShowStates.Maximized
  End If
End Sub

Public Sub RunShellExecuteToStartNewProcess( _
  ByVal ApplicationName, _
  ByVal Operation As String, _
  ByVal CommandString As String, _
  ByVal Parameters As String, _
  ByVal Directory As String, _
  ByVal ShowCmd As Long)
  
#If VBA7 Then
  Dim Result As LongPtr
#Else
  Dim Result As Long
#End If

  'Always pass 0 to hwnd for a new parent window
  Result = ShellExecute(0, Operation, CommandString, Parameters, Directory, ShowCmd)
  
  If Result <= 32 Then
    Dim FullCommand As String
    FullCommand = Operation & " " & CommandString
    If Parameters <> VBA.Constants.vbNullString Then
      FullCommand = FullCommand & Parameters
    End If
    If Directory <> VBA.Constants.vbNullString Then
      FullCommand = FullCommand & Directory
    End If
    pExceptions.Raise Exceptions.WindowsDriverFailedToStartProgram, ApplicationName, FullCommand
'    Err.Raise 1, "Failed to open URL. Error code: " & result, vbCritical
  End If
  
End Sub

Public Sub Snooze(ByVal SleepTimeInMilliseconds As Long)
  Dim CappedSleepTimeoutInMilliseconds As Long
  CappedSleepTimeoutInMilliseconds = Phosphorus.Configuration.GetValue("Processes", "CappedSleepTime", DEFAULT_CAPPED_SLEEP_TIME_IN_SECONDS) * 1000
  If SleepTimeInMilliseconds <= CappedSleepTimeoutInMilliseconds Then
    'Sleep safely!
    Sleep SleepTimeInMilliseconds
  Else
    'Raise Exception ...
    Phosphorus.pExceptions.Raise _
      Phosphorus.Exceptions.WindowsDriverCappedSleepTimeoutInSecondsExceeded, _
      SleepTimeInMilliseconds, _
      CappedSleepTimeoutInMilliseconds
  End If
End Sub

Public Sub KillProcessByID(pid As Long)
  'Use taskkill to force-terminate a process
  VBA.Interaction.Shell "taskkill /PID " & pid & " /F", vbHide
End Sub

Public Function GetProcessIdByCommandLine(ByVal commandLineSnippet As String) As Long
    
  Dim wmi As Object
  Dim processes As Object
  Dim process As Object
  Dim pid As Long
    
  On Error GoTo ErrorHandler
    
  ' Create WMI object
  Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  Set processes = wmi.ExecQuery("SELECT ProcessId, CommandLine FROM Win32_Process WHERE CommandLine LIKE '%" & VBA.Strings.Replace(commandLineSnippet, "\", "\\") & "%'")
    
  ' Find the process with matching command line
  For Each process In processes
    If VBA.Strings.InStr(1, process.commandLine, commandLineSnippet, vbTextCompare) > 0 Then
      pid = process.ProcessId
      Exit For
    End If
  Next
    
  GetProcessIdByCommandLine = pid
  Exit Function

ErrorHandler:
    GetProcessIdByCommandLine = 0
End Function




