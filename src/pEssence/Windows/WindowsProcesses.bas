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

' --------------------------------------------------------------
' ShellExecuteW – Unicode version (recommended)
' Compatible with 32-bit & 64-bit Office (VBA7+ = Office 2010 and newer)
' --------------------------------------------------------------

#If VBA7 Then   ' Office 2010+

  #If Win64 Then
    ' 64-bit Office
    Private Declare PtrSafe Function ShellExecuteW Lib "shell32.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal lpOperation As LongPtr, _
      ByVal lpFile As LongPtr, _
      ByVal lpParameters As LongPtr, _
      ByVal lpDirectory As LongPtr, _
      ByVal nShowCmd As Long) As LongPtr
  #Else
    ' 32-bit Office (even on 64-bit Windows)
    Private Declare PtrSafe Function ShellExecuteW Lib "shell32.dll" ( _
      ByVal hWnd As LongPtr, _
      ByVal lpOperation As LongPtr, _
      ByVal lpFile As LongPtr, _
      ByVal lpParameters As LongPtr, _
      ByVal lpDirectory As LongPtr, _
      ByVal nShowCmd As Long) As LongPtr
  #End If

#Else   ' VBA6 – Office 2007 or older (very rare in 2026)

  Private Declare Function ShellExecuteW Lib "shell32.dll" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

#End If

#If VBA7 Then
  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
#End If

Public Enum WindowStyle
  Hide = 0
  Normal = 1
  Minimized = 2
  Maximized = 3
  NoActivate = 4
End Enum

Public Sub RunShellExecuteToStartNewProcess( _
  ByVal ApplicationName, _
  ByVal Operation As String, _
  ByVal CommandString As String, _
  ByVal Parameters As String, _
  ByVal Directory As String, _
  ByVal ShowCmd As Long)
  
  Toaster.Message Operation & " " & ApplicationName
  
#If VBA7 Then
  Dim Result As LongPtr
#Else
  Dim Result As Long
#End If

  'Always pass 0 to hwnd for a new parent window
#If VBA7 Then   ' Office 2010+
  Result = ShellExecuteW(0&, StrPtr(Operation), StrPtr(CommandString), StrPtr(Parameters), StrPtr(Directory), ShowCmd)
#Else   ' VBA6 – Office 2007 or older (very rare in 2026)
  Result = ShellExecuteW(0, Operation, CommandString, Parameters, Directory, ShowCmd)
#End If
  
  If Result <= 32 Then
    Dim FullCommand As String
    FullCommand = Operation & " " & CommandString
    If Parameters <> VBA.Constants.vbNullString Then
      FullCommand = FullCommand & Parameters
    End If
    If Directory <> VBA.Constants.vbNullString Then
      FullCommand = FullCommand & Directory
    End If
    MsgBox "Failed to open " & ApplicationName & ". Error code: " & Result, vbExclamation
  End If
  
End Sub

Public Sub Snooze(ByVal SleepTimeInMilliseconds As Long)
  Dim CappedSleepTimeoutInMilliseconds As Long
  CappedSleepTimeoutInMilliseconds = 5000
  If SleepTimeInMilliseconds <= SleepTimeInMilliseconds Then
    Sleep SleepTimeInMilliseconds
  Else
     MsgBox "Snooze exceeds max snooze!"
  End If
End Sub

Public Sub LaunchCommandByProtocol(ByVal ApplicationName, Protocol As String, URL As String, ByVal ShowCmd As WindowStyle)
  RunShellExecuteToStartNewProcess ApplicationName, "open", Protocol & URL, VBA.Constants.vbNullString, VBA.Constants.vbNullString, ShowCmd
End Sub

