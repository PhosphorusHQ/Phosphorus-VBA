Attribute VB_Name = "Log4PStatic"
'@Folder Log4P
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

'Global Logger object available for use
Public Logger As Log4P
'Public global variable for the log root folder
Public LogRootFolder As String
Public LogFileNameDynamicPart1 As String
Public LogFileNameDynamicPart2 As String

' Public Enum for log levels
Public Enum LogLevel
  Trace = 0
  Debugging = 1
  Info = 2
  Warning = 3
  Error = 4
  Fatal = 5
End Enum

Public Enum LogCategory
  Internal
  External
End Enum

Public Sub GetLogger()
  If Logger Is Nothing Then
    Set Logger = New Phosphorus.Log4P
  End If
End Sub

Public Sub CloseLogger()
  Set Logger = Nothing
End Sub

Public Function GetLevelName(level As Phosphorus.LogLevel) As String
  Select Case level
    Case LogLevel.Trace: GetLevelName = "Trace"
    Case LogLevel.Debugging: GetLevelName = "Debugging"
    Case LogLevel.Info: GetLevelName = "Info"
    Case LogLevel.Warning: GetLevelName = "Warning"
    Case LogLevel.Error: GetLevelName = "Error"
    Case LogLevel.Fatal: GetLevelName = "Fatal"
  End Select
End Function

