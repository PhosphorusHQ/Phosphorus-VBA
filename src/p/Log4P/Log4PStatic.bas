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
  INTERNAL_TRACE = -6
  INTERNAL_DEBUG = -5
  INTERNAL_INFO = -4
  INTERNAL_WARNING = -3
  INTERNAL_ERROR = -2
  INTERNAL_FATAL = -1
  EXTERNAL_TRACE = 0
  EXTERNAL_DEBUG = 1
  EXTERNAL_INFO = 2
  EXTERNAL_WARNING = 3
  EXTERNAL_ERROR = 4
  EXTERNAL_FATAL = 5
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
    Case INTERNAL_TRACE: GetLevelName = "INTERNAL_TRACE"
    Case INTERNAL_DEBUG: GetLevelName = "INTERNAL_DEBUG"
    Case INTERNAL_INFO: GetLevelName = "INTERNAL_INFO"
    Case INTERNAL_WARNING: GetLevelName = "INTERNAL_WARNING"
    Case INTERNAL_ERROR: GetLevelName = "INTERNAL_ERROR"
    Case INTERNAL_FATAL: GetLevelName = "INTERNAL_FATAL"
    Case EXTERNAL_TRACE: GetLevelName = "EXTERNAL_TRACE"
    Case EXTERNAL_DEBUG: GetLevelName = "EXTERNAL_DEBUG"
    Case EXTERNAL_INFO: GetLevelName = "EXTERNAL_INFO"
    Case EXTERNAL_WARNING: GetLevelName = "EXTERNAL_WARNING"
    Case EXTERNAL_ERROR: GetLevelName = "EXTERNAL_ERROR"
    Case EXTERNAL_FATAL: GetLevelName = "EXTERNAL_FATAL"
  End Select
End Function

