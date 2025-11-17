Attribute VB_Name = "WindowsPowerShell"
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

' Global variable to control PowerShell server window visibility
Public PSSERVER_VISIBLE As Boolean

Private WindowsPowerShellPipeClient As Phosphorus.pWindowsPowerShellPipeClient

Private Sub GetWindowsPowerShellPipeClient()
  If WindowsPowerShellPipeClient Is Nothing Then
    ' Set PowerShell server visibility
    PSSERVER_VISIBLE = Phosphorus.Configuration.GetValue("PowerShellPipeClient", "Visible", False) ' Set to False to hide the PowerShell window
     Set WindowsPowerShellPipeClient = New Phosphorus.pWindowsPowerShellPipeClient
  End If
End Sub

Public Sub CloseWindowsPowerShellPipeClient()
  ' Clean up (automatically handled by Class_Terminate)
  Set WindowsPowerShellPipeClient = Nothing
End Sub

Public Function Execute(Command As String, CallingMethod As String) As String
  GetWindowsPowerShellPipeClient
  Execute = WindowsPowerShellPipeClient.Execute(Command, CallingMethod)
End Function

