Attribute VB_Name = "WindowsPowerShell"
'@Folder Windows
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

