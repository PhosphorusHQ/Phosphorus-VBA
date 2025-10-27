Attribute VB_Name = "WindowsRegistry"
'@Folder Windows
Option Explicit

Private Sub Tests()
  Dim RegSetting As String
  RegSetting = Now()
  Debug.Print RegSetting
  Dim RegKey As String
  RegKey = "HKEY_CURRENT_USER\Software\Phosphorus\TestWriteSetting"
  Phosphorus.WindowsRegistry.WriteRegistry RegKey, RegSetting, "REG_SZ"
  Phosphorus.Configuration.GetGlobalConfigurationReader
  Phosphorus.WindowsProcesses.Snooze 2000
  Debug.Print Phosphorus.WindowsRegistry.ReadRegistry(RegKey)
  DeleteRegistry RegKey
End Sub

Public Function ReadRegistry(KeyPath As String) As Variant
  On Error Resume Next
  Dim WShell As wshShell
  Set WShell = New wshShell
  ReadRegistry = WShell.RegRead(KeyPath)
'  If Err.Number <> 0 Then
'    ReadRegistry = Null
'    MsgBox "Error reading registry: " & Err.Description, vbExclamation
'  End If
  Set WShell = Nothing
End Function

Public Sub WriteRegistry(KeyPath As String, Value As Variant, ValueType As String)
  On Error Resume Next
  Dim WShell As wshShell
  Set WShell = New wshShell
  WShell.RegWrite KeyPath, Value, ValueType
  If Err.Number <> 0 Then
    MsgBox "Error writing to registry: " & Err.Description, vbExclamation
  End If
  Set WShell = Nothing
End Sub

Sub DeleteRegistry(KeyPath As String)
  On Error Resume Next
  Dim WShell As wshShell
  Set WShell = New wshShell
  WShell.RegDelete KeyPath
  If Err.Number <> 0 Then
    MsgBox "Error deleting from registry: " & Err.Description, vbExclamation
  End If
  Set WShell = Nothing
End Sub

