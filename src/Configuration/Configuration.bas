Attribute VB_Name = "Configuration"
'@Folder Configuration
Option Explicit

Public ConfigReader As Phosphorus.ConfigurationReader

Public Sub GetGlobalConfigurationReader()
  If ConfigReader Is Nothing Then
    Set ConfigReader = New Phosphorus.ConfigurationReader
    ConfigReader.Initialize ThisWorkbook.Path & "\config.ini"
  End If
End Sub

Public Sub GetLocalConfigurationReader(FilePath As String)
  Set ConfigReader = New Phosphorus.ConfigurationReader
  ConfigReader.Initialize FilePath
End Sub

Public Function GetValue(ByVal section As String, ByVal Key As String, Optional ByVal defaultValue As String = "") As String
  If ConfigReader Is Nothing Then
    GetGlobalConfigurationReader
  End If
  If ConfigReader.IsInitialized Then
    GetValue = ConfigReader.GetValue(section, Key, defaultValue)
  End If
End Function
