Attribute VB_Name = "ConfigurationTests"
'@Folder Configuration
' Purpose: Example usage of ConfigReader
Option Explicit

Public Sub TestConfigReader()

  Phosphorus.Configuration.GetGlobalConfigurationReader
  
  Dim server As String
  Dim port As Long
  Dim enabled As Boolean
  Dim LogLevel As String
  Dim maxRetries As Long
    
  ' Initialize with the path to the INI file
  Phosphorus.ConfigReader.Initialize ThisWorkbook.path & "\config.ini"
    
  If Phosphorus.ConfigReader.IsInitialized Then
        
    server = Phosphorus.ConfigReader.GetValue("Database", "Server", "default_server")
    port = Phosphorus.ConfigReader.GetInteger("Database", "Port", 0)
    enabled = Phosphorus.ConfigReader.GetBoolean("Database", "Enabled", False)
    LogLevel = Phosphorus.ConfigReader.GetValue("Application", "LogLevel", "INFO")
    maxRetries = Phosphorus.ConfigReader.GetInteger("Application", "MaxRetries", 1)
        
    Debug.Print "Database.Server: " & server
    Debug.Print "Database.Port: " & port
    Debug.Print "Database.Enabled: " & enabled
    Debug.Print "Application.LogLevel: " & LogLevel
    Debug.Print "Application.MaxRetries: " & maxRetries
  
    LogLevel = Phosphorus.ConfigReader.GetValue("Logging", "LogLevel", "INFO")
    Debug.Print "Logging.LogLevel: " & LogLevel
  
  Else
    Debug.Print "Failed to initialize configuration"
  End If

End Sub
