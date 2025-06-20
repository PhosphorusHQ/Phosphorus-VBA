Attribute VB_Name = "ConfigurationTests"
'@Folder Configuration
' Purpose: Example usage of ConfigReader
Option Explicit

Public Sub TestConfigReader()

  GetConfigurationReader
  
  Dim server As String
  Dim port As Long
  Dim enabled As Boolean
  Dim LogLevel As String
  Dim maxRetries As Long
    
  ' Initialize with the path to the INI file
  Configuration.Initialize ThisWorkbook.Path & "\config.ini"
    
  If Configuration.IsInitialized Then
        
    server = Configuration.GetValue("Database", "Server", "default_server")
    port = Configuration.GetInteger("Database", "Port", 0)
    enabled = Configuration.GetBoolean("Database", "Enabled", False)
    LogLevel = Configuration.GetValue("Application", "LogLevel", "INFO")
    maxRetries = Configuration.GetInteger("Application", "MaxRetries", 1)
        
    Debug.Print "Database.Server: " & server
    Debug.Print "Database.Port: " & port
    Debug.Print "Database.Enabled: " & enabled
    Debug.Print "Application.LogLevel: " & LogLevel
    Debug.Print "Application.MaxRetries: " & maxRetries
  
    LogLevel = Configuration.GetValue("Logging", "LogLevel", "INFO")
    Debug.Print "Logging.LogLevel: " & LogLevel
  
  Else
    Debug.Print "Failed to initialize configuration"
  End If

End Sub
