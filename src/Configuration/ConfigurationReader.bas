VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigurationReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Configuration
' Purpose: Reads runtime configuration settings from an INI file
Option Explicit

Private Type Settings
  configFilePath As String
  configData As Scripting.dictionary ' Requires Microsoft Scripting Runtime
  IsInitialized As Boolean
End Type

Private this As Settings

' Initialize the ConfigReader with the path to the INI file
Public Sub Initialize(ByVal filePath As String)
  Set this.configData = New Scripting.dictionary
  this.configFilePath = filePath
  this.IsInitialized = LoadConfigFile()
End Sub

' Loads the INI file into the configData dictionary
Private Function LoadConfigFile() As Boolean
    
  On Error GoTo ErrorHandler
    
  Dim fileNum As Integer
  Dim line As String
  Dim currentSection As String
  Dim key As String, value As String
  Dim pos As Integer
    
  ' Check if file exists
  If Not FileExists(this.configFilePath) Then
    MsgBox "Configuration file not found: " & this.configFilePath, vbCritical
    LoadConfigFile = False
    Exit Function
  End If
    
  fileNum = FreeFile
  Open this.configFilePath For Input As #fileNum
    
  currentSection = ""
    
  While Not EOF(fileNum)
    Line Input #fileNum, line
    line = VBA.Strings.Trim(line)
        
    ' Skip empty lines or comments
    If VBA.Strings.Len(line) = 0 Or VBA.Strings.Left(line, 1) = ";" Or VBA.Strings.Left(line, 1) = "#" Then
      GoTo Continue
    End If
        
    ' Check for section header
    If VBA.Strings.Left(line, 1) = "[" And VBA.Strings.Right(line, 1) = "]" Then
      currentSection = VBA.Strings.Mid(line, 2, VBA.Strings.Len(line) - 2)
      GoTo Continue
    End If
        
    ' Parse key-value pair
    pos = VBA.Strings.InStr(line, "=")
    If pos > 0 Then
      key = VBA.Strings.Trim(VBA.Strings.Left(line, pos - 1))
      value = VBA.Strings.Trim(VBA.Strings.Mid(line, pos + 1))
      If VBA.Strings.Len(currentSection) > 0 Then
        this.configData.Add currentSection & "." & key, value
      Else
        this.configData.Add key, value
      End If
    End If

Continue:
  Wend
    
  Close #fileNum
  LoadConfigFile = True
  Exit Function

ErrorHandler:
  If Not fileNum = 0 Then Close #fileNum
  MsgBox "Error loading configuration file: " & err.Description, vbCritical
  LoadConfigFile = False
  
End Function

' Get a configuration value as a string
Public Function GetValue(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As String = "") As String
  Dim lookupKey As String
  If IsInitialized Then
    lookupKey = IIf(Len(section) > 0, section & "." & key, key)
    If this.configData.exists(lookupKey) Then
      GetValue = this.configData(lookupKey)
    Else
      GetValue = defaultValue
    End If
  Else
    GetValue = defaultValue
  End If
End Function

' Get a configuration value as an integer
Public Function GetInteger(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As Long = 0) As Long
  Dim value As String
  value = GetValue(section, key, CStr(defaultValue))
  If IsNumeric(value) Then
    GetInteger = VBA.Conversion.CLng(value)
  Else
    GetInteger = defaultValue
  End If
End Function

' Get a configuration value as a boolean
Public Function GetBoolean(ByVal section As String, ByVal key As String, Optional ByVal defaultValue As Boolean = False) As Boolean
  Dim value As String
  value = VBA.Strings.UCase(GetValue(section, key, CStr(defaultValue)))
  If value = "TRUE" Or value = "1" Or value = "YES" Then
    GetBoolean = True
  ElseIf value = "FALSE" Or value = "0" Or value = "NO" Then
    GetBoolean = False
  Else
    GetBoolean = defaultValue
  End If
End Function

' Check if a file exists
Private Function FileExists(ByVal filePath As String) As Boolean
  FileExists = Dir(filePath) <> ""
End Function

' Property to check if the configuration is initialized
Public Property Get IsInitialized() As Boolean
  IsInitialized = this.IsInitialized
End Property
