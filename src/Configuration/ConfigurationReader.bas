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

Private This As Settings

' Initialize the ConfigReader with the path to the INI file
Public Sub Initialize(ByVal FilePath As String)
  Set This.configData = New Scripting.dictionary
  This.configFilePath = FilePath
  This.IsInitialized = LoadConfigFile()
End Sub

' Loads the INI file into the configData dictionary
Private Function LoadConfigFile() As Boolean
    
  On Error GoTo ErrorHandler
    
  Dim fileNum As Integer
  Dim line As String
  Dim currentSection As String
  Dim Key As String, Value As String
  Dim pos As Integer
    
  ' Check if file exists
  If Not FileExists(This.configFilePath) Then
    MsgBox "Configuration file not found: " & This.configFilePath, vbCritical
    LoadConfigFile = False
    Exit Function
  End If
    
  fileNum = FreeFile
  Open This.configFilePath For Input As #fileNum
    
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
      Key = VBA.Strings.Trim(VBA.Strings.Left(line, pos - 1))
      Value = VBA.Strings.Trim(VBA.Strings.Mid(line, pos + 1))
      If VBA.Strings.Len(currentSection) > 0 Then
        This.configData.Add currentSection & "." & Key, Value
      Else
        This.configData.Add Key, Value
      End If
    End If

Continue:
  Wend
    
  Close #fileNum
  LoadConfigFile = True
  Exit Function

ErrorHandler:
  If Not fileNum = 0 Then Close #fileNum
  MsgBox "Error loading configuration file: " & Err.Description, vbCritical
  LoadConfigFile = False
  
End Function

' Get a configuration value as a string
Public Function GetValue(ByVal section As String, ByVal Key As String, Optional ByVal defaultValue As String = "") As String
  Dim lookupKey As String
  If IsInitialized Then
    lookupKey = IIf(Len(section) > 0, section & "." & Key, Key)
    If This.configData.exists(lookupKey) Then
      GetValue = This.configData(lookupKey)
    Else
      GetValue = defaultValue
    End If
  Else
    GetValue = defaultValue
  End If
End Function

' Get a configuration value as an integer
Public Function GetInteger(ByVal section As String, ByVal Key As String, Optional ByVal defaultValue As Long = 0) As Long
  Dim Value As String
  Value = GetValue(section, Key, CStr(defaultValue))
  If IsNumeric(Value) Then
    GetInteger = VBA.Conversion.CLng(Value)
  Else
    GetInteger = defaultValue
  End If
End Function

' Get a configuration value as a boolean
Public Function GetBoolean(ByVal section As String, ByVal Key As String, Optional ByVal defaultValue As Boolean = False) As Boolean
  Dim Value As String
  Value = VBA.Strings.UCase(GetValue(section, Key, CStr(defaultValue)))
  If Value = "TRUE" Or Value = "1" Or Value = "YES" Then
    GetBoolean = True
  ElseIf Value = "FALSE" Or Value = "0" Or Value = "NO" Then
    GetBoolean = False
  Else
    GetBoolean = defaultValue
  End If
End Function

' Check if a file exists
Private Function FileExists(ByVal FilePath As String) As Boolean
  FileExists = Dir(FilePath) <> ""
End Function

' Property to check if the configuration is initialized
Public Property Get IsInitialized() As Boolean
  IsInitialized = This.IsInitialized
End Property
