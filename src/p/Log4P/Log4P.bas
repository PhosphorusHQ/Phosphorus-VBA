VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Log4P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Option Base 1

' Windows API Declarations
#If VBA7 Then
  Private Declare PtrSafe Function CreateFileW Lib "kernel32" ( _
    ByVal lpFileName As LongPtr, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As LongPtr, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As LongPtr) As LongPtr

  Private Declare PtrSafe Function WriteFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As LongPtr) As Long

  Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As LongPtr) As Long

  Private Declare PtrSafe Function FlushFileBuffers Lib "kernel32" ( _
    ByVal hFile As LongPtr) As Long

  Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" ( _
    ByRef lpPerformanceCount As LongLong) As Long

  Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" ( _
    ByRef lpFrequency As LongLong) As Long

  Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal codePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, _
    ByVal lpUsedDefaultChar As LongPtr) As Long
#Else
  Private Declare Function CreateFileW Lib "kernel32" ( _
    ByVal lpFileName As Long, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

  Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) As Long

  Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

  Private Declare Function FlushFileBuffers Lib "kernel32" ( _
    ByVal hFile As Long) As Long

  Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long

#End If

Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const OPEN_ALWAYS = 4
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const BUFFER_SIZE As Long = 8192 ' Bytes
Private Const DEFAULT_FLUSH_INTERVAL As Long = 100
Private Const CP_UTF8 As Long = 65001 ' Code page for UTF-8

Private logFilePath As String
#If VBA7 Then
    Private hFile As LongPtr
    Private lastLogTicks As LongLong
    Private perfFrequency As LongLong
    Private startTicks As LongLong
#Else
    Private hFile As Long
    Private lastLogTime As Double
    Private startTime As Double
#End If
Private currentLogLevel As LogLevel
Private previousLogLevel As LogLevel
Private writeCount As Long
Private flushInterval As Long
Private logBuffer As String
Private isBufferingEnabled As Boolean

Private Sub Class_Initialize()
  
  Dim rootFolder As String
  If VBA.Strings.Len(Log4PStatic.LogRootFolder) > 0 Then
    rootFolder = Log4PStatic.LogRootFolder
  Else
    rootFolder = ThisWorkbook.Path & "\Logs\"
  End If

  CreateFolder rootFolder
  'This must run AFTER the create folder method
  If VBA.Strings.Right(rootFolder, 1) <> "\" Then rootFolder = rootFolder & "\"
    
  currentLogLevel = LogLevel.Info
  previousLogLevel = LogLevel.Info
  writeCount = 0
  flushInterval = DEFAULT_FLUSH_INTERVAL
  logBuffer = ""
  isBufferingEnabled = True
    
  #If VBA7 Then
     If QueryPerformanceFrequency(perfFrequency) = 0 Then
       Err.Raise vbObjectError + 1, "Logger", "Failed to get performance frequency"
     End If
     QueryPerformanceCounter startTicks
     lastLogTicks = startTicks
  #Else
     startTime = Timer
     lastLogTime = startTime
  #End If
    
  'Build file path
  logFilePath = _
    rootFolder & _
    VBA.Strings.Format(Now, "yyyymmdd_hhmmss") & "_" & _
    Log4PStatic.LogFileNameDynamicPart1 & "_" & _
    Log4PStatic.LogFileNameDynamicPart2 & "_" & _
    VBA.Interaction.Environ$("COMPUTERNAME") & "_" & _
    VBA.Interaction.Environ$("USERNAME") & ".txt"
  If Dir(logFilePath) <> "" Then
    'Wait 1 second and create new filename
    Application.Wait (Now + TimeValue("0:00:01"))
    logFilePath = _
      rootFolder & _
      VBA.Strings.Format(Now, "yyyymmdd_hhmmss") & "_" & _
      Log4PStatic.LogFileNameDynamicPart1 & "_" & _
      Log4PStatic.LogFileNameDynamicPart2 & "_" & _
      VBA.Interaction.Environ$("COMPUTERNAME") & "_" & _
      VBA.Interaction.Environ$("USERNAME") & ".txt"
  End If
  ' Open file with CreateFileW for Unicode path
  hFile = CreateFileW(StrPtr(logFilePath), GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
  If hFile = -1 Then
    LogMessage "Failed to open log file: " & logFilePath, LogLevel.Error, Internal
    Err.Raise vbObjectError + 1, "Logger", "Failed to open log file"
  End If
    
  ' Write UTF-8 BOM
  Dim bom(0 To 2) As Byte
  bom(0) = &HEF
  bom(1) = &HBB
  bom(2) = &HBF
  Dim BytesWritten As Long
  If WriteFile(hFile, bom(0), 3, BytesWritten, 0) = 0 Then
    CloseHandle hFile
    LogMessage "Failed to write UTF-8 BOM", LogLevel.Error, Internal
    Err.Raise vbObjectError + 1, "Logger", "Failed to initialize log file"
  End If
    
  Info "Logger initialized with file: " & logFilePath, Internal

  'Always log file opened & closed messages - these can be filtered out in the log reader
  Dim strHeaderLine As String
  strHeaderLine = "Timestamp" & VBA.Constants.vbTab
  strHeaderLine = strHeaderLine & "LogEntryType" & VBA.Constants.vbTab & "LogLevel" & VBA.Constants.vbTab & "Delta" & VBA.Constants.vbTab & "Message"
    
  LogMessage strHeaderLine, LogLevel.Fatal, Internal
  LogMessage "Log file opened", LogLevel.Fatal, Internal, True
    
End Sub

Private Sub CreateFolder(folderPath As String)
    
  Dim parts() As String
  Dim i As Long
  Dim currentPath As String
    
  If VBA.Strings.Right(folderPath, 1) = "\" Then folderPath = VBA.Strings.Left(folderPath, VBA.Strings.Len(folderPath) - 1)
  parts = VBA.Strings.Split(folderPath, "\")
  currentPath = parts(0) & "\"
  For i = 1 To UBound(parts)
    currentPath = currentPath & parts(i) & "\"
    If Dir(currentPath, vbDirectory) = "" Then
      On Error Resume Next
      MkDir currentPath
      If Err.Number <> 0 Then
        LogMessage "Failed to create folder: " & currentPath & " (Error " & Err.Number & ")", LogLevel.Error, Internal
        Err.Raise vbObjectError + 1, "Logger", "Failed to create folder"
      End If
      On Error GoTo 0
    End If
  Next i

End Sub

Private Sub Class_Terminate()
  'Always log file closed message - these can be filtered out in the log reader
  LogMessage "Log file closed", LogLevel.Fatal, Internal
  If VBA.Strings.Len(logBuffer) > 0 Then WriteBuffer
  If hFile <> 0 Then
    FlushFileBuffers hFile
    CloseHandle hFile
    hFile = 0
  End If
End Sub

Public Property Let level(Value As Long)
  currentLogLevel = Value
End Property

Public Property Get level() As Long
  level = currentLogLevel
End Property

Public Sub SetTempLevel(newLevel As LogLevel)
  Logger.Info "Log level set to " & Log4PStatic.GetLevelName(newLevel), Internal
  previousLogLevel = currentLogLevel
  currentLogLevel = newLevel
End Sub

Public Sub RestoreLevel()
  Logger.Info "Log level reset to " & Log4PStatic.GetLevelName(previousLogLevel), Internal
  currentLogLevel = previousLogLevel
End Sub

Private Sub WriteBuffer()
  If hFile <> 0 And VBA.Strings.Len(logBuffer) > 0 Then
    Dim bytes() As Byte
    bytes = StringToUtf8Bytes(logBuffer)
    Dim BytesWritten As Long
    If WriteFile(hFile, bytes(0), UBound(bytes) + 1, BytesWritten, 0) = 0 Then
      LogMessage "Failed to write buffer to log file", LogLevel.Error, Internal
      Err.Raise vbObjectError + 1, "Logger", "Failed to write buffer to log file"
    End If
    logBuffer = ""
  End If
End Sub

Private Function StringToUtf8Bytes(ByVal str As String) As Byte()

  If VBA.Strings.Len(str) = 0 Then
    StringToUtf8Bytes = vbNullString
    Exit Function
  End If
    
  ' Get required buffer size
  Dim byteCount As Long
  byteCount = WideCharToMultiByte(CP_UTF8, 0, StrPtr(str), VBA.Strings.Len(str), 0, 0, 0, 0)
    
  ' Allocate byte array
  Dim bytes() As Byte
  ReDim bytes(0 To byteCount - 1)
    
  ' Convert to UTF-8
  byteCount = WideCharToMultiByte(CP_UTF8, 0, StrPtr(str), VBA.Strings.Len(str), VarPtr(bytes(0)), byteCount, 0, 0)
    
  StringToUtf8Bytes = bytes

End Function

Private Sub LogMessage(message As String, level As LogLevel, Cat As LogCategory, Optional forceFlush As Boolean = False)
    
  If level < currentLogLevel Then Exit Sub
    
  Dim timestamp As String
  Dim levelStr As String
  Dim logEntry As String
  Dim timeDiff As Double
  Dim microSeconds As Long
  
  levelStr = Log4PStatic.GetLevelName(level)
  
  'Do not change the log entry if we are writing the header row to the log file
  If VBA.Strings.InStr(1, message, "Timestamp") = 1 Then
    logEntry = message & vbCrLf
  Else
  
    #If VBA7 Then
       Dim currentTicks As LongLong
       QueryPerformanceCounter currentTicks
       timeDiff = VBA.Conversion.CDbl(currentTicks - lastLogTicks) / VBA.Conversion.CDbl(perfFrequency) * 1000
       ' Calculate microseconds since start for timestamp
       microSeconds = VBA.Conversion.CLng((VBA.Conversion.CDbl(currentTicks - startTicks) / VBA.Conversion.CDbl(perfFrequency)) * 1000000) Mod 1000000
       timestamp = VBA.Strings.Format(Now, "yyyy-mm-dd hh:nn:ss") & "." & VBA.Strings.Right("000000" & microSeconds, 6)
    #Else
       Dim currentTime As Double
       currentTime = Timer
       If currentTime < lastLogTime Then
         timeDiff = ((currentTime + 86400) - lastLogTime) * 1000
       Else
         timeDiff = (currentTime - lastLogTime) * 1000
       End If
       ' Use Timer for milliseconds (less precise)
       microSeconds = VBA.Conversion.CLng((currentTime - startTime) * 1000) Mod 1000
       timestamp = VBA.Strings.Format(Now, "yyyy-mm-dd hh:nn:ss") & "." & VBA.Strings.Right("000" & microSeconds, 3)
    #End If
      
    'Build up the log entry
    logEntry = timestamp & VBA.Constants.vbTab
        
    'Add the log level string
    logEntry = logEntry & VBA.Constants.vbTab & levelStr & VBA.Constants.vbTab
       
    'The accurary of the delta varies
    #If VBA7 Then
       logEntry = logEntry & VBA.Strings.ChrW(916) & VBA.Strings.Format(timeDiff, "0.000") & "ms" & VBA.Constants.vbTab
    #Else
       logEntry = logEntry & VBA.Strings.ChrW(916) & VBA.Strings.Format(timeDiff, "0.00") & "ms | "
    #End If
   
    ' Add the message
    logEntry = logEntry & message & vbCrLf
    
  End If
    
  If isBufferingEnabled Then
    
    logBuffer = logBuffer & logEntry
      
    ' Estimate UTF-8 byte size (conservative: up to 4 bytes per char)
    If VBA.Strings.Len(logBuffer) * 4 >= BUFFER_SIZE Then
      WriteBuffer
    End If
        
    If forceFlush Then
      WriteBuffer
      FlushFileBuffers hFile
    End If
        
    writeCount = writeCount + 1
    If writeCount Mod flushInterval = 0 Then
      WriteBuffer
      FlushFileBuffers hFile
    End If
    
  Else
    
    Dim bytes() As Byte
    bytes = StringToUtf8Bytes(logEntry)
    Dim BytesWritten As Long
    If hFile <> 0 Then
      If WriteFile(hFile, bytes(0), UBound(bytes) + 1, BytesWritten, 0) = 0 Then
        Err.Raise vbObjectError + 1, "Logger", "Failed to write to log file"
      End If
    End If
    writeCount = writeCount + 1
    If forceFlush Then FlushFileBuffers hFile
  
  End If
    
  #If VBA7 Then
     QueryPerformanceCounter lastLogTicks
  #Else
    lastLogTime = Timer
  #End If

End Sub

' Public logging methods for each log level

Public Sub Trace(message As String, Optional Cat As LogCategory = External, Optional forceFlush As Boolean = False)
  LogMessage message, LogLevel.Trace, Cat, forceFlush
End Sub

Public Sub Debugging(message As String, Optional Cat As LogCategory = External, Optional forceFlush As Boolean = False)
  LogMessage message, LogLevel.Debugging, Cat, forceFlush
End Sub

Public Sub Info(message As String, Optional Cat As LogCategory = External, Optional forceFlush As Boolean = False)
  LogMessage message, LogLevel.Info, Cat, forceFlush
End Sub

Public Sub Warning(message As String, Optional Cat As LogCategory = External, Optional forceFlush As Boolean = False)
  LogMessage message, LogLevel.Warning, Cat, forceFlush
End Sub

Public Sub Error(message As String, Optional Cat As LogCategory = External, Optional forceFlush As Boolean = False)
  LogMessage message, LogLevel.Error, Cat, forceFlush
End Sub

Public Sub Fatal(message As String, Optional Cat As LogCategory = External, Optional forceFlush As Boolean = False)
  LogMessage message, LogLevel.Fatal, Cat, forceFlush
End Sub

Public Sub Flush()
  If VBA.Strings.Len(logBuffer) > 0 Then WriteBuffer
  If hFile <> 0 Then FlushFileBuffers hFile
End Sub

Public Sub DisableBuffering()
  If VBA.Strings.Len(logBuffer) > 0 Then WriteBuffer
  isBufferingEnabled = False
  flushInterval = 0
End Sub

Public Sub RestoreBuffering()
  isBufferingEnabled = True
  flushInterval = DEFAULT_FLUSH_INTERVAL
End Sub

Public Function GetFilePath() As String
  GetFilePath = logFilePath
End Function
