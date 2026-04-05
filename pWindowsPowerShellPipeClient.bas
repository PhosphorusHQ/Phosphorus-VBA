VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pWindowsPowerShellPipeClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
 
'https://learn.microsoft.com/en-us/windows/win32/ipc/named-pipes

#If VBA7 Then
   
  Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As LongPtr, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As LongPtr) As LongPtr
 
  Private Declare PtrSafe Function WriteFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As LongPtr) As Long

  Private Declare PtrSafe Function ReadFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As LongPtr) As Long

  Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

  Private Declare PtrSafe Function WaitNamedPipe Lib "kernel32" Alias "WaitNamedPipeA" ( _
    ByVal lpNamedPipeName As String, _
    ByVal nTimeOut As Long) As Long

  Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

  Private Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As LongPtr

  Private Declare PtrSafe Function TerminateProcess Lib "kernel32" ( _
    ByVal hProcess As LongPtr, _
    ByVal uExitCode As Long) As Long

  Private Const INVALID_HANDLE_VALUE As LongPtr = -1

#Else
  
  Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

  Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToWrite As Long, _
          lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) As Long

  Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
          lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Long) As Long

  Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

  Private Declare Function WaitNamedPipe Lib "kernel32" Alias "WaitNamedPipeA" ( _
    ByVal lpNamedPipeName As String, _
    ByVal nTimeOut As Long) As Long

  Private Declare Function GetLastError Lib "kernel32" () As Long

  Private Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

  Private Declare Function TerminateProcess Lib "kernel32" ( _
    ByVal hProcess As Long, _
    ByVal uExitCode As Long) As Long

  Private Const INVALID_HANDLE_VALUE As Long = -1

#End If

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const PROCESS_TERMINATE As Long = &H1
Private Const NMPWAIT_WAIT_FOREVER As Long = -1

Private PipeHandle As LongPtr
Private ProcessId As Long
Private ProcessHandle As LongPtr
Private myWshShell As wshShell
'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/legacy/bb776890(v=vs.85)
'Microsoft Shell Controls and Automation (Shell32)
Private ScriptPath As String
Private Const PIPE_NAME As String = "\\.\pipe\VBAPowerShellPipe"
Private Const TempDirectory As String = "C:\Temp\"
Private Const ReadyFlagFile As String = "C:\Temp\PSReady.txt"

Private Sub Class_Initialize()
    
  Dim Result As Long
  Dim commandLine As String
  
  'Ensure we clean up after errors
  On Error GoTo ErrorHandler
    
  ' Set script path to workbook directory
  If ThisWorkbook.Path = "" Then
    'Raise Exception
    Phosphorus.pExceptions.Raise Phosphorus.Exceptions.PowerShellPipeClientWorkbookMustBeSaved
  End If
  ScriptPath = ThisWorkbook.Path & "\PSPipeServer.ps1"
    
  ' Write PowerShell script if it doesn't exist
  WritePowerShellScript
  
  'Create the temp directory where the ready flag will be created
  If VBA.FileSystem.Dir(TempDirectory, VBA.VbFileAttribute.vbDirectory) = "" Then
    VBA.FileSystem.MkDir TempDirectory
  End If
  'Delete the flag in case it already exists
  If VBA.FileSystem.Dir(ReadyFlagFile) <> "" Then
    VBA.FileSystem.Kill ReadyFlagFile
  End If
  
  ' Start the PowerShell server
  Set myWshShell = New wshShell 'CreateObject("WScript.Shell")
  
  Dim windowStyle As Integer
  windowStyle = VBA.Interaction.IIf(Phosphorus.WindowsPowerShell.PSSERVER_VISIBLE, 1, 0) ' 1 = visible, 0 = hidden
  commandLine = "powershell -NoExit -ExecutionPolicy Bypass -File """ & ScriptPath & """"
  'Use .Run rather than .Exec as .Exec does not let us set the window style
  myWshShell.Run commandLine, windowStyle, False
  
  ' Get the process ID using WMI
  ' .Run does not give us the PID lile .Exec does
  ProcessId = GetProcessIdByCommandLine(ScriptPath)
  If ProcessId = 0 Then
    'Raise Exception
    Phosphorus.pExceptions.Raise Phosphorus.Exceptions.PowerShellPipeClientFailedToRetrievePowerShellProcessID
  End If
  
  'Wait for the named pipe to be available - check for ReadFlagFile exists
  While VBA.FileSystem.Dir(ReadyFlagFile) = ""
    Phosphorus.WindowsProcesses.Snooze 100
  Wend
  'Delete the flag ready for next launch
  VBA.FileSystem.Kill ReadyFlagFile
  
  Dim TimoutSeconds As Integer
  Dim TimeOutMilliseconds As Long
  TimoutSeconds = 30
  TimeOutMilliseconds = TimoutSeconds * 1000
  Result = WaitNamedPipe(PIPE_NAME, TimeOutMilliseconds)
  If Result = 0 Then
    'Raise Exception
    Phosphorus.pExceptions.Raise Phosphorus.Exceptions.PowerShellPipeClientPipeNotAvailableWithinTimeout, TimoutSeconds
  End If
    
  ' Connect to the named pipe
  PipeHandle = CreateFile(PIPE_NAME, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
  If PipeHandle = INVALID_HANDLE_VALUE Then
    'Raise Exception
    Phosphorus.pExceptions.Raise Phosphorus.Exceptions.PowerShellPipeClientFailedToConnectToNamedPipe, GetLastError
  End If
    
  Exit Sub

ErrorHandler:
  ' Clean up in case of error during initialization
  Dim ErrNumber As Long
  Dim ErrDescription As String
  ErrNumber = Err.Number
  ErrDescription = Err.Description
  CleanUp
  Err.Raise ErrNumber, , ErrDescription
End Sub

Private Sub Class_Terminate()
  CleanUp
End Sub

Private Sub CleanUp(Optional Force As Boolean)

  Dim BytesWritten As Long
  
  If Not Force Then
    ''Force' bypasses the clean exit
    ' Send exit command to PowerShell server if pipe is open
    If PipeHandle <> INVALID_HANDLE_VALUE And PipeHandle <> 0 Then
      WriteFile PipeHandle, "exit" & vbCrLf, Len("exit" & vbCrLf), BytesWritten, 0
      CloseHandle PipeHandle
      PipeHandle = 0
    End If
  End If
  
  ' Terminate the PowerShell process if still running
  If ProcessId <> 0 Then
    ProcessHandle = OpenProcess(PROCESS_TERMINATE, 0, ProcessId)
    If ProcessHandle <> 0 Then
      TerminateProcess ProcessHandle, 0
      CloseHandle ProcessHandle
      ProcessHandle = 0
    End If
    ProcessId = 0
  End If
    
  ' Release WScript.Shell
  Set myWshShell = Nothing

End Sub

Public Function Execute(ByVal Command As String, CallingMethod As String) As String
    
  Dim BytesWritten As Long
  Dim Result As Long
  Dim LastError As Long
  
  ' Validate pipe handle
  If PipeHandle = INVALID_HANDLE_VALUE Or PipeHandle = 0 Then
    'Raise an exception
    Phosphorus.pExceptions.Raise Phosphorus.Exceptions.PowerShellPipeClientPipeHandleIsInvalid
  End If
    
  ' Send the command
  Command = Command & vbCrLf
  Result = WriteFile(PipeHandle, Command, Len(Command), BytesWritten, 0)
  If Result = 0 Then
    LastError = GetLastError
    'We need to Force a cleanup first here as the unforced clean will hang otherwise
    CleanUp Force:=True
    'Raise an exception
    Phosphorus.pExceptions.Raise Phosphorus.Exceptions.PowerShellPipeClientFailedToWriteToPipe, LastError
  End If
    
  ' Read the response
  Execute = ReadPipeResponse(PipeHandle)
  If Execute = "" Then
    'Raise an exception
    Phosphorus.pExceptions.Raise Phosphorus.Exceptions.PowerShellPipeClientFailedToReadResponse, LastError
  End If

  'Check for error messages
  If VBA.Strings.InStr(1, Execute, "Error: ") Then
    'Raise an exception
    Phosphorus.pExceptions.Raise Phosphorus.Exceptions.PowerShellPipeClientPowerShellError, CallingMethod, Command
  End If

End Function

Private Function ReadPipeResponse(ByVal PipeHandle As LongPtr) As String
    
  Dim buffer As String
  Dim bytesRead As Long
  Dim Result As Long
  Dim totalOutput As String
  Dim bufferSize As Long
    
  ' Set initial buffer size
  bufferSize = 4096
  buffer = String(bufferSize, Chr(0))
  totalOutput = ""
    
  ' Read until no more data or error
  Do
    Result = ReadFile(PipeHandle, buffer, bufferSize, bytesRead, 0)
    If Result = 0 Or bytesRead = 0 Then
      If GetLastError = 109 Then ' ERROR_BROKEN_PIPE
        Exit Do
      End If
      If totalOutput = "" Then
        totalOutput = "Error: " & GetLastError
        Exit Do
      End If
    End If
    totalOutput = totalOutput & Left(buffer, bytesRead)
  Loop While bytesRead = bufferSize ' Continue if buffer was filled
    
  ReadPipeResponse = totalOutput
  
End Function

Private Sub WritePowerShellScript()

  Dim fso As Object
  Dim file As Object
  Dim scriptContent As String
    
  On Error GoTo ErrorHandler
    
  ' Check if the script file already exists
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(ScriptPath) Then
    Exit Sub
  End If
        
  ' PowerShell script content
  scriptContent = _
    "# Define the named pipe name" & vbCrLf & _
    "$pipeName = ""VbaPowerShellPipe""" & vbCrLf & _
    "$pipe = New-Object System.IO.Pipes.NamedPipeServerStream($pipeName, [System.IO.Pipes.PipeDirection]::InOut)" & vbCrLf & _
    "" & vbCrLf & _
    "try {" & vbCrLf & _
    "    Write-Host ""Waiting for client connection...""" & vbCrLf & _
    "    # ----------  CREATE READY FLAG  ----------" & vbCrLf & _
    "    New-Item -Path """ & ReadyFlagFile & """ -ItemType File -Force | Out-Null" & vbCrLf & _
    "    $pipe.WaitForConnection()" & vbCrLf & _
    "    Write-Host ""Client connected.""" & vbCrLf & _
    "" & vbCrLf & _
    "    $reader = New-Object System.IO.StreamReader($pipe)" & vbCrLf & _
    "    $writer = New-Object System.IO.StreamWriter($pipe)" & vbCrLf & _
    "    $writer.AutoFlush = $true" & vbCrLf & _
    "" & vbCrLf & _
    "    # Keep the server running to process commands" & vbCrLf & _
    "    while ($true) {" & vbCrLf & _
    "        # Read command from VBA" & vbCrLf & _
    "        $command = $reader.ReadLine()" & vbCrLf & _
    "        if ($command -eq ""exit"") { break } # Exit condition" & vbCrLf & _
    "" & vbCrLf & _
    "        # Echo the received command to the console" & vbCrLf & _
    "        Write-Host ""Received command: $command""" & vbCrLf

  scriptContent = scriptContent & vbCrLf & _
    "        try {" & vbCrLf & _
    "            # Execute the command and capture output" & vbCrLf & _
    "            $result = Invoke-Expression $command -ErrorAction Stop | Out-String" & vbCrLf & _
    "            $writer.WriteLine($result)" & vbCrLf & _
    "" & vbCrLf & _
    "        # Echo the result to the console" & vbCrLf & _
    "        Write-Host ""Result: $result""" & vbCrLf & _
    "        }" & vbCrLf & _
    "        catch {" & vbCrLf & _
    "            # Send error back to client" & vbCrLf & _
    "            $writer.WriteLine(""Error: $_"")" & vbCrLf & _
    "        }" & vbCrLf & _
    "    }" & vbCrLf & _
    "}" & vbCrLf & _
    "finally {" & vbCrLf & _
    "    # Clean up" & vbCrLf & _
    "    $reader.Dispose()" & vbCrLf & _
    "    $writer.Dispose()" & vbCrLf & _
    "    $pipe.Dispose()" & vbCrLf & _
    "}"

  ' Write the script to the file
  Set file = fso.CreateTextFile(ScriptPath, True)
  file.Write scriptContent
  file.Close
  Exit Sub

ErrorHandler:
  'Raise an exception
  Phosphorus.pExceptions.Raise Phosphorus.Exceptions.PowerShellPipeClientFailedFailedToWritePowerShellScript, Err.Number, Err.Description
End Sub

