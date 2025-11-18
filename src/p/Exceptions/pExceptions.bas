Attribute VB_Name = "pExceptions"
'@Folder Exceptions
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public Enum Exceptions
'Catch All Errors 0+
  UnhandledException = 1
  UnexpectedError = 2
'PowerShellPipeClientErrors '10,000+
  PowerShellPipeClientWorkbookMustBeSaved = 10001
  PowerShellPipeClientFailedFailedToWritePowerShellScript = 10002
  PowerShellPipeClientFailedToRetrievePowerShellProcessID = 10003
  PowerShellPipeClientPipeNotAvailableWithinTimeout = 10004
  PowerShellPipeClientFailedToConnectToNamedPipe = 10005
  PowerShellPipeClientPipeHandleIsInvalid = 10006
  PowerShellPipeClientFailedToWriteToPipe = 10007
  PowerShellPipeClientFailedToReadResponse = 10008
  PowerShellPipeClientPowerShellError = 100099
  
'Windows Driver Launch Errors 20,000+
  WindowsDriverUnhandledDriverType = 20001
  WindowsDriverUnhandledWebBrowserType = 20002
  WindowsDriverFailedToStartProgram = 20003
  WindowsDriverWindowsAppNotFound = 20004
  WindowsDriverUnhandledAppConfiguration = 20005
  WindowsDriverMisconfiguredWindowShellFolder = 20006
  WindowsDriverInvalidHTTPStatus = 20007
  WindowsDriverInvalidURL = 20008
  WindowsDriverUIElementNotFoundBeforeTimeout = 20009
  WindowsDriverUIElementCondtionOrStateNotMetBeforeTimeout = 20010
  WindowsDriverCappedSleepTimeoutInSecondsExceeded = 20011
  WindowsDriverUndefinedWindowInteractionState = 200012
  WindowsDriverPatternNotHandled = 200013
  
End Enum

Public Sub Raise(exception As Phosphorus.Exceptions, Optional parameter1 As Variant, Optional parameter2 As Variant, Optional parameter3 As Variant, Optional parameter4 As Variant)

  'Get string representation of the parameters passed
  parameter1 = VBA.Conversion.CStr(parameter1)
  parameter2 = VBA.Conversion.CStr(parameter2)
  parameter3 = VBA.Conversion.CStr(parameter3)
  parameter4 = VBA.Conversion.CStr(parameter4)
  
  parameter1 = TrimErrorMessageParameter(parameter1)
  parameter2 = TrimErrorMessageParameter(parameter2)
  parameter3 = TrimErrorMessageParameter(parameter3)
  parameter4 = TrimErrorMessageParameter(parameter4)
  
  Dim ExceptionName As String
  Dim ErrorNumber As Long
  Dim ErrorDescription As String
  ErrorNumber = exception
  
  'Handle each type of error
  Select Case exception
    
    Case Exceptions.UnexpectedError
      ExceptionName = "UnexpectedError"
      ErrorDescription = "Unexpected Error with error code #" & parameter1 & " - " & parameter2 & " (in method " & parameter3 & ")"
   
    'PowerShell Pipe Client Errors
   
    Case Exceptions.PowerShellPipeClientWorkbookMustBeSaved
      ExceptionName = "PowerShellPipeClientWorkbookMustBeSaved"
      ErrorDescription = "Power Shell Pipe Client: workbook must be saved to determine script path"
  
    Case Exceptions.PowerShellPipeClientFailedFailedToWritePowerShellScript
      ExceptionName = "PowerShellPipeClientFailedFailedToWritePowerShellScript"
      ErrorDescription = "Power Shell Pipe Client: Failed to write PowerShell script - Error #" & parameter1 & ": " & parameter2
  
    Case Exceptions.PowerShellPipeClientFailedToRetrievePowerShellProcessID
      ExceptionName = "PowerShellPipeClientFailedToRetrievePowerShellProcessID"
      ErrorDescription = "Power Shell Pipe Client: failed to retrieve PowerShell process ID"
      
    Case Exceptions.PowerShellPipeClientPipeNotAvailableWithinTimeout
      ExceptionName = "PowerShellPipeClientPipeNotAvailableWithinTimeout"
      ErrorDescription = "Power Shell Pipe Client: wipe not available within " & parameter1 & " seconds"
     
    Case Exceptions.PowerShellPipeClientFailedToConnectToNamedPipe
      ExceptionName = "PowerShellPipeClientFailedToConnectToNamedPipe"
      ErrorDescription = "Power Shell Pipe Client: failed to connect to named pipe - Error: " & parameter1
   
    Case Exceptions.PowerShellPipeClientPipeHandleIsInvalid
      ExceptionName = "PowerShellPipeClientPipeHandleIsInvalid"
      ErrorDescription = "Power Shell Pipe Client: Pipe handle is invalid. Ensure the class is properly initialized."
  
    Case Exceptions.PowerShellPipeClientFailedToWriteToPipe
      ExceptionName = "PowerShellPipeClientFailedToWriteToPipe"
      ErrorDescription = "Power Shell Pipe Client: failed to write to pipe, received error: " & parameter1
      
    Case Exceptions.PowerShellPipeClientFailedToReadResponse
      ExceptionName = "PowerShellPipeClientFailedToReadResponse"
      ErrorDescription = "Power Shell Pipe Client: failed to read response, received error: " & parameter1

    Case Exceptions.PowerShellPipeClientPowerShellError
      ExceptionName = "PowerShellPipeClientPowerShellError"
      ErrorDescription = "Power Shell Pipe Client: PowerShell error: " & parameter1 & " in method " & parameter2 & " for command: " & parameter3
   
    'Windows Driver Launch Errors
   
    Case Exceptions.WindowsDriverUnhandledDriverType
      ExceptionName = "UnhandledDriverType"
      ErrorDescription = "Unhandled Driver Type: " & parameter1
   
    Case Exceptions.WindowsDriverUnhandledWebBrowserType
      ExceptionName = "UnhandledWebBrowserType"
      ErrorDescription = "Unhandled Web Browser Type: " & parameter1
   
    Case Exceptions.WindowsDriverFailedToStartProgram
      ExceptionName = "FailedToStartProgram"
      ErrorDescription = "Failed to start program: " & parameter1 & " with command: " & parameter2
      
    Case Exceptions.WindowsDriverWindowsAppNotFound
      ExceptionName = "WindowsAppNotFound"
      ErrorDescription = "Windows App: '" & parameter1 & "' not found!"
  
    Case Exceptions.WindowsDriverUnhandledAppConfiguration
      ExceptionName = "UnhandledAppConfiguration"
      ErrorDescription = "Unhandled App Configuration: " & parameter1
  
    Case Exceptions.WindowsDriverMisconfiguredWindowShellFolder
      ExceptionName = "MisconfiguredWindowShellFolder"
      ErrorDescription = "Windows shell folder " & parameter1 & "' must have a Path or a CLSID specified"
   
    Case Exceptions.WindowsDriverInvalidHTTPStatus
      ExceptionName = "InvalidHTTPStatus"
      ErrorDescription = "HTTP Status Code #" & parameter2 & " returned for URL: " & parameter1

    Case Exceptions.WindowsDriverInvalidURL
      ExceptionName = "InvalidURL"
      ErrorDescription = "Invalid URL: " & parameter1 & ", message '" & parameter2 & "'"

    Case Exceptions.WindowsDriverUIElementNotFoundBeforeTimeout
      ExceptionName = "UIElementNotFoundBeforeTimeout"
      ErrorDescription = "Element not found before " & parameter1 & " seconds timeout for PPath: " & parameter2
   
    Case Exceptions.WindowsDriverUIElementCondtionOrStateNotMetBeforeTimeout
      ExceptionName = "UIElementCondtionOrStateNotMetBeforeTimeout"
      ErrorDescription = "UI Element '" & parameter1 & "' condition or state '" & parameter2 & "' not met - got '" & parameter3 & "', timout: " & parameter4 & " seconds"
    
    Case Exceptions.WindowsDriverCappedSleepTimeoutInSecondsExceeded
      ExceptionName = "CappedSleepTimeoutInSecondsExceeded"
      ErrorDescription = "A requested sleep time of " & parameter1 & " seconds is greater than the capped timeout of " & parameter2 & " seconds"
  
   Case Exceptions.WindowsDriverUndefinedWindowInteractionState
     ExceptionName = "UIElementNotFoundBeforeTimeout"
     ErrorDescription = "UndefinedWindowInteractionState #" & parameter1
   
   Case Exceptions.WindowsDriverPatternNotHandled
     ExceptionName = "WindowsDriverPatternNotHandled"
     ErrorDescription = "Windows Driver UIAutomation Pattern #" & parameter1 & " & (" & parameter2 & ") not handled!"
   
   Case Else
     ExceptionName = "UnhandledException"
     ErrorDescription = "Unhandled Exception for exception code #" & exception
     ErrorNumber = Phosphorus.Exceptions.UnhandledException
  
  End Select
  
  ErrorNumber = VBA.Constants.vbObjectError + ErrorNumber
  ErrorDescription = ExceptionName & "Exception: " & ErrorDescription
  ErrorDescription = VBA.Strings.Trim(ErrorDescription)
  
  'Log the exception
  Phosphorus.Log4PStatic.GetLogger 'Ensure we have a logger!
  Logger.ExternalFatal ErrorDescription
  
  'Now raise a VBA error
  Err.Raise ErrorNumber, "pExceptions", ErrorDescription
 
 End Sub

Private Function TrimErrorMessageParameter(ByVal parameter As String) As String
  parameter = VBA.Strings.Replace(parameter, vbLf, " ")
  parameter = VBA.Strings.Replace(parameter, vbCr, " ")
  parameter = VBA.Strings.Replace(parameter, vbCrLf, " ")
  parameter = VBA.Strings.Trim(parameter)
  TrimErrorMessageParameter = parameter
End Function

