Attribute VB_Name = "Log4PTest"
'@Folder Logging
Option Explicit

Sub TestLogger1()

  ' Set a root folder that doesn’t exist (will be created)
  'LogRootFolder = ThisWorkbook.Path & "\Logs" - This will be the default folder anyway!

  Dim myLogger As Log4P
  Set myLogger = New Log4P

  ' Test all log levels using public methods
  myLogger.level = -5 ' Enable all levels (INTERNAL_DEBUG)
  myLogger.InternalTrace "Testing internal trace"
  myLogger.InternalDebug "Testing internal debug"
  myLogger.InternalInfo "Testing internal info"
  myLogger.InternalWarning "Testing internal warning"
  myLogger.InternalError "Testing internal error"
  myLogger.InternalFatal "Testing internal fatal"
  myLogger.ExternalTrace "Testing external debug"
  myLogger.ExternalDebug "Testing external debug"
  myLogger.ExternalInfo "Testing external info"
  myLogger.ExternalWarning "Testing external warning"
  myLogger.ExternalError "Testing external error"
  myLogger.ExternalFatal "Testing external critical", True ' Force flush

  ' Test filtering
  myLogger.level = -2 ' INTERNAL_ERROR
  myLogger.InternalWarning "This won’t log" ' Skipped
  myLogger.InternalError "This will log"

  myLogger.Flush
  Debug.Print "Log file: " & myLogger.GetFilePath

End Sub

Sub TestLogger2()

  Dim myLogger As Log4P
  Set myLogger = New Log4P
    
  ' Test all log levels with Unicode characters built using ChrW
  myLogger.level = -5 ' Enable all levels (INTERNAL_DEBUG)
  myLogger.InternalDebug "Testing internal debug with Unicode: " & _
    ChrW(&H3053) & ChrW(&H3093) & ChrW(&H306B) & ChrW(&H3061) & ChrW(&H306F) & _
    ChrW(&H4E16) & ChrW(&H754C) & _
    " " & ChrW(&HD83C) & ChrW(&HDF0D)
  myLogger.InternalInfo "Testing internal info with Unicode: " & _
    ChrW(&H41F) & ChrW(&H440) & ChrW(&H438) & ChrW(&H432) & ChrW(&H435) & ChrW(&H442) & _
    ", " & ChrW(&H43C) & ChrW(&H438) & ChrW(&H440) & ChrW(&H21)
  myLogger.InternalWarning "Testing internal warning with Unicode: " & _
    ChrW(&H4F60) & ChrW(&H597D) & _
    ChrW(&HFF0C) & _
    ChrW(&H4E16) & ChrW(&H754C) & ChrW(&HFF01)
  myLogger.InternalError "Testing internal error with Unicode: " & _
    ChrW(&HA1) & ChrW(&H48) & ChrW(&H6F) & ChrW(&H6C) & ChrW(&H61) & _
    ", " & ChrW(&H6D) & ChrW(&H75) & ChrW(&H6E) & ChrW(&H64) & ChrW(&H6F) & ChrW(&H21)
  myLogger.InternalFatal "Testing internal critical with Unicode: " & _
    ChrW(&H42) & ChrW(&H6F) & ChrW(&H6E) & ChrW(&H6A) & ChrW(&H6F) & ChrW(&H75) & ChrW(&H72) & _
    " " & ChrW(&H6C) & ChrW(&H65) & " " & _
    ChrW(&H6D) & ChrW(&H6F) & ChrW(&H6E) & ChrW(&H64) & ChrW(&H65) & ChrW(&H21) ' monde!
  myLogger.ExternalDebug "Testing application debug with Unicode: " & _
    ChrW(&H48) & ChrW(&H61) & ChrW(&H6C) & ChrW(&H6C) & ChrW(&H6F) & _
    " " & ChrW(&H57) & ChrW(&H65) & ChrW(&H6C) & ChrW(&H74) & ChrW(&H21)
  myLogger.ExternalWarning "Testing application warning with Unicode: " & _
    ChrW(&H645) & ChrW(&H631) & ChrW(&H62D) & ChrW(&H628) & ChrW(&H627) & _
    " " & ChrW(&H628) & ChrW(&H627) & ChrW(&H644) & ChrW(&H639) & ChrW(&H627) & ChrW(&H644) & ChrW(&H645)
  myLogger.ExternalError "Testing application error with Unicode: " & _
    ChrW(&H43) & ChrW(&H69) & ChrW(&H61) & ChrW(&H6F) & _
    " " & ChrW(&H6D) & ChrW(&H6F) & ChrW(&H6E) & ChrW(&H64) & ChrW(&H6F) & ChrW(&H21)
  myLogger.ExternalFatal "Testing application critical with Unicode: " & _
    ChrW(&H4F) & ChrW(&H6C) & ChrW(&HE1) & _
    " " & ChrW(&H4D) & ChrW(&H75) & ChrW(&H6E) & ChrW(&H64) & ChrW(&H6F) & ChrW(&H21), True ' Mundo!
    
  ' Test filtering
  myLogger.level = -2 ' INTERNAL_ERROR
  myLogger.InternalWarning "This won’t log" ' Skipped
  myLogger.InternalError "This will log with Unicode: " & _
    ChrW(&H417) & ChrW(&H434) & ChrW(&H440) & ChrW(&H430) & ChrW(&H432) & ChrW(&H441) & ChrW(&H442) & _
    ChrW(&H432) & ChrW(&H443) & ChrW(&H439) & ChrW(&H442) & ChrW(&H435) & ChrW(&H21)  ' !
 
  myLogger.Flush
  Debug.Print "Log file: " & myLogger.GetFilePath

End Sub

