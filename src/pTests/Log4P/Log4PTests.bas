Attribute VB_Name = "Log4PTests"
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

Sub TestLogger1()

  ' Set a root folder that doesn’t exist (will be created)
  'Log4PStatic.LogRootFolder = ThisWorkbook.Path & "\Logs" - This will be the default folder anyway!

  Log4PStatic.LogFileNameDynamicPart1 = "Log4PTests"
  Log4PStatic.LogFileNameDynamicPart2 = "TestLogger1"

  Dim myLogger As Phosphorus.Log4P
  Set myLogger = Phosphorus.Factory.GetNewPhosphorusLog4P
  
  ' Test all log levels using public methods
  myLogger.level = LogLevel.Debugging ' Enable all levels
  
  myLogger.Info "Testing internal info", Internal
  myLogger.Warning "Testing internal warning", Internal
  
  myLogger.Info "Testing external info", External
  
  myLogger.Error "Testing internal error", Internal
  myLogger.Fatal "Testing internal fatal", Internal
    
  myLogger.Trace "Testing external trace", External
  
  myLogger.Debugging "Testing external debug", External
  
  myLogger.Info "Testing external info", External
  myLogger.Warning "Testing external warning", External
  
  
  myLogger.Error "Testing external error", External
  myLogger.Fatal "Testing external fatal", External, True  ' Force flush

  ' Test filtering
  myLogger.level = LogLevel.Error
  myLogger.Warning "This won’t log", Internal ' Skipped
  myLogger.Error "This will log", Internal

  myLogger.Flush
  Debug.Print "Log file: " & myLogger.GetFilePath
  Set myLogger = Nothing
  
End Sub

Sub TestLogger2()

  Log4PStatic.LogFileNameDynamicPart1 = "Log4PTests"
  Log4PStatic.LogFileNameDynamicPart2 = "TestLogger2"

  Dim myLogger As Phosphorus.Log4P
  Set myLogger = Phosphorus.Factory.GetNewPhosphorusLog4P
    
  ' Test all log levels with Unicode characters built using ChrW
  myLogger.level = -5 ' Enable all levels (INTERNAL_DEBUG)
  
  myLogger.Debugging "Testing internal debug with Unicode: " & _
    VBA.Strings.ChrW(&H3053) & VBA.Strings.ChrW(&H3093) & VBA.Strings.ChrW(&H306B) & VBA.Strings.ChrW(&H3061) & VBA.Strings.ChrW(&H306F) & _
    VBA.Strings.ChrW(&H4E16) & VBA.Strings.ChrW(&H754C) & _
    " " & VBA.Strings.ChrW(&HD83C) & VBA.Strings.ChrW(&HDF0D), _
    Internal
  myLogger.Info "Testing internal info with Unicode: " & _
    VBA.Strings.ChrW(&H41F) & VBA.Strings.ChrW(&H440) & VBA.Strings.ChrW(&H438) & VBA.Strings.ChrW(&H432) & VBA.Strings.ChrW(&H435) & VBA.Strings.ChrW(&H442) & _
    ", " & VBA.Strings.ChrW(&H43C) & VBA.Strings.ChrW(&H438) & VBA.Strings.ChrW(&H440) & VBA.Strings.ChrW(&H21), _
    Internal
  myLogger.Warning "Testing internal warning with Unicode: " & _
    VBA.Strings.ChrW(&H4F60) & VBA.Strings.ChrW(&H597D) & _
    VBA.Strings.ChrW(&HFF0C) & _
    VBA.Strings.ChrW(&H4E16) & VBA.Strings.ChrW(&H754C) & VBA.Strings.ChrW(&HFF01), Internal
  
  myLogger.Error "Testing internal error with Unicode: " & _
    VBA.Strings.ChrW(&HA1) & VBA.Strings.ChrW(&H48) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H6C) & VBA.Strings.ChrW(&H61) & _
    ", " & VBA.Strings.ChrW(&H6D) & VBA.Strings.ChrW(&H75) & VBA.Strings.ChrW(&H6E) & VBA.Strings.ChrW(&H64) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H21), _
    Internal
  myLogger.Fatal "Testing internal critical with Unicode: " & _
    VBA.Strings.ChrW(&H42) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H6E) & VBA.Strings.ChrW(&H6A) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H75) & VBA.Strings.ChrW(&H72) & _
    " " & VBA.Strings.ChrW(&H6C) & VBA.Strings.ChrW(&H65) & " " & _
    VBA.Strings.ChrW(&H6D) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H6E) & VBA.Strings.ChrW(&H64) & VBA.Strings.ChrW(&H65) & VBA.Strings.ChrW(&H21), _
    Internal ' monde!
  myLogger.Debugging "Testing application debug with Unicode: " & _
    VBA.Strings.ChrW(&H48) & VBA.Strings.ChrW(&H61) & VBA.Strings.ChrW(&H6C) & VBA.Strings.ChrW(&H6C) & VBA.Strings.ChrW(&H6F) & _
    " " & VBA.Strings.ChrW(&H57) & VBA.Strings.ChrW(&H65) & VBA.Strings.ChrW(&H6C) & VBA.Strings.ChrW(&H74) & VBA.Strings.ChrW(&H21), _
    Internal
  myLogger.Warning "Testing application warning with Unicode: " & _
    VBA.Strings.ChrW(&H645) & VBA.Strings.ChrW(&H631) & VBA.Strings.ChrW(&H62D) & VBA.Strings.ChrW(&H628) & VBA.Strings.ChrW(&H627) & _
    " " & VBA.Strings.ChrW(&H628) & VBA.Strings.ChrW(&H627) & VBA.Strings.ChrW(&H644) & VBA.Strings.ChrW(&H639) & VBA.Strings.ChrW(&H627) & VBA.Strings.ChrW(&H644) & VBA.Strings.ChrW(&H645), _
    External
  myLogger.Error "Testing application error with Unicode: " & _
    VBA.Strings.ChrW(&H43) & VBA.Strings.ChrW(&H69) & VBA.Strings.ChrW(&H61) & VBA.Strings.ChrW(&H6F) & _
    " " & VBA.Strings.ChrW(&H6D) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H6E) & VBA.Strings.ChrW(&H64) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H21), _
    External
  myLogger.Fatal "Testing application critical with Unicode: " & _
    VBA.Strings.ChrW(&H4F) & VBA.Strings.ChrW(&H6C) & VBA.Strings.ChrW(&HE1) & _
    " " & VBA.Strings.ChrW(&H4D) & VBA.Strings.ChrW(&H75) & VBA.Strings.ChrW(&H6E) & VBA.Strings.ChrW(&H64) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H21), _
    Internal, True ' Mundo!
    
  ' Test filtering
  myLogger.level = LogLevel.Error
  myLogger.Warning "This won’t log", Internal ' Skipped
  myLogger.Error "This will log with Unicode: " & _
    VBA.Strings.ChrW(&H417) & VBA.Strings.ChrW(&H434) & VBA.Strings.ChrW(&H440) & VBA.Strings.ChrW(&H430) & VBA.Strings.ChrW(&H432) & VBA.Strings.ChrW(&H441) & VBA.Strings.ChrW(&H442) & _
    VBA.Strings.ChrW(&H432) & VBA.Strings.ChrW(&H443) & VBA.Strings.ChrW(&H439) & VBA.Strings.ChrW(&H442) & VBA.Strings.ChrW(&H435) & VBA.Strings.ChrW(&H21), _
    Internal
 
  myLogger.Flush
  Debug.Print "Log file: " & myLogger.GetFilePath
  Set myLogger = Nothing

End Sub

