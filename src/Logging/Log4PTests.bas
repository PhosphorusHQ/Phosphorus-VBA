Attribute VB_Name = "Log4PTests"
'@Folder Logging
Option Explicit

Sub TestLogger1()

  ' Set a root folder that doesn’t exist (will be created)
  'Log4PStatic.LogRootFolder = ThisWorkbook.Path & "\Logs" - This will be the default folder anyway!

  Log4PStatic.LogFileNameDynamicPart1 = "Log4PTests"
  Log4PStatic.LogFileNameDynamicPart2 = "TestLogger1"

  Dim myLogger As Log4P
  Set myLogger = New Log4P

  ' Test all log levels using public methods
  myLogger.level = -5 ' Enable all levels (INTERNAL_DEBUG)
  
  myLogger.NextAnalysisCodeLevel "Group1"
  myLogger.NextAnalysisCodeLevel "Group1.1"
  myLogger.NextAnalysisCodeLevel "Group1.1.1"
  myLogger.InternalTrace "Testing internal trace"
  myLogger.PreviousAnalysisCodeLevel
  myLogger.NextAnalysisCodeLevel "Group1.1.2"
  myLogger.PreviousAnalysisCodeLevel
  myLogger.InternalDebug "Testing internal debug"
  myLogger.PreviousAnalysisCodeLevel
  myLogger.NextAnalysisCodeLevel "Group1.2"
  myLogger.InternalInfo "Testing internal info"
  myLogger.InternalWarning "Testing internal warning"
  
  myLogger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="Test Output", _
    AnalysisCode2:="Test1", _
    message:="Fixed Message"
  
  myLogger.PreviousAnalysisCodeLevel
  myLogger.NextAnalysisCodeLevel "Group1.3"
  myLogger.InternalError "Testing internal error"
  myLogger.InternalFatal "Testing internal fatal"
  
  myLogger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="Test Output", _
    AnalysisCode2:="Test2", _
    message:="Fixed Message"
  
  myLogger.PreviousAnalysisCodeLevel
  myLogger.PreviousAnalysisCodeLevel
  myLogger.NextAnalysisCodeLevel "Group2"
  myLogger.ExternalTrace "Testing external debug"
  myLogger.ExternalDebug "Testing external debug"
  myLogger.PreviousAnalysisCodeLevel
  myLogger.ExternalInfo "Testing external info"
  myLogger.ExternalWarning "Testing external warning"
  myLogger.PreviousAnalysisCodeLevel
  
  myLogger.NextAnalysisCodeLevel "Group3"
  myLogger.ExternalError "Testing external error"
  myLogger.ExternalFatal "Testing external critical", True ' Force flush
  myLogger.PreviousAnalysisCodeLevel

  ' Test filtering
  myLogger.level = -2 ' INTERNAL_ERROR
  myLogger.InternalWarning "This won’t log" ' Skipped
  myLogger.InternalError "This will log"

  myLogger.Flush
  Debug.Print "Log file: " & myLogger.GetFilePath
  Set myLogger = Nothing
  
End Sub

Sub TestLogger2()

  Log4PStatic.LogFileNameDynamicPart1 = "Log4PTests"
  Log4PStatic.LogFileNameDynamicPart2 = "TestLogger2"

  Dim myLogger As Log4P
  Set myLogger = New Log4P
    
  ' Test all log levels with Unicode characters built using ChrW
  myLogger.level = -5 ' Enable all levels (INTERNAL_DEBUG)
  
  myLogger.NextAnalysisCodeLevel "Group1"
  myLogger.InternalDebug "Testing internal debug with Unicode: " & _
    VBA.Strings.ChrW(&H3053) & VBA.Strings.ChrW(&H3093) & VBA.Strings.ChrW(&H306B) & VBA.Strings.ChrW(&H3061) & VBA.Strings.ChrW(&H306F) & _
    VBA.Strings.ChrW(&H4E16) & VBA.Strings.ChrW(&H754C) & _
    " " & VBA.Strings.ChrW(&HD83C) & VBA.Strings.ChrW(&HDF0D)
  myLogger.InternalInfo "Testing internal info with Unicode: " & _
    VBA.Strings.ChrW(&H41F) & VBA.Strings.ChrW(&H440) & VBA.Strings.ChrW(&H438) & VBA.Strings.ChrW(&H432) & VBA.Strings.ChrW(&H435) & VBA.Strings.ChrW(&H442) & _
    ", " & VBA.Strings.ChrW(&H43C) & VBA.Strings.ChrW(&H438) & VBA.Strings.ChrW(&H440) & VBA.Strings.ChrW(&H21)
  myLogger.InternalWarning "Testing internal warning with Unicode: " & _
    VBA.Strings.ChrW(&H4F60) & VBA.Strings.ChrW(&H597D) & _
    VBA.Strings.ChrW(&HFF0C) & _
    VBA.Strings.ChrW(&H4E16) & VBA.Strings.ChrW(&H754C) & VBA.Strings.ChrW(&HFF01)
  myLogger.PreviousAnalysisCodeLevel
  
  myLogger.NextAnalysisCodeLevel "Group2"
  myLogger.InternalError "Testing internal error with Unicode: " & _
    VBA.Strings.ChrW(&HA1) & VBA.Strings.ChrW(&H48) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H6C) & VBA.Strings.ChrW(&H61) & _
    ", " & VBA.Strings.ChrW(&H6D) & VBA.Strings.ChrW(&H75) & VBA.Strings.ChrW(&H6E) & VBA.Strings.ChrW(&H64) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H21)
  myLogger.InternalFatal "Testing internal critical with Unicode: " & _
    VBA.Strings.ChrW(&H42) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H6E) & VBA.Strings.ChrW(&H6A) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H75) & VBA.Strings.ChrW(&H72) & _
    " " & VBA.Strings.ChrW(&H6C) & VBA.Strings.ChrW(&H65) & " " & _
    VBA.Strings.ChrW(&H6D) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H6E) & VBA.Strings.ChrW(&H64) & VBA.Strings.ChrW(&H65) & VBA.Strings.ChrW(&H21) ' monde!
  myLogger.ExternalDebug "Testing application debug with Unicode: " & _
    VBA.Strings.ChrW(&H48) & VBA.Strings.ChrW(&H61) & VBA.Strings.ChrW(&H6C) & VBA.Strings.ChrW(&H6C) & VBA.Strings.ChrW(&H6F) & _
    " " & VBA.Strings.ChrW(&H57) & VBA.Strings.ChrW(&H65) & VBA.Strings.ChrW(&H6C) & VBA.Strings.ChrW(&H74) & VBA.Strings.ChrW(&H21)
  myLogger.PreviousAnalysisCodeLevel
  
  myLogger.NextAnalysisCodeLevel "Group3"
  myLogger.ExternalWarning "Testing application warning with Unicode: " & _
    VBA.Strings.ChrW(&H645) & VBA.Strings.ChrW(&H631) & VBA.Strings.ChrW(&H62D) & VBA.Strings.ChrW(&H628) & VBA.Strings.ChrW(&H627) & _
    " " & VBA.Strings.ChrW(&H628) & VBA.Strings.ChrW(&H627) & VBA.Strings.ChrW(&H644) & VBA.Strings.ChrW(&H639) & VBA.Strings.ChrW(&H627) & VBA.Strings.ChrW(&H644) & VBA.Strings.ChrW(&H645)
  myLogger.ExternalError "Testing application error with Unicode: " & _
    VBA.Strings.ChrW(&H43) & VBA.Strings.ChrW(&H69) & VBA.Strings.ChrW(&H61) & VBA.Strings.ChrW(&H6F) & _
    " " & VBA.Strings.ChrW(&H6D) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H6E) & VBA.Strings.ChrW(&H64) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H21)
  myLogger.ExternalFatal "Testing application critical with Unicode: " & _
    VBA.Strings.ChrW(&H4F) & VBA.Strings.ChrW(&H6C) & VBA.Strings.ChrW(&HE1) & _
    " " & VBA.Strings.ChrW(&H4D) & VBA.Strings.ChrW(&H75) & VBA.Strings.ChrW(&H6E) & VBA.Strings.ChrW(&H64) & VBA.Strings.ChrW(&H6F) & VBA.Strings.ChrW(&H21), True ' Mundo!
  myLogger.PreviousAnalysisCodeLevel
    
  myLogger.NextAnalysisCodeLevel "Group4"
  ' Test filtering
  myLogger.level = -2 ' INTERNAL_ERROR
  myLogger.InternalWarning "This won’t log" ' Skipped
  myLogger.InternalError "This will log with Unicode: " & _
    VBA.Strings.ChrW(&H417) & VBA.Strings.ChrW(&H434) & VBA.Strings.ChrW(&H440) & VBA.Strings.ChrW(&H430) & VBA.Strings.ChrW(&H432) & VBA.Strings.ChrW(&H441) & VBA.Strings.ChrW(&H442) & _
    VBA.Strings.ChrW(&H432) & VBA.Strings.ChrW(&H443) & VBA.Strings.ChrW(&H439) & VBA.Strings.ChrW(&H442) & VBA.Strings.ChrW(&H435) & VBA.Strings.ChrW(&H21)  ' !
  myLogger.PreviousAnalysisCodeLevel
 
  myLogger.Flush
  Debug.Print "Log file: " & myLogger.GetFilePath
  Set myLogger = Nothing

End Sub

