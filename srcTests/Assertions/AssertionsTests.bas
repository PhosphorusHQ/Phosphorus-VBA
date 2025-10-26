Attribute VB_Name = "AssertionsTests"
'@Folder Assertions
Option Explicit

Sub RunAllAssertionTests()

  Phosphorus.Log4PStatic.LogFileNameDynamicPart1 = "AssertionsTests"
  Log4PStatic.LogFileNameDynamicPart2 = "RunAssertionTests"

  Phosphorus.Log4PStatic.GetLogger
  Logger.level = LogLevel.EXTERNAL_TRACE 'Log everything
      
  RunAssertionTests1
  RunAssertionTests2
  RunAssertionTests3
  RunAssertionTests4
  
  Logger.Flush
  Log4PStatic.CloseLogger

End Sub

Sub RunAssertionTests1()
  
  ' Test examples
  
  Phosphorus.AssertionsStatic.GetAssert
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests1", _
    AnalysisCode2:="AssertEqual", _
    AnalysisCode3:="Numbers", _
    message:="Numbers should be equal"
  pAssert.Equal 5, 5, "Numbers should be equal", isCritical:=True
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests1", _
    AnalysisCode2:="AssertEqual", _
    AnalysisCode3:="Strings", _
    message:="Strings should be equal"
  pAssert.Equal "test", "test", "Strings should be equal", isCritical:=True
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests1", _
    AnalysisCode2:="AssertNotEqual", _
    AnalysisCode3:="Numbers", _
    message:="Numbers should not be equal"
  pAssert.NotEqual 5, 6, "Numbers should not be equal", isCritical:=True
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests1", _
    AnalysisCode2:="AssertTrue", _
    AnalysisCode3:="Basic math", _
    message:="Basic math should be true"
  pAssert.IsTrue (2 + 2 = 4), "Basic math should be true", isCritical:=True
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests1", _
    AnalysisCode2:="AssertFalse", _
    AnalysisCode3:="Comparison", _
    message:="Comparison should be false"
  pAssert.IsFalse (1 > 2), "Comparison should be false", isCritical:=True
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests1", _
    AnalysisCode2:="AssertNull", _
    AnalysisCode3:="Comparison", _
    message:="Value should be null"
  pAssert.ThisIsNull Null, "Value should be null", isCritical:=True
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests1", _
    AnalysisCode2:="AssertNotNull", _
    AnalysisCode3:="Comparison", _
    message:="Value should not be null"
  pAssert.ThisIsNotNull "hello", "Value should not be null", isCritical:=True

  ' Intentionally failing tests for demonstration
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests1", _
    AnalysisCode2:="AssertEqual", _
    AnalysisCode3:="Comparison", _
    message:="Numbers should not be equal - This should fail"
  pAssert.Equal 5, 6, "This should fail", isCritical:=True
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests1", _
    AnalysisCode2:="AssertTrue", _
    AnalysisCode3:="False", _
    message:="True is not False - This should fail too"
  pAssert.IsTrue False, "This should fail too", isCritical:=True

  ' Print results
  pAssert.PrintSummary
  
  Set pAssert = Nothing

End Sub

Sub RunAssertionTests2()
    
  Phosphorus.AssertionsStatic.GetAssert
  
  ' Test Step 1
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests2", _
    AnalysisCode2:="Test Step 1", _
    message:="Test Step 1"
    
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests2", _
    AnalysisCode2:="Test Step 1", _
    AnalysisCode3:="Numbers should be equal", _
    message:="Assert 5 = 5"
  pAssert.Equal 5, 5, "Numbers should be equal", isCritical:=True

  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests2", _
    AnalysisCode2:="Test Step 1", _
    AnalysisCode3:="Basic math should be true", _
    message:="Assert 2 + 2 = 4"
  pAssert.IsTrue (2 + 2 = 4), "Basic math should be true", isCritical:=True
    
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests2", _
    AnalysisCode2:="Test Step 1", _
    AnalysisCode3:="Comparison should be false", _
    message:="Assert 2 + 2 = 4"
  pAssert.IsFalse (1 > 2), "Comparison should be false", isCritical:=True
  
  'No Critical Failures
  If pAssert.HasCriticalFailure Then
    Logger.LogFixedLevelMessage _
      level:=LogLevel.EXTERNAL_INFO, _
      AnalysisCode1:="AssertionTests2", _
      AnalysisCode2:="Check for Has Critical Failure", _
      message:="RunAssertionTests2 Step 1 has critical failures"
    Exit Sub
  End If
    
  ' Test Step 2
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests2", _
    AnalysisCode2:="Test Step 2", _
    message:="Test Step 2"
    
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests2", _
    AnalysisCode2:="Test Step 2", _
    AnalysisCode3:="Numbers should not be equal - this should fail - critical", _
    message:="Assert 5 = 5"
  pAssert.Equal 5, 6, "This should fail", isCritical:=True ' Critical assertion that will fail
    
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests2", _
    AnalysisCode2:="Test Step 2", _
    AnalysisCode3:="The same string is not not equal - this should fail too", _
    message:="Assert 5 = 5"
  pAssert.NotEqual "test", "test", "This should fail too", isCritical:=True
    
  If pAssert.HasCriticalFailure Then
    Logger.LogFixedLevelMessage _
      level:=LogLevel.EXTERNAL_INFO, _
      AnalysisCode1:="AssertionTests2", _
      AnalysisCode2:="Check for Has Critical Failure", _
      message:="RunAssertionTests2 Step 2 aborted due to critical failure"
    GoTo PrintSummary
  End If
    
  ' Test Step 3 (won't run if critical failure occurred)
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests2", _
    AnalysisCode2:="Test Step 3", _
    message:="Test Step 3"
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests2", _
    AnalysisCode2:="Test Step 3", _
    AnalysisCode3:="A string is not null - this should not run", _
    message:="Assert 'hello' is not "
  pAssert.ThisIsNotNull "hello", "Value should not be null", isCritical:=True

PrintSummary:
  ' Print results
  pAssert.PrintSummary

  Set pAssert = Nothing

End Sub

Sub RunAssertionTests3()
  
  Phosphorus.AssertionsStatic.GetAssert
  
  Dim testArray(1 To 3) As Integer
    testArray(1) = 1
    testArray(2) = 2
    testArray(3) = 3
    
  ' Test Step 1
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests3", _
    AnalysisCode2:="Test Step 1", _
    message:="Test Step 1"
    
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests3", _
    AnalysisCode2:="Test Step 1", _
    AnalysisCode3:="Value should be in range", _
    message:="Assert 5 is between 1 and 10"
  pAssert.InRange 5, 1, 10, "Value should be in range", isCritical:=True

  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests3", _
    AnalysisCode2:="Test Step 1", _
    AnalysisCode3:="Array should contain", _
    message:="Array should contain 2"
  pAssert.Contains testArray, 2, "Array should contain 2", isCritical:=True

  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests3", _
    AnalysisCode2:="Test Step 1", _
    AnalysisCode3:="String should contain numbers", _
    message:="'Hello123' contain '*[0-9]*'"
  pAssert.StringMatchesSimplePattern "Hello123", "*[0-9]*", "String should contain numbers", isCritical:=True
    
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests3", _
    AnalysisCode2:="Test Step 1", _
    AnalysisCode3:="Values should be approximately equal", _
    message:="'Hello123' contain '*[0-9]*'"
  pAssert.ApproximatelyEqual 10#, 10.1, 0.2, "Values should be approximately equal", isCritical:=True
    
  pAssert.ObjectExists Application, "Application object should exist"
    
  ' Check for critical failure
  If pAssert.HasCriticalFailure Then
    Debug.Print "Step 1 aborted due to critical failure"
    Exit Sub
  End If
    
  pAssert.PrintSummary

  Set pAssert = Nothing

End Sub

Sub RunAssertionTests4()

  Phosphorus.AssertionsStatic.GetAssert
  
  pAssert.Throws "AssertionThrowsErrTest", 5, "Should raise error 5"
    
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests4", _
    AnalysisCode2:="AssertEmpty", _
    AnalysisCode3:="Collection", _
    message:="Collection should be empty"
  Dim testCollection As New collection
  pAssert.IsEmpty testCollection, "Collection should be empty", isCritical:=True
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests4", _
    AnalysisCode2:="AssertEmpty", _
    AnalysisCode3:="Array", _
    message:="Array should be empty"
  Dim emptyArray() As Integer
  pAssert.IsEmpty emptyArray(), "Array should be empty", isCritical:=True
    
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests4", _
    AnalysisCode2:="AssertType", _
    AnalysisCode3:="String", _
    message:="Should be string type"
  pAssert.IsType "Hello", "String", "Should be string type", isCritical:=True
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests4", _
    AnalysisCode2:="AssertGreaterThan", _
    AnalysisCode3:="Numbers 10, 5", _
    message:="10 should be greater than 5"
  pAssert.GreaterThan 10, 5, "10 should be greater than 5", isCritical:=True
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests4", _
    AnalysisCode2:="AssertLessThan", _
    AnalysisCode3:="Numbers 3, 7", _
    message:="3 should be less than 7"
  pAssert.LessThan 3, 7, "3 should be less than 7", isCritical:=True
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="AssertionTests4", _
    AnalysisCode2:="AssertFileExists", _
    AnalysisCode3:="C:\Windows\explorer.exe", _
    message:="Windows Explorer should exist"
  pAssert.FileExists "C:\Windows\explorer.exe", "Windows Explorer should exist", isCritical:=True
    
  ' Check for critical failure
  If pAssert.HasCriticalFailure Then
    Debug.Print "Step 1 aborted due to critical failure"
    Exit Sub
  End If
    
  pAssert.PrintSummary
  
  Set pAssert = Nothing
    
End Sub

'Test method called in Assert.Throws test above
Public Sub AssertionThrowsErrTest()
  On Error Resume Next
  Err.Raise 5
End Sub

