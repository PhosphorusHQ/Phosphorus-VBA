Attribute VB_Name = "AssertionsTests"
'@Folder Assertions
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Sub RunAllAssertionTests()

  Phosphorus.Log4PStatic.LogFileNameDynamicPart1 = "AssertionsTests"
  Log4PStatic.LogFileNameDynamicPart2 = "RunAssertionTests"

  Phosphorus.Log4PStatic.GetLogger
  Logger.level = LogLevel.Trace 'Log everything
      
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
  Logger.Info "Numbers should be equal", External
  pAssert.Equal 5, 5, "Numbers should be equal", isCritical:=True
  
  Logger.Info "Strings should be equal", External
  pAssert.Equal "test", "test", "Strings should be equal", isCritical:=True
  
  Logger.Info "Numbers should not be equal", External
  pAssert.NotEqual 5, 6, "Numbers should not be equal", isCritical:=True
  
  Logger.Info "Basic math should be true", External
  pAssert.IsTrue (2 + 2 = 4), "Basic math should be true", isCritical:=True
  
  Logger.Info "Comparison should be false", External
  pAssert.IsFalse (1 > 2), "Comparison should be false", isCritical:=True
  
  Logger.Info "Value should be null", External
  pAssert.ThisIsNull Null, "Value should be null", isCritical:=True

  Logger.Info "Value should not be null", External
  pAssert.ThisIsNotNull "hello", "Value should not be null", isCritical:=True

  ' Intentionally failing tests for demonstration
  
  Logger.Info "Numbers should not be equal - This should fail", External
  pAssert.Equal 5, 6, "This should fail", isCritical:=True
  
  Logger.Info "True is not False - This should fail too", External
  pAssert.IsTrue False, "This should fail too", isCritical:=True

  ' Print results
  pAssert.PrintSummary
  
  Set pAssert = Nothing

End Sub

Sub RunAssertionTests2()
    
  Phosphorus.AssertionsStatic.GetAssert
  
  ' Test Step 1
  Logger.Info "Test Step 1", External
    
  Logger.Info "Assert 5 = 5", External
  pAssert.Equal 5, 5, "Numbers should be equal", isCritical:=True

  Logger.Info "Assert 2 + 2 = 4", External
  pAssert.IsTrue (2 + 2 = 4), "Basic math should be true", isCritical:=True
    
  Logger.Info "Assert 1>3", External
  pAssert.IsFalse (1 > 2), "Comparison should be false", isCritical:=True
  
  'No Critical Failures
  If pAssert.HasCriticalFailure Then
    Logger.Info "RunAssertionTests2 Step 1 has critical failures", External
    Exit Sub
  End If
    
  ' Test Step 2
  Logger.Info "Test Step 2", External

  Logger.Info "Numbers should not be equal - this should fail - critical", External
  pAssert.Equal 5, 6, "This should fail", isCritical:=True ' Critical assertion that will fail
    
  Logger.Info "The same string is not not equal - this should fail too", External
  pAssert.NotEqual "test", "test", "This should fail too", isCritical:=True

  If pAssert.HasCriticalFailure Then
    Logger.Info "RunAssertionTests2 Step 2 aborted due to critical failure", External
    GoTo PrintSummary
  End If
    
  ' Test Step 3 (won't run if critical failure occurred)
  Logger.Info "Test Step 3 (won't run if critical failure occurred)", External
 
  Logger.Info "A string is not null - this should not run", External
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
  Logger.Info "Test Step 1", External
 
  Logger.Info "Assert 5 is between 1 and 10", External
  pAssert.InRange 5, 1, 10, "Value should be in range", isCritical:=True

  Logger.Info "Array should contain 2", External
  pAssert.Contains testArray, 2, "Array should contain 2", isCritical:=True

  Logger.Info "'Hello123' contain '*[0-9]*'", External
  pAssert.StringMatchesSimplePattern "Hello123", "*[0-9]*", "String should contain numbers", isCritical:=True

  Logger.Info "Values should be approximately equal", External
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

  Logger.Info "Collection should be empty", External
  Dim testCollection As New collection
  pAssert.IsEmpty testCollection, "Collection should be empty", isCritical:=True

  Logger.Info "Array should be empty", External
  Dim emptyArray() As Integer
  pAssert.IsEmpty emptyArray(), "Array should be empty", isCritical:=True
    
  Logger.Info "Should be string type", External
  pAssert.IsType "Hello", "String", "Should be string type", isCritical:=True

  Logger.Info "10 should be greater than 5", External
  pAssert.GreaterThan 10, 5, "10 should be greater than 5", isCritical:=True
  
  Logger.Info "3 should be less than 7", External
  pAssert.LessThan 3, 7, "3 should be less than 7", isCritical:=True
  
  Logger.Info "Windows Explorer should exist", External
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

