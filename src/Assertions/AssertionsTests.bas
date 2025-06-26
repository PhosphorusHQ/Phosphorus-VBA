Attribute VB_Name = "AssertionsTests"
'@Folder Assertions
' Assertions.cls
'Option Explicit

'Sub RunAssertionTests()
'
'  Log4PStatic.LogFileNameDynamicPart1 = "AssertionsTests"
'  Log4PStatic.LogFileNameDynamicPart2 = "RunAssertionTests"
'
'  Log4PStatic.GetLogger
'  Logger.level = LogLevel.EXTERNAL_TRACE 'Log everything
'
'  Dim Assert As New Assertions
'
'  ' Test examples
'  Assert.AssertEqual 5, 5, "Numbers should be equal"
'  Assert.AssertEqual "test", "test", "Strings should be equal"
'  Assert.AssertNotEqual 5, 6, "Numbers should not be equal"
'  Assert.AssertTrue (2 + 2 = 4), "Basic math should be true"
'  Assert.AssertFalse (1 > 2), "Comparison should be false"
'  Assert.AssertNull Null, "Value should be null"
'  Assert.AssertNotNull "hello", "Value should not be null"
'
'  ' Intentionally failing tests for demonstration
'  Assert.AssertEqual 5, 6, "This should fail"
'  Assert.AssertTrue False, "This should fail too"
'
'  ' Print results
'  Assert.PrintSummary
'
'  Logger.Flush
'  Debug.Print "Log file: " & Logger.GetFilePath
'
'  Log4PStatic.CloseLogger
'
'End Sub
' TestModule.bas
Option Explicit

Sub RunTests()
    Dim assert As New Assertions
    
    ' Test Step 1
    Debug.Print "Test Step 1:"
    assert.AssertEqual 5, 5, "Numbers should be equal"
    assert.AssertTrue (2 + 2 = 4), "Basic math should be true", True ' Critical assertion
    assert.AssertFalse (1 > 2), "Comparison should be false"
    If assert.HasCriticalFailure Then
        Debug.Print "Step 1 aborted due to critical failure"
        Exit Sub
    End If
    
    ' Test Step 2
    Debug.Print vbNewLine & "Test Step 2:"
    assert.AssertEqual 5, 6, "This should fail", True ' Critical assertion that will fail
    assert.AssertNotEqual "test", "test", "This should fail too"
    If assert.HasCriticalFailure Then
        Debug.Print "Step 2 aborted due to critical failure"
        Exit Sub
    End If
    
    ' Test Step 3 (won't run if critical failure occurred)
    Debug.Print vbNewLine & "Test Step 3:"
    assert.AssertNotNull "hello", "Value should not be null"
    
    ' Print results
    assert.PrintSummary
End Sub

' TestModule.bas
Sub TestAdditionalAssertions()
    Dim assert As New Assertions
    Dim testArray(1 To 3) As Integer
    testArray(1) = 1
    testArray(2) = 2
    testArray(3) = 3
    
    ' Test Step 1
    Debug.Print "Test Step 1:"
    assert.AssertInRange 5, 1, 10, "Value should be in range"
    assert.AssertContains testArray, 2, "Array should contain 2", True
    assert.AssertStringMatches "Hello123", "*[0-9]*", "String should contain numbers"
    assert.AssertApproximatelyEqual 10#, 10.1, 0.2, "Values should be approximately equal"
    assert.AssertObjectExists Application, "Application object should exist"
    
    ' Check for critical failure
    If assert.HasCriticalFailure Then
        Debug.Print "Step 1 aborted due to critical failure"
        Exit Sub
    End If
    
    assert.PrintSummary
End Sub

' TestModule.bas
Sub TestAllAssertions()
    Dim assert As New Assertions
    Dim emptyArray() As Integer
    Dim testCollection As New collection
    
    ' Test Step 1
    Debug.Print "Test Step 1:"
    assert.AssertThrows "Sub Test() : Err.Raise 5 : End Sub", 5, "Should raise error 5"
    assert.AssertEmpty testCollection, "Collection should be empty", True
    assert.AssertType "Hello", "String", "Should be string type"
    assert.AssertGreaterThan 10, 5, "10 should be greater than 5"
    assert.AssertLessThan 3, 7, "3 should be less than 7"
    assert.AssertFileExists "C:\Windows\explorer.exe", "Windows Explorer should exist"
    
    ' Check for critical failure
    If assert.HasCriticalFailure Then
        Debug.Print "Step 1 aborted due to critical failure"
        Exit Sub
    End If
    
    assert.PrintSummary
End Sub
