Attribute VB_Name = "TestModule01"
'@Folder pUnit
'@TestModule
Option Explicit

'@SmokeTest
'@TestMethod
Public Sub TestAddition()
  On Error GoTo ErrorHandler
  Dim result As Integer
  result = 2 + 2
  Phosphorus.AssertionsStatic.pAssert.Equal 4, result, "2 + 2 should equal 4"
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Slow
'@TestMethod
Public Sub TestDivision()
  On Error GoTo ErrorHandler
  Dim result As Double
  result = 10 / 2
  Phosphorus.AssertionsStatic.pAssert.Equal 5, result, "10 / 2 should equal 5"
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(Regression,Slow)
'@TestMethod
Public Sub TestFailure()
  On Error GoTo ErrorHandler
  Phosphorus.AssertionsStatic.pAssert.IsTrue False, "This test is designed to fail"
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub
