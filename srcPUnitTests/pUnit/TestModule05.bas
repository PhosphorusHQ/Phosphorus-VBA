Attribute VB_Name = "TestModule05"
'@Folder pUnit
'@TestModule
Option Explicit

'@TestModule
'@SmokeTest
'@Regression

Public Sub BeforeModule()
  Phosphorus.Logger.ExternalInfo "Running BeforeModule in pUnitSampleTestModule5"
  On Error GoTo ErrorHandler
  'Do nothing
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Public Sub AfterModule()
  On Error GoTo ErrorHandler
  Err.Raise vbObjectError + 1001, "AfterModule", "Intentional failure in AfterModule"
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Public Sub BeforeTest()
End Sub

Public Sub AfterTest()
End Sub

'@TestMethod
'@SmokeTest
Public Sub TestAddition()
  On Error GoTo ErrorHandler
  Dim result As Integer
  result = 2 + 2
  Phosphorus.AssertionsStatic.pAssert.Equal 4, result, "2 + 2 should equal 4"
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@TestMethod
'@Slow
Public Sub TestDivision()
  On Error GoTo ErrorHandler
  Dim result As Double
  result = 10 / 2
  Phosphorus.AssertionsStatic.pAssert.Equal 5, result, "10 / 2 should equal 5"
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@TestMethod
'@Regression
Public Sub TestFailure()
  On Error GoTo ErrorHandler
  Phosphorus.AssertionsStatic.pAssert.IsTrue False, "This test is designed to fail"
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

