Attribute VB_Name = "TestModule08"
'@Folder pUnit
'@TestModule
Option Explicit

Public Sub BeforeModule()
  On Error GoTo ErrorHandler
  Err.Raise vbObjectError + 1000, "BeforeModule", "Intentional failure in BeforeModule"
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Public Sub AfterModule()
End Sub

Public Sub BeforeTest()
End Sub

Public Sub AfterTest()
End Sub

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

'@TestMethod
Public Sub TestMultiplication()
  On Error GoTo ErrorHandler
    Dim result As Integer
    result = 3 * 4
    Phosphorus.AssertionsStatic.pAssert.Equal 12, result, "3 * 4 should equal 12"
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

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

'@TestMethod
Public Sub TestFailure()
  On Error GoTo ErrorHandler
  Phosphorus.AssertionsStatic.pAssert.IsTrue False, "This test is designed to fail"
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub
