Attribute VB_Name = "TestModule02"
'@Folder pUnit
'@TestModule
Option Explicit

Public Sub BeforeModule()
  ' Setup code for the module (e.g., initialize resources)
End Sub

Public Sub AfterModule()
  ' Teardown code for the module (e.g., clean up resources)
End Sub

Public Sub BeforeTest()
  ' Setup code for each test (e.g., reset state)
End Sub

Public Sub AfterTest()
  ' Teardown code for each test (e.g., clean up test state)
End Sub

'@Regression
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

'@SmokeTest
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

'@Slow
'@TestMethod
Public Sub TestFailure()
  On Error GoTo ErrorHandler
  Phosphorus.AssertionsStatic.pAssert.IsTrue False, "This test is designed to fail"
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub
