Attribute VB_Name = "TestModule06"
'@Folder pUnit
'@TestModule
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public Sub BeforeModule()
  On Error GoTo ErrorHandler
  'Do nothing
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Public Sub AfterModule()
  On Error GoTo ErrorHandler
  'Do nothing
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Public Sub BeforeTest()
  On Error GoTo ErrorHandler
  Err.Raise vbObjectError + 1003, "BeforeTest", "Intentional failure in BeforeTest"
  Exit Sub

ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
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

