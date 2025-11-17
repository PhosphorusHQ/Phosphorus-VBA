Attribute VB_Name = "TestModule09"
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
End Sub

Public Sub AfterModule()
End Sub

Public Sub BeforeTest()
  On Error GoTo ErrorHandler
  If Phosphorus.pUnit.CurrentTestName = "TestMultiplication" Then
    Err.Raise vbObjectError + 1002, "BeforeTest", "Intentional failure in BeforeTest for TestMultiplication"
  End If
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Public Sub AfterTest()
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
Public Sub TestSubtraction()
  On Error GoTo ErrorHandler
    Dim result As Integer
    result = 5 - 3
    Phosphorus.AssertionsStatic.pAssert.Equal 1, result, "5 - 3 should equal 2"
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub
