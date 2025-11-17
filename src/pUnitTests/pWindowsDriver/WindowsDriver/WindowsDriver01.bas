Attribute VB_Name = "WindowsDriver01"
'@Folder WindowsDriver
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

'@TestMethod
Public Sub TestInvalidGetWindowInteractionStateDescription()
'Check for no error raised if give an invalid WindowInteractionStateDescription
Arrange:
  On Error GoTo ErrorHandler
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = VBA.Constants.vbObjectError + Exceptions.WindowsDriverUndefinedWindowInteractionState
Act:
  pWinDriver.pWindowsDriverStatic.GetWindowInteractionStateDescription -1
  Exit Sub
ErrorHandler:
Assert:
  Dim ErrorNumber As Long
  ErrorNumber = Err.Number
  Phosphorus.AssertionsStatic.pAssert.Equal ExpectedErrorNumber, ErrorNumber, isCritical:=True
  If ErrorNumber <> ExpectedErrorNumber Then
    Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
  End If
End Sub
