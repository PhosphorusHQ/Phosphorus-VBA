Attribute VB_Name = "WindowsDriver01"
'@Folder WindowsDriver
'@TestModule
Option Explicit

'@TestMethod
Public Sub TestInvalidGetWindowInteractionStateDescription()
'Check for no error raised if give an invalid WindowInteractionStateDescription
Arrange:
  On Error GoTo ErrorHandler
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = VBA.Constants.vbObjectError + Exceptions.WindowsDriverUndefinedWindowInteractionState
Act:
  Phosphorus.pWindowsDriverStatic.GetWindowInteractionStateDescription -1
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
