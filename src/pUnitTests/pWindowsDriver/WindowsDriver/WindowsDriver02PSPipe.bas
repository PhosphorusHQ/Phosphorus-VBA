Attribute VB_Name = "WindowsDriver02PSPipe"
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

Private Sub BeforeModule()
  On Error GoTo ErrorHandler
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Private Sub AfterModule()
  On Error GoTo ErrorHandler
  WindowsPowerShell.CloseWindowsPowerShellPipeClient
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@TestMethod
Public Sub TestWindowsPowerShellPipeClient_ValidCommands()
Arrange:
  On Error GoTo ErrorHandler
  Dim ThisMethodName As String
  ThisMethodName = "pUnitTests.WindowsDriver01PSPipe.TestWindowsPowerShellPipeClient_ValidCommands"
  Dim response1 As String
  Dim response2 As String
Act:
  response1 = Phosphorus.WindowsPowerShell.Execute("$x = 10; Write-Output $x", ThisMethodName)
  response2 = Phosphorus.WindowsPowerShell.Execute("$x = $x + 5; Write-Output $x", ThisMethodName)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(10), VBA.Conversion.CLng(response1), isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(15), VBA.Conversion.CLng(response2), isCritical:=True
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

' Cmd =
'@TestMethod
Public Sub TestWindowsPowerShellPipeClient_InvalidCommand()
Arrange:
  On Error GoTo ErrorHandler
  Dim ThisMethodName As String
  ThisMethodName = "pUnitTests.WindowsDriver01PSPipe.TestWindowsPowerShellPipeClient_InvalidCommand"
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = VBA.Constants.vbObjectError + Exceptions.PowerShellPipeClientPowerShellError
Act:
  Phosphorus.WindowsPowerShell.Execute "NOT Get-AppxPackage -Name '*AppName*'", ThisMethodName
Assert:
  Exit Sub
ErrorHandler:
  Dim ErrorNumber As Long
  ErrorNumber = Err.Number
  Phosphorus.AssertionsStatic.pAssert.Equal ExpectedErrorNumber, ErrorNumber, isCritical:=True
  If ErrorNumber <> ExpectedErrorNumber Then
    Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
  End If
End Sub

