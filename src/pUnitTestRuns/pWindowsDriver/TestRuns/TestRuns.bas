Attribute VB_Name = "TestRuns"
'@Folder TestRuns
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Sub RunSingleTest()
  pUnit.TestRunner.RunAllTests "WindowsDriver03LaunchWDs", "ValidPageLoadElement", "pUnitTests_pWindowsDriver"
End Sub

Sub RunAllTests()
  pUnit.TestRunner.RunAllTests "WindowsDriver02PSPipe", "", "pUnitTests_pWindowsDriver"
End Sub

Sub TestPowerShellPipeClient()
  pUnit.TestRunner.RunAllTests "WindowsDriver01PSPipe", "TestWindowsPowerShellPipeClient_InvalidCommand", "pUnitTests_pWindowsDriver"
End Sub

Sub RunSingleTest2()
  pUnit.TestRunner.RunAllTests "WindowsDriver01", "TestInvalidGetWindowInteractionStateDescription", "pUnitTests_pWindowsDriver"
End Sub

