Attribute VB_Name = "WindowsDriverTestRuns"
'@Folder WindowsDriver
Option Explicit

Sub RunSingleTest()
  pUnit.TestRunner.RunAllTests "WindowsDriver03LaunchWDs", "ValidPageLoadElement", "pUnitTests"
End Sub

Sub RunAllTests()
  pUnit.TestRunner.RunAllTests "WindowsDriver02PSPipe", "", "pUnitTests"
End Sub

Sub TestPowerShellPipeClient()
  pUnit.TestRunner.RunAllTests "WindowsDriver01PSPipe", "TestWindowsPowerShellPipeClient_InvalidCommand", "pUnitTests"
End Sub

Sub RunSingleTest2()
  pUnit.TestRunner.RunAllTests "WindowsDriver01", "TestInvalidGetWindowInteractionStateDescription", "pUnitTests"
End Sub

