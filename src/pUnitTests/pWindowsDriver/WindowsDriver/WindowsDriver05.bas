Attribute VB_Name = "WindowsDriver05"
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

Private pWindowsDriver As pWinDriver.pWindowsDriver
Const WEB_APP_NAME = "Selectors Hub Xpath Practice Page"
Const TARGET_PAGE_URL = "https://selectorshub.com/xpath-practice-page/"
Const TARGET_PAGE_TITLE = "Xpath Practice Page"

Private Sub BeforeModule()
  On Error GoTo ErrorHandler
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Private Sub AfterModule()
  On Error GoTo ErrorHandler
'  Set pWindowsDriver = Nothing
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub
 
Public Sub BeforeTest()
  On Error GoTo ErrorHandler
  Set pWindowsDriver = pWinDriver.pWindowsDriverStatic.GetNewPDriver(pWinDriver.pWindowsDriverType.WebBrowser)
  pWindowsDriver.SetWindowsDriverWebBrowserType _
    WebBrowserType:=pWinDriver.pWebBrowserType.MicrosoftEdge, _
    InstanceType:=pWinDriver.pInstanceType.ReuseACurrentOpenInstance
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Public Sub AfterTest()
  On Error GoTo ErrorHandler
  pWindowsDriver.Terminate
  Set pWindowsDriver = Nothing
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(UIADriverAndElements)
'@TestMethod
Public Sub UIADriverAndElements03_001()
'Check for no error if we open the URL in the current browser with a valid page load element - ReuseACurrentOpenInstance App Instance Type
Arrange:
'  On Error Resume Next
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = 0
  On Error GoTo ErrorHandler
Act:
  'Open the browser with the valid page
  pWindowsDriver.Launch WEB_APP_NAME, TARGET_PAGE_TITLE, TARGET_PAGE_URL, TimeoutInSeconds:=30
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

'Check for no error if we open the URL in the current browser with a valid page load element - Default(ReuseOpenInstance) App Instance Type


'@Tag(UIADriverAndElements)
'@TestMethod
Public Sub UIADriverAndElements03_004()
'Check for element not found error if we open the URL in the current browser with an invalid page load element
'Test 2 - check for an incorrect title element not found and that this kills the window
  On Error GoTo ErrorHandler
Arrange:
'  Dim InvalidPPath As String
'  CurrentPPath = "/Window[xp:starts-with(@Name,""" & TARGET_PAGE_TITLE & """)]"
'  InvalidPPath = CurrentPPath & "/NOTHING"
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = VBA.Constants.vbObjectError + Phosphorus.Exceptions.WindowsDriverUIElementNotFoundBeforeTimeout
Act:
  'Open the browser with the default page - Use a short timeout as is doesn't matter if the page hasn't fully loaded yet
  pWindowsDriver.Launch WEB_APP_NAME, "NOT " & TARGET_PAGE_TITLE, TARGET_PAGE_URL, TimeoutInSeconds:=5
  Exit Sub
ErrorHandler:
Assert:
  Dim ErrorNumber As Long
  ErrorNumber = Err.Number
  Phosphorus.AssertionsStatic.pAssert.Equal ExpectedErrorNumber, ErrorNumber, isCritical:=True
  If ErrorNumber <> ExpectedErrorNumber Then
    Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
  End If
  'Use the valid page load element PPath to get the PID from the found element so that the browser will get killed on termination of the driver
'TODO PJG  pWindowsDriver.FindElement CurrentPPath, TimeoutInSeconds:=30, GetPID:=True
End Sub

'@Tag(UIADriverAndElements)
'@TestMethod
Public Sub UIADriverAndElements03_005()
'Test when capped sleep exceeded
  On Error GoTo ErrorHandler
Arrange:
  Dim CappedSleepTimeoutToTestInSeconds As Long
  CappedSleepTimeoutToTestInSeconds = Configuration.GetValue("Processes", "CappedSleepTime", Phosphorus.WindowsProcesses.DEFAULT_CAPPED_SLEEP_TIME_IN_SECONDS) + 1
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = VBA.Constants.vbObjectError + Phosphorus.Exceptions.WindowsDriverCappedSleepTimeoutInSecondsExceeded
  On Error GoTo ErrorHandler
Act:
  Phosphorus.WindowsProcesses.Snooze CappedSleepTimeoutToTestInSeconds
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

