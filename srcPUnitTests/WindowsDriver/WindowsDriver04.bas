Attribute VB_Name = "WindowsDriver04"
'@Folder WindowsDriver
'@TestModule
Option Explicit

Private pWindowsDriver As Phosphorus.pWindowsDriver

Const WEB_APP_NAME = "Selectors Hub Xpath Practice Page"
Const TARGET_PAGE_URL = "https://selectorshub.com/xpath-practice-page/"
Const TARGET_PAGE_TITLE = "Xpath Practice Page"

Dim lstrCurrentPPath As String

Private Sub BeforeModule()
  On Error GoTo ErrorHandler
'  Set pWindowsDriver = Phosphorus.Factory.GetNewPDriver(Phosphorus.pWindowsDriverType.WebBrowser)
'  pWindowsDriver.WebBrowserType.driverType = Phosphorus.pWebBrowserType.MicrosoftEdge
'  pWindowsDriver.WebBrowserType.driverType = Phosphorus.pWebBrowserType.MicrosoftEdge
'  pWindowsDriver.WebBrowserType.InstanceType = Phosphorus.pInstanceType.ReuseACurrentOpenInstance
'
  'Check for an error raised if we give get invalid http status code
  pWindowsDriver.Launch WEB_APP_NAME, TARGET_PAGE_TITLE, TARGET_PAGE_URL & "invalid-url"

'TODO: Test for proper HTTP response in the before test?

  'Find the App window in the BeforeModule method so all tests are skipped if it fails
  lstrCurrentPPath = "/Window[xp:starts-with(@Name,""" & TARGET_PAGE_TITLE & """)]"
'TODO ... make the TIMEOUT for this page 30s as 10s is not always long enough for this page to load

  pWindowsDriver.FindElement lstrCurrentPPath

'Test 1 - check we can kill the page without doing anything ... and then reopen the page
'Test 2 - check for an incorrect title element not found and that this kills the window
'THEN let all other tests proceed only if these pass!

'NB The driver window is not closed if the page title element is not found in time
  Set pWindowsDriver = Nothing

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Private Sub AfterModule()
  On Error GoTo ErrorHandler
  pWindowsDriver.Terminate
  Set pWindowsDriver = Nothing
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Public Sub BeforeTest()
  On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Public Sub AfterTest()
  On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub


'@Tag(UIADriverAndElements)
'@TestMethod
Public Sub UIADriverAndElements01_001()
  On Error GoTo ErrorHandler
Arrange:

Act:
Assert:
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Test 2 - test what happens if the page title is not matched

