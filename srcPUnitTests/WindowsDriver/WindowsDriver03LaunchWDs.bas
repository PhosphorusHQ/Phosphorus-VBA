Attribute VB_Name = "WindowsDriver03LaunchWDs"
'@Folder WindowsDriver
'@TestModule
Option Explicit

Private pWindowsDriver As Phosphorus.pWindowsDriver
Const WEB_APP_NAME = "Example.com"
Const TARGET_PAGE_URL = "https://www.example.com"
Const TARGET_PAGE_TITLE = "Example Domain"

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
 
Public Sub BeforeTest()
  On Error GoTo ErrorHandler
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

'@TestMethod
Public Sub ValidURL()
'Check for no error raised if we give a valid url - use msedge with default options
Arrange:
  On Error GoTo ErrorHandler
  Set pWindowsDriver = Phosphorus.Factory.GetNewPDriver(Phosphorus.pWindowsDriverType.WebBrowser)
  pWindowsDriver.SetWindowsDriverWebBrowserType WebBrowserType:=Phosphorus.pWebBrowserType.MicrosoftEdge
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = 0
Act:
  'Open the browser with a valid url
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

'@TestMethod
Public Sub InvalidURL1()
'Check for the error raised if we get an invalid url - use msedge with default options
Arrange:
  On Error GoTo ErrorHandler
  Set pWindowsDriver = Phosphorus.Factory.GetNewPDriver(Phosphorus.pWindowsDriverType.WebBrowser)
  pWindowsDriver.SetWindowsDriverWebBrowserType WebBrowserType:=Phosphorus.pWebBrowserType.MicrosoftEdge
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = VBA.Constants.vbObjectError + Exceptions.WindowsDriverInvalidHTTPStatus
Act:
  'Open the browser with an invalid url - valid format, syntactically correct
  pWindowsDriver.Launch WEB_APP_NAME, TARGET_PAGE_TITLE, TARGET_PAGE_URL & "/invalid-url", TimeoutInSeconds:=30
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

'@TestMethod
Public Sub InvalidURL2()
'Check for the error raised if we get an invalid url - use msedge with default options
Arrange:
  On Error GoTo ErrorHandler
  Set pWindowsDriver = Phosphorus.Factory.GetNewPDriver(Phosphorus.pWindowsDriverType.WebBrowser)
  pWindowsDriver.SetWindowsDriverWebBrowserType WebBrowserType:=Phosphorus.pWebBrowserType.MicrosoftEdge
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = VBA.Constants.vbObjectError + Exceptions.WindowsDriverInvalidURL
Act:
  'Open the browser with an invalid url - valid format, syntactically incorrect
  pWindowsDriver.Launch WEB_APP_NAME, TARGET_PAGE_TITLE, TARGET_PAGE_URL & "invalid-url", TimeoutInSeconds:=30
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

'@TestMethod
Public Sub ValidPageLoadElement()
'Check for no error if we open the URL in the current browser with a valid page load element - ReuseACurrentOpenInstance App Instance Type
Arrange:
  On Error GoTo ErrorHandler
  Set pWindowsDriver = Phosphorus.Factory.GetNewPDriver(Phosphorus.pWindowsDriverType.WebBrowser)
  pWindowsDriver.SetWindowsDriverWebBrowserType WebBrowserType:=Phosphorus.pWebBrowserType.MicrosoftEdge, _
    InstanceType:=Phosphorus.pInstanceType.ReuseACurrentOpenInstance
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = 0
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

'@TestMethod
Public Sub InvalidPageLoadElement()
'Check for element not found error if we open the URL in the current browser with an invalid page load element
Arrange:
  On Error GoTo ErrorHandler
  Set pWindowsDriver = Phosphorus.Factory.GetNewPDriver(Phosphorus.pWindowsDriverType.WebBrowser)
  pWindowsDriver.SetWindowsDriverWebBrowserType WebBrowserType:=Phosphorus.pWebBrowserType.MicrosoftEdge
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = VBA.Constants.vbObjectError + Phosphorus.Exceptions.WindowsDriverUIElementNotFoundBeforeTimeout
Act:
  'Open the browser - Use a short timeout as is doesn't matter if the page hasn't fully loaded yet
  pWindowsDriver.Launch WEB_APP_NAME, "NOT " & TARGET_PAGE_TITLE, TARGET_PAGE_URL, TimeoutInSeconds:=5
  'Resumes here after Asser if all ok
  On Error GoTo ErrorHandler1
  ExpectedErrorNumber = 0
  'Use the valid page load element PPath to get the PID from the found element so that the browser will get killed on termination of the driver
  pWindowsDriver.FindElement "/Window[xp:starts-with(@Name,""" & TARGET_PAGE_TITLE & """)]", TimeoutInSeconds:=30, GetPID:=True
  Exit Sub
ErrorHandler:
Assert:
  Dim ErrorNumber As Long
  ErrorNumber = Err.Number
  Phosphorus.AssertionsStatic.pAssert.Equal ExpectedErrorNumber, ErrorNumber, isCritical:=True
  If ErrorNumber <> ExpectedErrorNumber Then
    Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
  End If
  Resume Next
ErrorHandler1:
  ErrorNumber = Err.Number
  Phosphorus.AssertionsStatic.pAssert.Equal ExpectedErrorNumber, ErrorNumber, isCritical:=True
  If ErrorNumber <> ExpectedErrorNumber Then
    Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
  End If
End Sub

'Reusable commom method
Private Sub ValidPageLoadElementByBrowserTypeAndInstanceType(WebBrowserType As Phosphorus.pWebBrowserType, InstanceType As Phosphorus.pInstanceType)
'Check for no error if we open the URL in the current browser with a valid page load element - parameterised WebBrowserType and InstanceType
Arrange:
  Set pWindowsDriver = Phosphorus.Factory.GetNewPDriver(Phosphorus.pWindowsDriverType.WebBrowser)
  pWindowsDriver.SetWindowsDriverWebBrowserType _
    WebBrowserType:=WebBrowserType, _
    InstanceType:=InstanceType
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = 0
  On Error GoTo ErrorHandler
Act:
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

'Test all options for launching MSEDGE

'@TestMethod
Public Sub ValidPageLoadElement_Edge_ReuseACurrentOpenInstance()
'Check for no error if we open the URL in the current browser with a valid page load element - ReuseACurrentOpenInstance Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType Phosphorus.pWebBrowserType.MicrosoftEdge, Phosphorus.pInstanceType.ReuseACurrentOpenInstance
End Sub

'@TestMethod
Public Sub ValidPageLoadElement_Edge_Executable()
'Check for no error if we open the URL in the current browser with a valid page load element - Executable Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType Phosphorus.pWebBrowserType.MicrosoftEdge, Phosphorus.pInstanceType.Executable
End Sub

'@TestMethod
Public Sub ValidPageLoadElement_Edge_NewWindow()
'Check for no error if we open the URL in the current browser with a valid page load element - NewWindow Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType Phosphorus.pWebBrowserType.MicrosoftEdge, Phosphorus.pInstanceType.NewWindow
End Sub

'@TestMethod
Public Sub ValidPageLoadElement_Edge_AppMode()
'Check for no error if we open the URL in the current browser with a valid page load element - AppMode Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType Phosphorus.pWebBrowserType.MicrosoftEdge, Phosphorus.pInstanceType.AppMode
End Sub

'@TestMethod
Public Sub ValidPageLoadElement_Edge_NewProfile()
'Check for no error if we open the URL in the current browser with a valid page load element - NewProfile Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType Phosphorus.pWebBrowserType.MicrosoftEdge, Phosphorus.pInstanceType.NewProfile
End Sub

'@TestMethod
Public Sub ValidPageLoadElement_Edge_ApplicationUserModelID()
'Check for no error if we open the URL in the current browser with a valid page load element - ApplicationUserModelID Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType Phosphorus.pWebBrowserType.MicrosoftEdge, Phosphorus.pInstanceType.ApplicationUserModelID
End Sub

'Test other browsers
'@TestMethod
Public Sub ValidPageLoadElement_DuckDuckGo_ApplicationUserModelID() 'DuckDuckGo is only an app?
'Check for no error if we open the URL in the current browser with a valid page load element - ApplicationUserModelID Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType Phosphorus.pWebBrowserType.DuckDuckGo, Phosphorus.pInstanceType.ApplicationUserModelID
End Sub

'@TestMethod
Public Sub ValidPageLoadElement_Chrome_Executable()
'Check for no error if we open the URL in the current browser with a valid page load element - Executable Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType Phosphorus.pWebBrowserType.Chrome, Phosphorus.pInstanceType.Executable
End Sub

'@TestMethod
Public Sub ValidPageLoadElement_Chrome_NewProfile()
'Check for no error if we open the URL in the current browser with a valid page load element - NewProfile Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType Phosphorus.pWebBrowserType.Chrome, Phosphorus.pInstanceType.NewProfile
End Sub

'@TestMethod
Public Sub ValidPageLoadElement_Chrome_GuestModeNoSignIn()
'Check for no error if we open the URL in the current browser with a valid page load element - GuestModeNoSignIn Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType Phosphorus.pWebBrowserType.Chrome, Phosphorus.pInstanceType.GuestModeNoSignIn
End Sub

'@TestMethod
Public Sub ValidPageLoadElement_Firefox_Executable()
'Check for no error if we open the URL in the current browser with a valid page load element - Executable Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType Phosphorus.pWebBrowserType.Firefox, Phosphorus.pInstanceType.Executable
End Sub

'@TestMethod
Public Sub ValidPageLoadElement_Opera_Executable()
'Check for no error if we open the URL in the current browser with a valid page load element - Executable Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType Phosphorus.pWebBrowserType.Opera, Phosphorus.pInstanceType.Executable
End Sub

'@TestMethod
Public Sub ValidPageLoadElement_Brave_Executable()
'Check for no error if we open the URL in the current browser with a valid page load element - Executable Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType Phosphorus.pWebBrowserType.Brave, Phosphorus.pInstanceType.Executable
End Sub

'brave
'https://brave.com/
'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe
'TOR
'https://www.torproject.org/download/


Private Sub ValidPageLoadElementForApps(App As Phosphorus.WindowsApp)
'Check for no error if we open an app
Arrange:
  Set pWindowsDriver = Phosphorus.Factory.GetNewPDriver(Phosphorus.pWindowsDriverType.WindowsApp)
  pWindowsDriver.SetWindowsApp App
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = 0
  On Error GoTo ErrorHandler
Act:
  pWindowsDriver.Launch App.FriendlyName, App.OfficialName, TimeoutInSeconds:=30
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

'@TestMethod
Public Sub ValidPageLoadElement_MicrosoftWindowsCalculator()
  ValidPageLoadElementForApps Phosphorus.WindowsWindowsApps.MicrosoftWindowsCalculator
End Sub


