Attribute VB_Name = "WindowsDriver03LaunchWDs"
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
  Set pWindowsDriver = pWinDriver.pWindowsDriverStatic.GetNewPDriver(pWinDriver.pWindowsDriverType.WebBrowser)
  pWindowsDriver.SetWindowsDriverWebBrowserType WebBrowserType:=pWinDriver.pWebBrowserType.MicrosoftEdge
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
  Set pWindowsDriver = pWinDriver.pWindowsDriverStatic.GetNewPDriver(pWinDriver.pWindowsDriverType.WebBrowser)
  pWindowsDriver.SetWindowsDriverWebBrowserType WebBrowserType:=pWinDriver.pWebBrowserType.MicrosoftEdge
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
  Set pWindowsDriver = pWinDriver.pWindowsDriverStatic.GetNewPDriver(pWinDriver.pWindowsDriverType.WebBrowser)
  pWindowsDriver.SetWindowsDriverWebBrowserType WebBrowserType:=pWinDriver.pWebBrowserType.MicrosoftEdge
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
  Set pWindowsDriver = pWinDriver.pWindowsDriverStatic.GetNewPDriver(pWinDriver.pWindowsDriverType.WebBrowser)
  pWindowsDriver.SetWindowsDriverWebBrowserType WebBrowserType:=pWinDriver.pWebBrowserType.MicrosoftEdge, _
    InstanceType:=pWinDriver.pInstanceType.ReuseACurrentOpenInstance
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = 0
Act:
  'Open the browser with the valid page
  pWindowsDriver.Launch WEB_APP_NAME, TARGET_PAGE_TITLE, TARGET_PAGE_URL, TimeoutInSeconds:=30
Assert:
  Phosphorus.AssertionsStatic.pAssert.IsTrue pWindowsDriver.ElementExists("Header", "//Document[@AutomationId=""RootWebArea""]//Text[And(@AriaRole=""heading"",@Name=""Example Domain"")]"), isCritical:=True
  Exit Sub
ErrorHandler:
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
  Set pWindowsDriver = pWinDriver.pWindowsDriverStatic.GetNewPDriver(pWinDriver.pWindowsDriverType.WebBrowser)
  pWindowsDriver.SetWindowsDriverWebBrowserType WebBrowserType:=pWinDriver.pWebBrowserType.MicrosoftEdge
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = VBA.Constants.vbObjectError + Phosphorus.Exceptions.WindowsDriverUIElementNotFoundBeforeTimeout
Act:
  'Open the browser - Use a short timeout as is doesn't matter if the page hasn't fully loaded yet
  pWindowsDriver.Launch WEB_APP_NAME, "NOT " & TARGET_PAGE_TITLE, TARGET_PAGE_URL, TimeoutInSeconds:=5
  'Resumes here after Asser if all ok
  On Error GoTo ErrorHandler1
  ExpectedErrorNumber = 0
  'Use the valid page load element PPath to get the PID from the found element so that the browser will get killed on termination of the driver
  pWindowsDriver.FindElement "InvalidPageLoadElement", "/Window[xp:starts-with(@Name,""" & TARGET_PAGE_TITLE & """)]", TimeoutInSeconds:=30, GetPID:=True
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
Private Sub ValidPageLoadElementByBrowserTypeAndInstanceType(WebBrowserType As pWinDriver.pWebBrowserType, InstanceType As pWinDriver.pInstanceType)
'Check for no error if we open the URL in the current browser with a valid page load element - parameterised WebBrowserType and InstanceType
Arrange:
  Set pWindowsDriver = pWinDriver.pWindowsDriverStatic.GetNewPDriver(pWinDriver.pWindowsDriverType.WebBrowser)
  pWindowsDriver.SetWindowsDriverWebBrowserType _
    WebBrowserType:=WebBrowserType, _
    InstanceType:=InstanceType
  Dim ExpectedErrorNumber As Long
  ExpectedErrorNumber = 0
  On Error GoTo ErrorHandler
Act:
  'Open the browser with the valid page
  pWindowsDriver.Launch WEB_APP_NAME, TARGET_PAGE_TITLE, TARGET_PAGE_URL, TimeoutInSeconds:=30
  
  Dim FullBrowserRootWebAreaPPath As String
  FullBrowserRootWebAreaPPath = pWindowsDriver.GetWebBrowserFullBrowserRootWebAreaPPath()
    
  Dim CurrentPPath As String
    
  Dim HeaderNodeText As String
  Dim HeaderNodePPath As String
  If WebBrowserType = pWebBrowserType.Opera Then
    HeaderNodeText = ""
  Else
    HeaderNodeText = "Example Domain"
  End If
  HeaderNodePPath = pWindowsDriver.GetWebBrowserPPathConfigurationItem(pWinDriver.pWebBrowserPPathConfigurationItems.HeaderNodePPath, HeaderNodeText)
  CurrentPPath = FullBrowserRootWebAreaPPath & HeaderNodePPath
  If HeaderNodePPath <> "" Then
    Phosphorus.AssertionsStatic.pAssert.IsTrue pWindowsDriver.ElementExists("Header", CurrentPPath), "Check for 'Example Domain' Header Element", isCritical:=True
  End If
  
  Dim TextNodeText As String
  Dim TextNodePPath As String
  TextNodeText = "This domain is for use in documentation examples without needing permission. Avoid use in operations."
  TextNodePPath = pWindowsDriver.GetWebBrowserPPathConfigurationItem(pWinDriver.pWebBrowserPPathConfigurationItems.TextNodePPath, TextNodeText)
  CurrentPPath = FullBrowserRootWebAreaPPath & TextNodePPath
  If TextNodePPath <> "" Then
    Phosphorus.AssertionsStatic.pAssert.IsTrue pWindowsDriver.ElementExists("Text", CurrentPPath), "Check for 'Example Domain' Text Element", isCritical:=True
  End If

  Dim HyperlinkNodeText As String
  Dim HyperlinkNodePPath As String
  HyperlinkNodeText = "Learn more"
  HyperlinkNodePPath = pWindowsDriver.GetWebBrowserPPathConfigurationItem(pWinDriver.pWebBrowserPPathConfigurationItems.HyperlinkNodePPath, HyperlinkNodeText)
  CurrentPPath = FullBrowserRootWebAreaPPath & HyperlinkNodePPath
  If HyperlinkNodePPath <> "" Then
    Phosphorus.AssertionsStatic.pAssert.IsTrue pWindowsDriver.ElementExists("Hyperlink", CurrentPPath), "Check for 'Example Domain' Hypelink Element", isCritical:=True
  End If

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

'**************************************
'Test all options for launching MSEDGE
'**************************************

'@WebBrowserLaunch
'@TestMethod
Public Sub ValidPageLoadElement_Edge_ReuseACurrentOpenInstance()
'Check for no error if we open the URL in the current browser with a valid page load element - ReuseACurrentOpenInstance Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType pWinDriver.pWebBrowserType.MicrosoftEdge, pWinDriver.pInstanceType.ReuseACurrentOpenInstance
End Sub

'@WebBrowserLaunch
'@TestMethod
Public Sub ValidPageLoadElement_Edge_Executable()
'Check for no error if we open the URL in the current browser with a valid page load element - Executable Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType pWinDriver.pWebBrowserType.MicrosoftEdge, pWinDriver.pInstanceType.Executable
End Sub

'@WebBrowserLaunch
'@TestMethod
Public Sub ValidPageLoadElement_Edge_NewWindow()
'Check for no error if we open the URL in the current browser with a valid page load element - NewWindow Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType pWinDriver.pWebBrowserType.MicrosoftEdge, pWinDriver.pInstanceType.NewWindow
End Sub

'@WebBrowserLaunch
'@TestMethod
Public Sub ValidPageLoadElement_Edge_AppMode()
'Check for no error if we open the URL in the current browser with a valid page load element - AppMode Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType pWinDriver.pWebBrowserType.MicrosoftEdge, pWinDriver.pInstanceType.AppMode
End Sub

'@WebBrowserLaunch
'@TestMethod
Public Sub ValidPageLoadElement_Edge_NewProfile()
'Check for no error if we open the URL in the current browser with a valid page load element - NewProfile Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType pWinDriver.pWebBrowserType.MicrosoftEdge, pWinDriver.pInstanceType.NewProfile
End Sub

'@WebBrowserLaunch
'@TestMethod
Public Sub ValidPageLoadElement_Edge_ApplicationUserModelID()
'Check for no error if we open the URL in the current browser with a valid page load element - ApplicationUserModelID Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType pWinDriver.pWebBrowserType.MicrosoftEdge, pWinDriver.pInstanceType.ApplicationUserModelID
End Sub

'**************************************
'Test other browsers
'**************************************

'@WebBrowserLaunch
'@TestMethod
Public Sub ValidPageLoadElement_DuckDuckGo_ApplicationUserModelID() 'DuckDuckGo is only an app?
'Check for no error if we open the URL in the current browser with a valid page load element - ApplicationUserModelID Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType pWinDriver.pWebBrowserType.DuckDuckGo, pWinDriver.pInstanceType.ApplicationUserModelID
End Sub

'@WebBrowserLaunch
'@TestMethod
Public Sub ValidPageLoadElement_Chrome_Executable()
'Check for no error if we open the URL in the current browser with a valid page load element - Executable Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType pWinDriver.pWebBrowserType.Chrome, pWinDriver.pInstanceType.Executable
End Sub

'@WebBrowserLaunch
'@TestMethod
Public Sub ValidPageLoadElement_Chrome_NewProfile()
'Check for no error if we open the URL in the current browser with a valid page load element - NewProfile Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType pWinDriver.pWebBrowserType.Chrome, pWinDriver.pInstanceType.NewProfile
End Sub

'@WebBrowserLaunch
'@TestMethod
Public Sub ValidPageLoadElement_Chrome_GuestModeNoSignIn()
'Check for no error if we open the URL in the current browser with a valid page load element - GuestModeNoSignIn Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType pWinDriver.pWebBrowserType.Chrome, pWinDriver.pInstanceType.GuestModeNoSignIn
End Sub

'@WebBrowserLaunch
'@TestMethod
Public Sub ValidPageLoadElement_Firefox_Executable()
'Check for no error if we open the URL in the current browser with a valid page load element - Executable Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType pWinDriver.pWebBrowserType.Firefox, pWinDriver.pInstanceType.Executable
End Sub

'@WebBrowserLaunch
'@TestMethod
Public Sub ValidPageLoadElement_Opera_Executable()
'Check for no error if we open the URL in the current browser with a valid page load element - Executable Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType pWinDriver.pWebBrowserType.Opera, pWinDriver.pInstanceType.Executable
End Sub

'@WebBrowserLaunch
'@TestMethod
Public Sub ValidPageLoadElement_Brave_Executable()
'Check for no error if we open the URL in the current browser with a valid page load element - Executable Instance Type
  ValidPageLoadElementByBrowserTypeAndInstanceType pWinDriver.pWebBrowserType.Brave, pWinDriver.pInstanceType.Executable
End Sub

'brave
'https://brave.com/
'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe
'TOR
'https://www.torproject.org/download/

Private Sub ValidPageLoadElementForApps(App As Phosphorus.WindowsApp)
'Check for no error if we open an app
Arrange:
  Set pWindowsDriver = pWinDriver.pWindowsDriverStatic.GetNewPDriver(pWinDriver.pWindowsDriverType.WindowsApp)
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


