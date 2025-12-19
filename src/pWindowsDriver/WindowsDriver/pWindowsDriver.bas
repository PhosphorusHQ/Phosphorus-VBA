VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pWindowsDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder WindowsDriver
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Private Const DEFAULT_IMPLICIT_TIMEOUT_IN_SECONDS = 30
Private Const DEFAULT_IMPLICIT_DELAY_BETWEEN_POLLS_IN_MILLISECOND = 10
 
Private This As DriverProperties
Private Type DriverProperties
  DriverType As pWinDriver.pWindowsDriverType
  WebBrowserType As pWinDriver.pWindowsDriverWebBrowserType
  TempDirectoryForCurrentAppInstance As String
  WindowsApp As Phosphorus.WindowsApp
  SubDriver As Object
  Navigate As pWindowsDriver_Navigate
  WebAppTitle As String
  PageLoadedElementPPath As String 'PPath of an element that indicates the driver has at least partially loaded
  PageLoadedElementExpectedWindowInteractionState As UIAutomationClient.WindowInteractionState
  PageUnloadedElementPPath As String 'PPath of an element that indicates the driver has now been closed (by it's non-existence)
  MasterWindowsDriverElement As pWinDriver.pWindowsDriverElement
  BrowserRootViewElement As pWinDriver.pWindowsDriverElement
  BrowserRootWebAreaElement As pWinDriver.pWindowsDriverElement
  ProcessId As Long
  ImplicitTimeoutInSeconds As Integer
  CurrentEvaluatedPPath As pPath.ReturnClass
'  CurrentUIAElement As UIAutomationClient.IUIAutomationElement
'NOT USED YET  ImplicitDelayBetweenPollsInMilliseconds As Integer
End Type

Private Sub Class_Initialize()
  This.ImplicitTimeoutInSeconds = Phosphorus.Configuration.GetValue("pDriver", "ImplicitTimeout", DEFAULT_IMPLICIT_TIMEOUT_IN_SECONDS)
End Sub

Public Sub Terminate()
  'Close the Driver cleanly
  This.MasterWindowsDriverElement.CloseWindow
  'Release the MasterWindowsDriverElement
  Set This.MasterWindowsDriverElement = Nothing
  'Try to kill the process in case the clean close didn't work
  If This.ProcessId <> 0 Then
    Phosphorus.WindowsProcesses.KillProcessByID This.ProcessId
  End If
  If This.PageUnloadedElementPPath <> "" Then
    Me.WaitUntilElementNotExists "PageUnloadedElement", This.PageUnloadedElementPPath
  End If
  DeleteAnyTempDirectory
End Sub

Public Sub SetDriverType(ByVal DriverType As pWinDriver.pWindowsDriverType)
  This.DriverType = DriverType
  If This.DriverType = pWinDriver.pWindowsDriverType.WebBrowser Then
    Set This.WebBrowserType = New pWinDriver.pWindowsDriverWebBrowserType
  End If
End Sub

Public Function GetDriverType() As pWinDriver.pWindowsDriverType
  GetDriverType = This.DriverType
End Function

Public Sub SetWindowsDriverWebBrowserType(ByVal WebBrowserType As pWinDriver.pWebBrowserType, Optional ByVal InstanceType As pWinDriver.pInstanceType)
  This.WebBrowserType.WebBrowserType = WebBrowserType
  If InstanceType > 0 Then
    This.WebBrowserType.InstanceType = InstanceType
  End If
End Sub

Public Function GetWindowsDriverWebBrowserType() As pWinDriver.pWindowsDriverWebBrowserType
  Set GetWindowsDriverWebBrowserType = This.WebBrowserType
End Function

Public Sub SetWindowsApp(WindowsApp As Phosphorus.WindowsApp)
  This.WindowsApp = WindowsApp
End Sub

Public Function GetWindowsApp() As Phosphorus.WindowsApp
  GetWindowsApp = This.WindowsApp
End Function

Public Function GetDefaultImplicitTimeoutInSeconds() As Integer
  GetDefaultImplicitTimeoutInSeconds = DEFAULT_IMPLICIT_TIMEOUT_IN_SECONDS
End Function

Public Sub SetPageLoadedElement( _
  PageLoadedElementPPath As String, _
  ExpectedWindowInteractionState As UIAutomationClient.WindowInteractionState)
  This.PageLoadedElementPPath = PageLoadedElementPPath
  This.PageLoadedElementExpectedWindowInteractionState = ExpectedWindowInteractionState
End Sub

Public Function GetPageLoadedElementPPath() As String
  GetPageLoadedElementPPath = This.PageLoadedElementPPath
End Function

Public Sub SetPageUnloadedElementPPath(ByVal PageUnloadedElementPPath As String)
  This.PageUnloadedElementPPath = This.PageLoadedElementPPath
End Sub

Public Function CreateTempDirectory() As String
  'Create a unique temp directory for app instance isolation
  This.TempDirectoryForCurrentAppInstance = Environ("TEMP") & "\PhosphorusTemp" & Format(Now, "yyyymmddhhmmss")
  VBA.FileSystem.MkDir This.TempDirectoryForCurrentAppInstance
  CreateTempDirectory = This.TempDirectoryForCurrentAppInstance
End Function

Private Function DeleteAnyTempDirectory() As String
  If This.TempDirectoryForCurrentAppInstance <> "" Then
    'We need to close or kill the browser before we can delete it!"
    'FileSystemObject requires a reference to the Microsoft Scripting Runtime Library (Scrrun.dll)
    Dim MyFSO As FileSystemObject
    Set MyFSO = New FileSystemObject
    MyFSO.DeleteFolder This.TempDirectoryForCurrentAppInstance
    Set MyFSO = Nothing
  End If
End Function

Public Sub Launch( _
  ByVal AppName As String, _
  ByVal WebAppTitle As String, _
  Optional ByVal Document As String, _
  Optional ByVal TimeoutInSeconds As Integer, _
  Optional ByVal CheckHTTPStatusCodeOnly As Boolean, _
  Optional ByVal WindowShowState As Phosphorus.WindowShowStates = Phosphorus.WindowShowStates.SW_SHOWMAXIMIZED)
  
  This.WebAppTitle = WebAppTitle
  
  'Check the page status?
  If VBA.Strings.Left(Document, 4) = "http" Then
    CheckHTTPStatusCode Document
  End If
  
  If Not CheckHTTPStatusCodeOnly Then
  
    'Initialise any subdriver
    Select Case This.DriverType
      
      Case pWinDriver.pWindowsDriverType.PreLaunched
        'The App is already lauched, so we can ignore this type - it should really be called anyway
      
      Case pWinDriver.pWindowsDriverType.WebBrowser
        Select Case This.WebBrowserType.WebBrowserType
          Case pWinDriver.pWebBrowserType.MicrosoftEdge
            Set This.SubDriver = New pWinDriver.pWindowsDriverMSEdge
          Case pWinDriver.pWebBrowserType.DuckDuckGo
            Set This.SubDriver = New pWinDriver.pWindowsDriverDuckDuckGo
          Case pWinDriver.pWebBrowserType.Chrome
            Set This.SubDriver = New pWinDriver.pWindowsDriverChrome
          Case pWinDriver.pWebBrowserType.Firefox
            Set This.SubDriver = New pWinDriver.pWindowsDriverFirefox
          Case pWinDriver.pWebBrowserType.Opera
            Set This.SubDriver = New pWinDriver.pWindowsDriverOpera
          Case pWinDriver.pWebBrowserType.Brave
            Set This.SubDriver = New pWinDriver.pWindowsDriverBrave
          Case Else
            Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverUnhandledWebBrowserType, "Web Brower Type: #" & This.WebBrowserType.WebBrowserType
         End Select
         If Not This.SubDriver Is Nothing Then
           This.SubDriver.IWindowsDriverWebBrowser_LaunchApp Me, AppName, WebAppTitle, Document
         End If
         
      Case pWinDriver.pWindowsDriverType.WindowsApp
        Phosphorus.WindowsProcesses.LaunchAppByAUMID This.WindowsApp
        This.PageLoadedElementPPath = This.WindowsApp.PageLoadedElementPPath
        This.PageLoadedElementExpectedWindowInteractionState = UIAutomationClient.WindowInteractionState.WindowInteractionState_ReadyForUserInteraction
      
      Case Else
        Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverUnhandledDriverType, "Driver Type: #" & This.DriverType
        
    End Select

    If This.PageLoadedElementPPath = "" Then
    
      Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverNoPageLoadElementDefined, AppName, This.WebAppTitle
    
    Else

      If This.DriverType = pWinDriver.pWindowsDriverType.WebBrowser Then
        This.SubDriver.IWindowsDriverWebBrowser_PostLaunchApp
      End If
      
      'Wait for the Page Load Element
      Set This.MasterWindowsDriverElement = FindElement("PageLoadedElement", This.PageLoadedElementPPath, TimeoutInSeconds)
      'Store this as the Master Window element so that we can easily close the current driver window when it is finished with
      This.MasterWindowsDriverElement.Name = "MasterWindowsDriverElement"
      Logger.InternalInfo "MasterWindowsDriverElement:=" & This.MasterWindowsDriverElement.FoundBypPath

      'Set the default element PPath which indicate the driver has been unloaded/closed
      Me.SetPageUnloadedElementPPath GetPageLoadedElementPPath
   
      'Wait until it is in the right WindowInteractionState
'TODO: This only applies to (top level?) Window controls?
      This.MasterWindowsDriverElement.WaitForWindowInteractionState This.PageLoadedElementExpectedWindowInteractionState

      'Get the App's PID from the MasterWindowsDriverElement
      This.ProcessId = This.MasterWindowsDriverElement.GetProcessID
      
      'Check for the Browser Root elements if this is a Web Browser
      If This.DriverType = WebBrowser Then
        CheckForBrowserRootElements
        If This.WebBrowserType.WebBrowserType = Opera Then
          CloseAllOtherTabs
        End If
      End If
      
    End If

'TODO: How to open a new tab for 2nd locator, etc?
  
  End If
End Sub

Private Sub CheckHTTPStatusCode(ByVal lstrUrl As String)
  ' Send Request
  Dim lRequest As WinHttpRequest
  Set lRequest = New WinHttpRequest
  lRequest.Open "GET", lstrUrl
  On Error Resume Next
  lRequest.Send
  'Check for errors
  Dim ErrNum As Long
  Dim ErrDescription As String
  ErrNum = Err.Number
  ErrDescription = Err.Description
  On Error GoTo 0
  If ErrNum <> 0 Then
    If ErrNum = -2147012889 Then
      Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverInvalidURL, lstrUrl, ErrDescription
    Else
      Phosphorus.pExceptions.Raise Phosphorus.Exceptions.UnexpectedError, ErrNum, ErrDescription, "Phosphorus.pWindowsDriver.CheckHTTPStatusCode"
    End If
  Else
    'Check the status code received
    If lRequest.Status <> 200 Then
      Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverInvalidHTTPStatus, lstrUrl, lRequest.Status
    End If
  End If
End Sub

Private Sub CheckForBrowserRootElements()

  Dim CurrentpPath As String
  Dim BrowserRootViewpPath As String
  Dim BrowserRootWebAreapPath As String
  BrowserRootViewpPath = GetWebBrowserRootViewpPath
  CurrentpPath = BrowserRootViewpPath
  If BrowserRootViewpPath <> "" Then
    Set This.BrowserRootViewElement = RefreshPageUntilBrowserRootElementExists("BrowserRootWebAreaPPath", CurrentpPath, RootElement:=This.MasterWindowsDriverElement)
    If This.BrowserRootViewElement Is Nothing Then
      Logger.ExternalFatal "Browser Root View Element not found!"
      Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverWebBrowserRootViewElementNotFound, CurrentpPath
    End If
  End If
  Logger.InternalInfo "BrowserRootViewpPath:=" & BrowserRootViewpPath

  BrowserRootWebAreapPath = GetWebBrowserRootWebAreapPath
  CurrentpPath = BrowserRootWebAreapPath
  If BrowserRootWebAreapPath <> "" Then
    Set This.BrowserRootWebAreaElement = RefreshPageUntilBrowserRootElementExists("BrowserRootWebAreaPPath", CurrentpPath, RootElement:=This.BrowserRootViewElement)
      If This.BrowserRootWebAreaElement Is Nothing Then
        Logger.ExternalFatal "Browser Root Web Area not found!"
        Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverWebBrowserRootWebAreaElementNotFound, CurrentpPath
      End If
  Else
    Logger.InternalTrace "Opera uses the same element for these two, so don't waste time searching for it again"
    Set This.BrowserRootWebAreaElement = This.BrowserRootViewElement
  End If
  Logger.InternalInfo "BrowserRootWebAreapPath:=" & BrowserRootWebAreapPath

End Sub

Private Function RefreshPageUntilBrowserRootElementExists(Name As String, pPath As String, RootElement As pWinDriver.pWindowsDriverElement) As pWinDriver.pWindowsDriverElement
  Dim i As Integer
  Dim Continue As Boolean
  Dim ReturnElement As pWinDriver.pWindowsDriverElement
  i = 1
  Continue = True
  While Continue
    Set ReturnElement = FindElement(Name, pPath, TimeoutInSeconds:=0, RootElement:=RootElement, CheckExistenceOnly:=True)
    If ReturnElement Is Nothing And i <= 10 Then
      Navigate.Refresh
      i = i + 1
    Else
      Continue = False
    End If
  Wend
  Set RefreshPageUntilBrowserRootElementExists = ReturnElement
End Function

Public Function GetMasterWindowsDriverElement() As pWinDriver.pWindowsDriverElement
  Set GetMasterWindowsDriverElement = This.MasterWindowsDriverElement
End Function

Public Function GetWebBrowserRootViewElement() As pWinDriver.pWindowsDriverElement
  Set GetWebBrowserRootViewElement = This.BrowserRootViewElement
End Function

Public Function GetWebBrowserRootWebAreaElement() As pWinDriver.pWindowsDriverElement
  Set GetWebBrowserRootWebAreaElement = This.BrowserRootWebAreaElement
End Function

Public Function GetWebBrowserFullBrowserRootWebAreaPPath() As String
  GetWebBrowserFullBrowserRootWebAreaPPath = GetWebBrowserRootViewpPath & GetWebBrowserRootWebAreapPath
End Function

Private Function GetWebBrowserRootViewpPath()
  Dim BrowserRootViewControlType As String
  Dim BrowserRootViewClassName As String
  Dim BrowserRootViewUseWebAppTitleAsName As Boolean
  BrowserRootViewControlType = GetWebBrowserpPathConfigurationItem(pWinDriver.pWebBrowserPPathConfigurationItems.BrowserRootViewControlType)
  BrowserRootViewClassName = GetWebBrowserpPathConfigurationItem(pWinDriver.pWebBrowserPPathConfigurationItems.BrowserRootViewClassName)
  BrowserRootViewUseWebAppTitleAsName = GetWebBrowserpPathConfigurationItem(pWinDriver.pWebBrowserPPathConfigurationItems.BrowserRootViewUseWebAppTitleAsName)
  If BrowserRootViewControlType <> "" Then
    GetWebBrowserRootViewpPath = "//" & BrowserRootViewControlType
    If BrowserRootViewClassName <> "" Then
      'NB: Opera needs the Name attribute here!
      Logger.InternalTrace "Opera needs the Name attribute here!"
      If BrowserRootViewUseWebAppTitleAsName Then
        GetWebBrowserRootViewpPath = GetWebBrowserRootViewpPath & _
          "[And(@ClassName=""" & BrowserRootViewClassName & """,@Name=""" & This.WebAppTitle & """)]"
      Else
        GetWebBrowserRootViewpPath = GetWebBrowserRootViewpPath & "[@ClassName=""" & BrowserRootViewClassName & """]"
      End If
    Else
      If BrowserRootViewUseWebAppTitleAsName Then
        GetWebBrowserRootViewpPath = GetWebBrowserRootViewpPath & "[@Name=""" & This.WebAppTitle & """]"
      End If
    End If
  End If
  Logger.InternalTrace "Got Web Browser Root View pPath as: " & GetWebBrowserRootViewpPath
End Function

Private Function GetWebBrowserRootWebAreapPath()
  Dim BrowserRootWebAreaControlType As String
  Dim BrowserRootWebAreaAutomationId As String
  BrowserRootWebAreaControlType = GetWebBrowserpPathConfigurationItem(pWinDriver.pWebBrowserPPathConfigurationItems.RootWebAreaControlType)
  BrowserRootWebAreaAutomationId = GetWebBrowserpPathConfigurationItem(pWinDriver.pWebBrowserPPathConfigurationItems.RootWebAreaAutomationID)
  If BrowserRootWebAreaControlType <> "" Then
    GetWebBrowserRootWebAreapPath = "//" & BrowserRootWebAreaControlType
    If BrowserRootWebAreaAutomationId <> "" And This.WebAppTitle <> "" Then
      GetWebBrowserRootWebAreapPath = GetWebBrowserRootWebAreapPath & _
        "[And(@AutomationId=""" & BrowserRootWebAreaAutomationId & """,@Name=""" & This.WebAppTitle & """)]"
    ElseIf This.WebAppTitle <> "" Then
      GetWebBrowserRootWebAreapPath = GetWebBrowserRootWebAreapPath & "[@Name=""" & This.WebAppTitle & """]"
    End If
  End If
  Logger.InternalTrace "Got Web Browser Root Area pPath as: " & GetWebBrowserRootWebAreapPath
End Function

Public Function GetWebBrowserpPathConfigurationItem(ItemType As pWinDriver.pWebBrowserPPathConfigurationItems, Optional Parameter1 As Variant) As Variant
  
  Select Case ItemType
    
    Case pWinDriver.pWebBrowserPPathConfigurationItems.HeaderNodePPath
      GetWebBrowserpPathConfigurationItem = GetWebBrowserpPathConfigurationTextItemPPath(pWinDriver.pWebBrowserPPathConfigurationItems.HeaderNodeAriaRole, VBA.Conversion.CStr(Parameter1))
    
    Case pWinDriver.pWebBrowserPPathConfigurationItems.TextNodePPath
      GetWebBrowserpPathConfigurationItem = GetWebBrowserpPathConfigurationTextItemPPath(pWinDriver.pWebBrowserPPathConfigurationItems.TextNodeAriaRole, VBA.Conversion.CStr(Parameter1))
    
    Case pWinDriver.pWebBrowserPPathConfigurationItems.HyperlinkNodePPath
      GetWebBrowserpPathConfigurationItem = GetWebBrowserpPathConfigurationTextItemPPath(pWinDriver.pWebBrowserPPathConfigurationItems.HyperlinkNodeAriaRole, VBA.Conversion.CStr(Parameter1))
    
    Case Else
      GetWebBrowserpPathConfigurationItem = This.SubDriver.IWindowsDriverWebBrowser_GetPPathConfigurationItem(ItemType)
  
  End Select

End Function

Private Function GetWebBrowserpPathConfigurationTextItemPPath(ItemType As pWinDriver.pWebBrowserPPathConfigurationItems, TextNodeText As String) As String
  
  Dim TextNodeControlType As String
  If ItemType = HyperlinkNodeAriaRole Then
    TextNodeControlType = "Hyperlink"
  Else
    If This.WebBrowserType.WebBrowserType = pWinDriver.pWebBrowserType.Opera Then
      TextNodeControlType = "*"
    Else
      TextNodeControlType = "Text"
    End If
  End If
  
  Dim TextNodeAriaRole As String
  TextNodeAriaRole = GetWebBrowserpPathConfigurationItem(ItemType)
  
  Dim TextNodePPath As String
  TextNodePPath = "//" & TextNodeControlType
  If TextNodeText = "" Then
    'We have to specify either text and maybe a role or both
    TextNodePPath = ""
  ElseIf TextNodeAriaRole = "" Then
    TextNodePPath = TextNodePPath & "[@Name=""" & TextNodeText & """]"
  Else
    TextNodePPath = TextNodePPath & "[And(@AriaRole=""" & TextNodeAriaRole & """,@Name=""" & TextNodeText & """)]"
  End If
  
  GetWebBrowserpPathConfigurationTextItemPPath = TextNodePPath

End Function

Public Function FindElement( _
  ByVal Name As String, _
  ByVal pPathString As String, _
  Optional ByVal TimeoutInSeconds As Integer, _
  Optional ByVal GetPID As Boolean, _
  Optional ByVal CheckExistenceOnly As Boolean = False, _
  Optional ByRef RootElement As pWinDriver.pWindowsDriverElement) As pWindowsDriverElement
'  Optional ByRef RootElement As UIAutomationClient.IUIAutomationElement) As pWindowsDriverElement
   
  Dim CurrentpPath As pPath.Core
  Set CurrentpPath = Nothing
  Set This.CurrentEvaluatedPPath = Nothing
  
  Set CurrentpPath = pPath.ConstantsAndStatic.GetNewPhosphorusPPath
  CurrentpPath.Initialise

  'We should have released the MasterWindowsDriverElement when it gets closed!
  Dim InitialpPath As String
  If Not (RootElement Is Nothing) Then
    InitialpPath = RootElement.FullFoundBypPath
    CurrentpPath.SetApplicationRootElement RootElement.GetUIAElement
  ElseIf Not (This.BrowserRootWebAreaElement Is Nothing) Then
    InitialpPath = This.BrowserRootWebAreaElement.FullFoundBypPath
    CurrentpPath.SetApplicationRootElement This.BrowserRootWebAreaElement.GetUIAElement
  ElseIf Not (This.MasterWindowsDriverElement Is Nothing) Then
    InitialpPath = This.MasterWindowsDriverElement.FullFoundBypPath
    CurrentpPath.SetApplicationRootElement This.MasterWindowsDriverElement.GetUIAElement
  Else
    InitialpPath = "{Desktop}"
    CurrentpPath.SetApplicationRootElement pWinDriver.pWindowsDriverStatic.gUIADesktopUIElement
  End If
  Logger.InternalDebug "Root Element pPath: " & InitialpPath

  'Get default timeout if none set (it might be 0!)
  If IsMissing(TimeoutInSeconds) Then
    TimeoutInSeconds = This.ImplicitTimeoutInSeconds
  End If
  
  'Calculate the end time
  Dim dtEndTime As Date
  dtEndTime = DateAdd("s", TimeoutInSeconds, Now)

  'Loop until element(s) found or timed out
  Dim boolElementFound As Boolean
  boolElementFound = False
  Dim boolPassedEndTime As Boolean
  boolPassedEndTime = False
  Dim i As Integer
  i = 0
  While Not (boolElementFound Or boolPassedEndTime)
    Set This.CurrentEvaluatedPPath = CurrentpPath.Evaluate(FullLocationPathExpression:=pPathString, InitialpPath:=InitialpPath)
    boolElementFound = This.CurrentEvaluatedPPath.ReturnedValue
    If Not boolElementFound Then
      boolPassedEndTime = (Now > dtEndTime)
      i = i + 1
      Logger.ExternalDebug "Sleeping " & i & ": " & DEFAULT_IMPLICIT_DELAY_BETWEEN_POLLS_IN_MILLISECOND
      Phosphorus.WindowsProcesses.Snooze DEFAULT_IMPLICIT_DELAY_BETWEEN_POLLS_IN_MILLISECOND
    End If
  Wend

  If Not boolElementFound And Not CheckExistenceOnly Then
    Logger.ExternalFatal "Element named '" & Name & "' not found by pPath: " & pPathString
    Phosphorus.pExceptions.Raise _
      Phosphorus.Exceptions.WindowsDriverUIElementNotFoundBeforeTimeout, _
      Name, _
      VBA.Conversion.CStr(TimeoutInSeconds), _
      pPathString
  End If
    
  If boolElementFound Then
    
    Logger.InternalDebug "Element has been found"
    
    'TODO: Check for only 1 matching element!
    
    'Store the current UIA element
    Dim FoundUIAElement As UIAutomationClient.IUIAutomationElement
    Set FoundUIAElement = This.CurrentEvaluatedPPath.GetMatchingElement(1)
    Dim NewWindowsDriverElement As pWindowsDriverElement
    Set NewWindowsDriverElement = New pWindowsDriverElement
    NewWindowsDriverElement.Initialise Name, Me, FoundUIAElement, pPathString, InitialpPath
    Set CurrentpPath = Nothing

    'Prepare to Kill PID at end of session! - Needed for test cases and error handling
    If GetPID Then
      This.ProcessId = NewWindowsDriverElement.GetProcessID
    End If
       
  End If
  
  'Return the found window element, if any
  Set FindElement = NewWindowsDriverElement

End Function

Public Function ElementExists( _
  ByVal ElementName As String, _
  ByVal pPathString As String, _
  Optional ByRef RootElement As UIAutomationClient.IUIAutomationElement, _
  Optional TimeoutInSeconds As Integer = 0) As Boolean
  
  Dim FoundElement As pWindowsDriverElement
  Set FoundElement = Me.FindElement(Name:=ElementName, pPathString:=pPathString, TimeoutInSeconds:=TimeoutInSeconds, CheckExistenceOnly:=True, RootElement:=RootElement)
  ElementExists = Not (FoundElement Is Nothing)

End Function

Public Sub WaitUntilElementNotExists( _
  ByVal ElementName, _
  ByVal pPathString As String, _
  Optional ByVal TimeoutInSeconds As Integer)
  
  If IsMissing(TimeoutInSeconds) Then
    TimeoutInSeconds = This.ImplicitTimeoutInSeconds
  End If
  If Not ElementExists(ElementName, pPathString) Then
    Exit Sub
  Else
    Stop
    'TODO: Need to loop and wait until the element doesn't exist here, or we reach a timeout
  End If

End Sub

Public Sub CloseAllOtherTabs()
   This.SubDriver.IWindowsDriverWebBrowser_CloseAllOtherTabs
End Sub

Public Function Navigate() As pWindowsDriver_Navigate
  If This.Navigate Is Nothing Then
    Set This.Navigate = New pWindowsDriver_Navigate
    This.Navigate.Initialise Me, This.SubDriver
    Set Navigate = This.Navigate
  Else
    Set Navigate = This.Navigate
  End If
End Function
