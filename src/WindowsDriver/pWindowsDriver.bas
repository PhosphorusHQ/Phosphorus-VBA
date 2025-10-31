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
Option Explicit

Private This As DriverProperties

Private Const DEFAULT_IMPLICIT_TIMEOUT_IN_SECONDS = 30
Private Const DEFAULT_IMPLICIT_DELAY_BETWEEN_POLLS_IN_MILLISECOND = 10
 
Private Type DriverProperties
  DriverType As Phosphorus.pWindowsDriverType
  WebBrowserType As Phosphorus.pWindowsDriverWebBrowserType
  TempDirectoryForCurrentAppInstance As String
  WindowsApp As Phosphorus.WindowsApp
  SubDriver As Object
  PageLoadedElementPPath As String 'PPath of an element that indicates the driver has at least partially loaded
  PageLoadedElementExpectedWindowInteractionState As UIAutomationClient.WindowInteractionState
  PageUnloadedElementPPath As String 'PPath of an element that indicates the driver has now been closed (by it's non-existence)
'  MasterWindowUIAElement As UIAutomationClient.IUIAutomationElement
  MasterWindowsDriverElement As Phosphorus.pWindowsDriverElement
  ProcessId As Long
  ImplicitTimeoutInSeconds As Integer
  CurrentEvaluatedPPath As Phosphorus.PPathReturnClass
'  CurrentUIAElement As UIAutomationClient.IUIAutomationElement
'NOT USED YET  ImplicitDelayBetweenPollsInMilliseconds As Integer
End Type

Private Sub Class_Initialize()
  This.ImplicitTimeoutInSeconds = Phosphorus.Configuration.GetValue("pDriver", "ImplicitTimeout", DEFAULT_IMPLICIT_TIMEOUT_IN_SECONDS)
End Sub

Public Sub Terminate()
  'Close the Driver cleanly
  This.MasterWindowsDriverElement.CloseWindow
  'Try to kill the process in case the clean close didn't work
  If This.ProcessId <> 0 Then
    Phosphorus.WindowsProcesses.KillProcessByID This.ProcessId
  End If
  If This.PageUnloadedElementPPath <> "" Then
    Me.WaitUntilElementNotExists "PageUnloadedElement", This.PageUnloadedElementPPath
  End If
  DeleteAnyTempDirectory
End Sub

Public Sub SetDriverType(ByVal DriverType As Phosphorus.pWindowsDriverType)
  This.DriverType = DriverType
  If This.DriverType = Phosphorus.pWindowsDriverType.WebBrowser Then
    Set This.WebBrowserType = New Phosphorus.pWindowsDriverWebBrowserType
  End If
End Sub

Public Function GetDriverType() As Phosphorus.pWindowsDriverType
  GetDriverType = This.DriverType
End Function

Public Sub SetWindowsDriverWebBrowserType(ByVal WebBrowserType As Phosphorus.pWebBrowserType, Optional ByVal InstanceType As Phosphorus.pInstanceType)
  This.WebBrowserType.WebBrowserType = WebBrowserType
  If InstanceType > 0 Then
    This.WebBrowserType.InstanceType = InstanceType
  End If
End Sub

Public Function GetWindowsDriverWebBrowserType() As Phosphorus.pWindowsDriverWebBrowserType
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
  ByVal AppTitle As String, _
  Optional ByVal Document As String, _
  Optional ByVal TimeoutInSeconds As Integer, _
  Optional ByVal CheckHTTPStatusCodeOnly As Boolean, _
  Optional ByVal WindowShowState As Phosphorus.WindowShowStates = Phosphorus.WindowShowStates.SW_SHOWMAXIMIZED)
  
  'Check the page status?
  If VBA.Strings.Left(Document, 4) = "http" Then
    CheckHTTPStatusCode Document
  End If
  
  If Not CheckHTTPStatusCodeOnly Then
  
    'Initialise any subdriver
    Select Case This.DriverType
      
      Case Phosphorus.pWindowsDriverType.PreLaunched
        'The App is already lauched, so we can ignore this type - it should really be called anyway
      
      Case Phosphorus.pWindowsDriverType.WebBrowser
        Select Case This.WebBrowserType.WebBrowserType
          Case Phosphorus.pWebBrowserType.MicrosoftEdge
            Set This.SubDriver = New Phosphorus.pWindowsDriverMSEdge
          Case Phosphorus.pWebBrowserType.DuckDuckGo
            Set This.SubDriver = New Phosphorus.pWindowsDriverDuckDuckGo
          Case Phosphorus.pWebBrowserType.Chrome
            Set This.SubDriver = New Phosphorus.pWindowsDriverChrome
          Case Phosphorus.pWebBrowserType.Firefox
            Set This.SubDriver = New Phosphorus.pWindowsDriverFirefox
          Case Phosphorus.pWebBrowserType.Opera
            Set This.SubDriver = New Phosphorus.pWindowsDriverOpera
          Case Phosphorus.pWebBrowserType.Brave
            Set This.SubDriver = New Phosphorus.pWindowsDriverBrave
          Case Else
            Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverUnhandledWebBrowserType, "Web Brower Type: #" & This.WebBrowserType.WebBrowserType
         End Select
         If Not This.SubDriver Is Nothing Then
           This.SubDriver.IWindowsDriverWebBrowser_LaunchApp Me, AppName, AppTitle, Document
         End If
         
      Case Phosphorus.pWindowsDriverType.WindowsApp
        Phosphorus.WindowsProcesses.LaunchAppByAUMID This.WindowsApp
        This.PageLoadedElementPPath = This.WindowsApp.PageLoadedElementPPath
        This.PageLoadedElementExpectedWindowInteractionState = UIAutomationClient.WindowInteractionState.WindowInteractionState_ReadyForUserInteraction
      
      Case Else
        Phosphorus.pExceptions.Raise Phosphorus.Exceptions.WindowsDriverUnhandledDriverType, "Driver Type: #" & This.DriverType
        
    End Select

    If This.PageLoadedElementPPath <> "" Then
    
      'Wait for the Page Load Element
      Dim FoundElement As Phosphorus.pWindowsDriverElement
      Set FoundElement = FindElement("PageLoadedElement", This.PageLoadedElementPPath, TimeoutInSeconds)
      
      'Store this as the Master Window element so that we can easily close the current driver window when it is finished with
      'Set This.MasterWindowUIAElement = FoundElement.GetUIAElement
      Set This.MasterWindowsDriverElement = FoundElement
      This.MasterWindowsDriverElement.SetName "MasterWindowsDriverElement"
      
      'Set the default element PPath which indicate the driver has been unloaded/closed
      Me.SetPageUnloadedElementPPath GetPageLoadedElementPPath
   
      'Wait until it is in the right WindowInteractionState
'TODO: This only applies to (top level?) Window controls?
      FoundElement.WaitForWindowInteractionState This.PageLoadedElementExpectedWindowInteractionState
            
      'Get the App's PID from the PageLoadedElement
      This.ProcessId = FoundElement.GetProcessID
      
    End If

'TODO: How to open a new tab for 2nd locator, etc?
  
  End If
End Sub

'Wait for Ready For User Interaction State of the page load element here!
'    UIAutomationClient.WindowInteractionState.WindowInteractionState_ReadyForUserInteraction
'WindowInteractionState_BlockedByModalWindow

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

Public Function FindElement( _
  ByVal Name As String, _
  ByVal PPathString As String, _
  Optional ByVal TimeoutInSeconds As Integer, _
  Optional ByVal GetPID As Boolean, _
  Optional ByVal CheckExistenceOnly As Boolean = False) As pWindowsDriverElement
   
  Dim CurrentPPath As Phosphorus.PPath
  Set CurrentPPath = Nothing
  Set This.CurrentEvaluatedPPath = Nothing
  
  Set CurrentPPath = Phosphorus.Factory.GetNewPhosphorusPPath

  CurrentPPath.Initialise
  CurrentPPath.SetApplicationRootElement Phosphorus.pWindowsDriverStatic.gUIADesktopUIElement
  
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
    Set This.CurrentEvaluatedPPath = CurrentPPath.Evaluate(PPathString)
    boolElementFound = This.CurrentEvaluatedPPath.ReturnedValue
    If Not boolElementFound Then
      boolPassedEndTime = (Now > dtEndTime)
      i = i + 1
      Logger.ExternalDebug "Sleeping " & i & ": " & DEFAULT_IMPLICIT_DELAY_BETWEEN_POLLS_IN_MILLISECOND
      Phosphorus.WindowsProcesses.Snooze DEFAULT_IMPLICIT_DELAY_BETWEEN_POLLS_IN_MILLISECOND
    End If
  Wend

  If Not boolElementFound And Not CheckExistenceOnly Then
    
    'Element not found - raise exception
    Phosphorus.pExceptions.Raise _
      Phosphorus.Exceptions.WindowsDriverUIElementNotFoundBeforeTimeout, _
      VBA.Conversion.CStr(TimeoutInSeconds), _
      PPathString
  
  End If
  
  Dim FoundUIAElement As UIAutomationClient.IUIAutomationElement
  
  If boolElementFound Then
    
    'Element has been found if we get here!
  
    'TODO: Check for only 1 matching element!
    
    'Store the current UIA element
    Set FoundUIAElement = This.CurrentEvaluatedPPath.GetMatchingElement(1)
    Dim NewWindowsDriverElement As pWindowsDriverElement
    Set NewWindowsDriverElement = New pWindowsDriverElement
    NewWindowsDriverElement.SetUIAElement Name, Me, FoundUIAElement, PPathString
    Set CurrentPPath = Nothing

    'Prepare to Kill PID at end of session! - Needed for test cases and error handling
    If GetPID Then
      This.ProcessId = NewWindowsDriverElement.GetProcessID
    End If
       
  End If
  
  'Return the found window element, if any
  Set FindElement = NewWindowsDriverElement

End Function

Public Function ElementExists(ByVal Name, ByVal PPathString As String) As Boolean
  Me.FindElement Name:=Name, PPathString:=PPathString, TimeoutInSeconds:=0, CheckExistenceOnly:=True
End Function

Public Sub WaitUntilElementNotExists( _
  ByVal Name, _
  ByVal PPathString As String, _
  Optional ByVal TimeoutInSeconds As Integer)
  If IsMissing(TimeoutInSeconds) Then
    TimeoutInSeconds = This.ImplicitTimeoutInSeconds
  End If
  MsgBox ElementExists(Name, PPathString)
End Sub

