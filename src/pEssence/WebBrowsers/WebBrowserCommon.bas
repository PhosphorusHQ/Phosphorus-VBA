Attribute VB_Name = "WebBrowserCommon"
'@Folder WebBrowsers
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public DownloadSpeedMbps As Integer
Public UploadSpeedMbps As Integer

Public Sub ForgetInternetSpeeds()
  DownloadSpeedMbps = 0
  UploadSpeedMbps = 0
End Sub

Public Sub SetDummyBroadbandInternetSpeeds()
  DownloadSpeedMbps = 200
  UploadSpeedMbps = 200
End Sub

Public Sub SetDummyMobileDataInternetSpeeds()
  DownloadSpeedMbps = 20
  UploadSpeedMbps = 20
End Sub

Public Sub GetInternetSpeeds()
  If DownloadSpeedMbps = 0 Or UploadSpeedMbps = 0 Then
    GetInternetSpeedsFromOokla
  End If
End Sub

Private Sub GetInternetSpeedsFromOokla()

  Dim WebBrowser As Object
  Dim Go As pLocator
  Dim ResultID As pLocator
  Dim DOWNLOAD As pLocator
  Dim UPLOAD As pLocator

Phosphorus.Log4PStatic.GetLogger

  Factory.CurrentWebBrowserType = Chrome 'Switched to Chrome from Edge as using Edge here causes problems for launching later Edge browser
  Set WebBrowser = Factory.GetNewWebBrowser
  WebBrowser.Start _
    "Speedtest by Ookla", "https://www.speedtest.net/", "Speedtest by Ookla - The Global Broadband Speed Test", _
    AbsoluteWaitTimeSeconds:=10

  Dim Root As pLocator
  Set Root = WebBrowser.GetRootWebArea
  
  Set Go = Factory.GetNewLocator
  With Go
'Edge
'    .Initialise _
      "Go", Root, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleHeading, NameIs)", FindFirst:=True: .AriaRoleHeading: .NameIs "GO"
'Chrome
    .Initialise _
      "Go", Root, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleButton, NameIs)", FindFirst:=True: .AriaRoleButton: .NameIs "start speed test - connection type multi"
    .Find 10
    Snooze 2000
    .Element.Click
  End With

  Set ResultID = Factory.GetNewLocator
  With ResultID
'Edge
'    .Initialise _
      "ResultID", Root, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleDescription, NameIs)": .AriaRoleDescription: .NameIs "Result ID:":
'Chrome
'     .Initialise _
'      "ResultID", Root, TreeScope_Descendants, By.pConditions, _
'      "AND(AriaRoleDescription, NameIs)": .AriaRoleDescription: .NameIs "Result ID":
'     .Find 120
    End With

    WebBrowser.WaitForNewURL 120
    
  Set DOWNLOAD = New pLocator
  With DOWNLOAD
'Edge
'    .Initialise _
      "DOWNLOAD", Root, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleDescription, NameIs)": .AriaRoleDescription: .NameIs "DOWNLOAD", True: .RelativeElementNumber 3
'Chrome
'    .Initialise _
'      "DOWNLOAD", Root, TreeScope_Descendants, By.pConditions, _
'      "AND(AriaRoleDescription, NameIs)": .AriaRoleDescription: .NameIs "DOWNLOAD", True: .RelativeElementNumber 3
    .Initialise "DOWNLOAD", Root, TreeScope.None, By.pPath, "//*[@ClassName=""py-2 font-mono text-5xl MuiBox-root css-s31qlv""][1]"
    .Find 2
  End With
  DownloadSpeedMbps = DOWNLOAD.Element.Name
  
  Set UPLOAD = New pLocator
  With UPLOAD
'Edge
'    .Initialise _
      "UPLOAD", Root, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleDescription, NameIs)": .AriaRoleDescription: .NameIs "UPLOAD", True: .RelativeElementNumber 3
'Chrome
    .Initialise _
      "UPLOAD", Root, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleDescription, NameIs)": .AriaRoleDescription: .NameIs "UPLOAD", True: .RelativeElementNumber 3
    .Find 2
  End With
  UploadSpeedMbps = UPLOAD.Element.Name

  Set WebBrowser = Nothing
  Set Go = Nothing
  Set ResultID = Nothing
  Set DOWNLOAD = Nothing
  Set UPLOAD = Nothing
  
Phosphorus.Log4PStatic.CloseLogger
  
  Debug.Print "Download speed is (Mbps): " & DownloadSpeedMbps & ", " & "Upload speed is (Mbps): " & UploadSpeedMbps

End Sub

Public Sub Navigate(CurrentWebBrowser As Object, NavigationButton As pLocator, AddressElement As pLocator, RootWebArea As pLocator)
  Dim CurrentURL As String
  CurrentURL = CurrentWebBrowser.GetCurrentURL
  With NavigationButton
    .Find 10
    .Element.Click
  End With
  AddressElement.Element.WaitForPatternState UIAPatterns.Value, CurrentURL, 10, True
  RootWebArea.Find 10, FindElementAgain:=True
End Sub

Public Sub WaitForNewURL(CurrentURL As String, AddressElement As pLocator, RootWebArea As pLocator, TimeoutInSeconds As Integer)
  AddressElement.Element.WaitForPatternState UIAPatterns.Value, CurrentURL, TimeoutInSeconds, True
  RootWebArea.Find 10, FindElementAgain:=True
End Sub
