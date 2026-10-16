Attribute VB_Name = "WebBrowserCommon"
'@Folder WebBrowsers
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2026 Peter Jeffrey Gale
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
    GetInternetSpeedsFromGoogle
  End If
End Sub

Private Sub GetInternetSpeedsFromOokla()

  Dim WebBrowser As Object
  Dim Root As pLocator
  Dim Go As pLocator
  Dim DOWNLOAD As pLocator
  Dim UPLOAD As pLocator

  Factory.CurrentWebBrowserType = Chrome 'Switched to Chrome from Edge as using Edge here causes problems for launching later Edge browser
  Set WebBrowser = Factory.GetNewWebBrowser
  WebBrowser.Start _
    "Speedtest by Ookla", "https://www.speedtest.net/", "Speedtest by Ookla - The Global Broadband Speed Test", _
    AbsoluteWaitTimeSeconds:=10

  Set Root = WebBrowser.GetRootWebArea
  
  Set Go = Factory.GetNewLocator
  With Go
    .Initialise _
      "Go", Root, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleButton, NameIs)", FindFirst:=True: .AriaRoleButton: .NameIs "start speed test - connection type multi"
    .Find 10
    Snooze 2000
    .Element.Click
  End With

  WebBrowser.WaitForNewURL 120
    
  Set DOWNLOAD = New pLocator
  With DOWNLOAD
    .Initialise "DOWNLOAD", Root, TreeScope.Descendants, By.NameIs, "Mbps": .PositionInMatchingSet 1: .RelativeElementNumber 1
    .Find 2
  End With
  DownloadSpeedMbps = DOWNLOAD.Element.Name
  
  Set UPLOAD = New pLocator
  With UPLOAD
    .Initialise "UPLOAD", Root, TreeScope.Descendants, By.NameIs, "Mbps": .PositionInMatchingSet 2: .RelativeElementNumber 1
    .Find 2
  End With
  UploadSpeedMbps = UPLOAD.Element.Name

  Set WebBrowser = Nothing
  Set Go = Nothing
  Set DOWNLOAD = Nothing
  Set UPLOAD = Nothing
  
  Debug.Print "Ookla: Download speed is (Mbps): " & DownloadSpeedMbps & ", " & "Upload speed is (Mbps): " & UploadSpeedMbps

End Sub

Private Sub GetInternetSpeedsFromGoogle()

  Dim WebBrowser As Object
  Dim Root As pLocator
  Dim RunTest As pLocator
  Dim TestAgain As pLocator
  Dim DOWNLOAD As pLocator
  Dim UPLOAD As pLocator

  Factory.CurrentWebBrowserType = Chrome 'Switched to Chrome from Edge as using Edge here causes problems for launching later Edge browser
  Set WebBrowser = Factory.GetNewWebBrowser
  WebBrowser.Start _
    "Speedtest by Google", "https://www.google.com/search?q=Speed+Test", "Speed Test - Google Search", _
    AbsoluteWaitTimeSeconds:=10
  
  Set Root = WebBrowser.GetRootWebArea
  
  Set RunTest = Factory.GetNewLocator
  With RunTest
    .Initialise _
      "RunTest", Root, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleButton, NameIs)", FindFirst:=True: .AriaRoleButton: .NameIs "RUN SPEED TEST"
    .Find 10
    .Element.Click
  End With

  Set TestAgain = Factory.GetNewLocator
  With TestAgain
    .Initialise _
      "TestAgain", Root, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleButton, NameIs)", FindFirst:=True: .AriaRoleButton: .NameIs "TEST AGAIN"
    .ElementExists 60
  End With
    
  Set DOWNLOAD = New pLocator
  With DOWNLOAD
    .Initialise "DOWNLOAD", Root, TreeScope.Descendants, By.NameIs, "Mbps download": .RelativeElementNumber -1
    .Find 2
  End With
  DownloadSpeedMbps = DOWNLOAD.Element.Name
  
  Set UPLOAD = New pLocator
  With UPLOAD
    .Initialise "UPLOAD", Root, TreeScope.Descendants, By.NameIs, "Mbps upload": .RelativeElementNumber -1
    .Find 2
  End With
  UploadSpeedMbps = UPLOAD.Element.Name

  Set WebBrowser = Nothing
  Set RunTest = Nothing
  Set TestAgain = Nothing
  Set DOWNLOAD = Nothing
  Set UPLOAD = Nothing
  
  Debug.Print "Google: Download speed is (Mbps): " & DownloadSpeedMbps & ", " & "Upload speed is (Mbps): " & UploadSpeedMbps

End Sub

Public Sub Navigate( _
  CurrentWebBrowser As Object, _
  NavigationButton As pLocator, _
  AddressElement As pLocator, _
  RootWebArea As pLocator, _
  Optional RefreshButton As pLocator)
  Dim CurrentURL As String
  CurrentURL = CurrentWebBrowser.GetCurrentURL
  With NavigationButton
    .Find 10
    .Element.Click
  End With
  AddressElement.Element.WaitForPatternState UIAPatterns.Value, CurrentURL, 10, True
  If RefreshButton Is Nothing Then
    RootWebArea.Find 10, FindElementAgain:=True
  Else
    'We are able to try refreshing!
    Dim Continue As Boolean, i As Integer
    Continue = True
    i = 0
    While Continue
      i = i + 1
      Continue = (i <= 3)
      If Continue Then
        If RootWebArea.ElementExists(10) Then
          Continue = False
        Else
          RefreshButton.Element.Click
        End If
      End If
    Wend
  End If
End Sub

Public Sub WaitForNewURL(CurrentURL As String, AddressElement As pLocator, RootWebArea As pLocator, TimeoutInSeconds As Integer)
  AddressElement.Element.WaitForPatternState UIAPatterns.Value, CurrentURL, TimeoutInSeconds, True
  RootWebArea.Find 10, FindElementAgain:=True
End Sub
