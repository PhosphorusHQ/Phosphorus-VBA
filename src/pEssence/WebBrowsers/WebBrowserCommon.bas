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
  
  Factory.CurrentWebBrowserType = Edge
  Set WebBrowser = Factory.GetNewWebBrowser
  WebBrowser.Start _
    "Speedtest by Ookla", "https://www.speedtest.net/", "Speedtest by Ookla - The Global Broadband Speed Test", _
    AbsoluteWaitTimeSeconds:=10

  Dim Root As pLocator
  Set Root = WebBrowser.GetRootWebArea
  
  Set Go = Factory.GetNewLocator
  With Go
    .Initialise _
      "Go", Root, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleHeading, NameIs)", FindFirst:=True: .AriaRoleHeading: .NameIs "GO"
    .Find 10
    Snooze 2000
    .Element.Click
  End With

  Set ResultID = Factory.GetNewLocator
  With ResultID
    .Initialise _
      "ResultID", Root, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleDescription, NameIs)": .AriaRoleDescription: .NameIs "Result ID:":
     .Find 120
    End With

  Set DOWNLOAD = New pLocator
  With DOWNLOAD
    .Initialise _
      "DOWNLOAD", Root, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleDescription, NameIs)": .AriaRoleDescription: .NameIs "DOWNLOAD": .RelativeElementNumber 3
    .Find 2
  End With
  DownloadSpeedMbps = DOWNLOAD.Element.Name
  
  Set UPLOAD = New pLocator
  With UPLOAD
    .Initialise _
      "UPLOAD", Root, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleDescription, NameIs)": .AriaRoleDescription: .NameIs "UPLOAD": .RelativeElementNumber 3
    .Find 2
  End With
  UploadSpeedMbps = UPLOAD.Element.Name

  Set WebBrowser = Nothing
  Set Go = Nothing
  Set ResultID = Nothing
  Set DOWNLOAD = Nothing
  Set UPLOAD = Nothing
  
  Debug.Print "Download speed is (Mbps): " & DownloadSpeedMbps & ", " & "Upload speed is (Mbps): " & UploadSpeedMbps
  
End Sub

