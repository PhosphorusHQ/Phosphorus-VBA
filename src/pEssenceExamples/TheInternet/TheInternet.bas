Attribute VB_Name = "TheInternet"
'@Folder TheInternet
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Function TheInternet() As Boolean

  Dim Succeeded As Boolean
  Succeeded = False
  
  On Error GoTo ErrorHandler

  If Not RunningAllExamples Then
    WebBrowserCommon.GetInternetSpeeds
    Window.HighlightElements = True
    Factory.CurrentWebBrowserType = 1
  End If

  Dim The_Internet As TheInternetPage
  Set The_Internet = New TheInternetPage

  With The_Internet
    .Initialize
    .RunHomePageChecks
    .Checkboxes
    .DragAndDrop
    .FormAuthentication
  End With
  Succeeded = True
  GoTo ExitSub

ErrorHandler:
  If RunningAllExamples Then
    Debug.Print Err.Description & " (Error Number #" & Err.Number & ") for " & Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser in 'TheInternet' Example!"
  Else
    MsgBox _
      Err.Description & " (Error Number #" & Err.Number & ") " & _
        "for " & Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser " & _
        "in 'TheInternet' Example!", _
      vbCritical
      GoTo ExitSub
  End If
  
ExitSub:
  Set The_Internet = Nothing
  If Not RunningAllExamples Then
    Window.HighlightElements = False
  End If

  TheInternet = Succeeded

End Function
