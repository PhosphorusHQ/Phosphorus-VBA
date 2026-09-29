Attribute VB_Name = "ExampleDomainDotCom"
'@Folder ExampleDomainDotCom
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Private ExampleDomain As ExampleDomainDotComPage

Function ExampleDomainDotCom() As Boolean

  Dim Succeeded As Boolean
  Succeeded = False
  
  On Error GoTo ErrorHandler
    
  If Not RunningAllExamples Then
    WebBrowserCommon.GetInternetSpeeds
    Window.HighlightElements = True
    Factory.CurrentWebBrowserType = 1
  End If
  
  Set ExampleDomain = New ExampleDomainDotComPage
  
  ExampleDomain.Initialize
  ExampleDomain.RunChecks
  Succeeded = True
  GoTo ExitSub

ErrorHandler:
  If RunningAllExamples Then
    Debug.Print Err.Description & " (Error Number #" & Err.Number & ") for " & Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser in 'ExampleDomainDotCom' Example!"
  Else
    MsgBox _
      Err.Description & " (Error Number #" & Err.Number & ") " & _
        "for " & Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser " & _
        "in 'ExampleDomainDotCom' Example!", _
      vbCritical
  End If
  GoTo ExitSub
  
ExitSub:
  Set ExampleDomain = Nothing
  If Not RunningAllExamples Then
    Window.HighlightElements = False
  End If
  
  ExampleDomainDotCom = Succeeded

End Function
