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

Sub ExampleDomainDotCom()

  On Error GoTo ErrorHandler

  If Not RunningAllExamples Then
    Window.HighlightElements = True
    Factory.CurrentWebBrowserType = Edge
  End If
  
  Set ExampleDomain = New ExampleDomainDotComPage
  
  ExampleDomain.Initialize
  ExampleDomain.RunChecks
  GoTo ExitSub

ErrorHandler:
  MsgBox Err.Description & " (Error Number #" & Err.Number & ")"
  GoTo ExitSub
  
ExitSub:
  Set ExampleDomain = Nothing
  If Not RunningAllExamples Then
    Window.HighlightElements = False
  End If
  
End Sub
