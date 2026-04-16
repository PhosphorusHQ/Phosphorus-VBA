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

Private The_Internet As TheInternetPage

Sub TheInternet()

  On Error GoTo ErrorHandler

  If Not RunningAllExamples Then
    Window.HighlightElements = True
    Factory.CurrentWebBrowserType = 1
  End If
  
  Set The_Internet = New TheInternetPage
        
  With The_Internet
    .Initialize
    .RunHomePageChecks
    .Checkboxes
    .DragAndDrop
    .FormAuthentication
  End With
  GoTo ExitSub

ErrorHandler:
  MsgBox Err.Description & " (Error Number #" & Err.Number & ")"
  GoTo ExitSub
  
ExitSub:
  Set The_Internet = Nothing
  If Not RunningAllExamples Then
    Window.HighlightElements = False
  End If
  
End Sub
