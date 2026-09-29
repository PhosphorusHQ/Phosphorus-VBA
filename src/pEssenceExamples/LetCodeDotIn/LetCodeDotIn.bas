Attribute VB_Name = "LetCodeDotIn"
'@Folder LetCodeDotIn
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Function RadioButtonsAndCheckboxes() As Boolean

  Dim Succeeded As Boolean
  Succeeded = False
  
  On Error GoTo ErrorHandler

  If Not RunningAllExamples Then
    WebBrowserCommon.GetInternetSpeeds
    Window.HighlightElements = True
    Factory.CurrentWebBrowserType = Chromium
  End If
  
  Dim LetCodeDotIn As LetCodeDotInPageRadio
  Set LetCodeDotIn = New LetCodeDotInPageRadio
  
  With LetCodeDotIn
    .Initialize
    .Automate
  End With
  Succeeded = True
  GoTo ExitSub

ErrorHandler:
  If RunningAllExamples Then
    Debug.Print Err.Description & " (Error Number #" & Err.Number & ") for " & Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser in 'LetCodeDotIn' Example!"
  Else
    MsgBox _
      Err.Description & " (Error Number #" & Err.Number & ") " & _
        "for " & Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser " & _
        "in 'LetCodeDotIn' Example!", _
      vbCritical
  End If

ExitSub:
  Set LetCodeDotIn = Nothing
  If Not RunningAllExamples Then
    Window.HighlightElements = False
  End If
  
  RadioButtonsAndCheckboxes = Succeeded

End Function

