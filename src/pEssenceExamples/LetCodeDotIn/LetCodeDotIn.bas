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

Sub LetCode_RadioButtonsAndCheckboxes()

  On Error GoTo ErrorHandler

  Window.HighlightElements = True
  
  Dim LetCodeDotIn As LetCodeDotInPageRadio
  Set LetCodeDotIn = New LetCodeDotInPageRadio
  
  With LetCodeDotIn
    .Initialize
    .Automate
  End With
  GoTo ExitSub

ErrorHandler:
  MsgBox Err.Description & " (Error Number #" & Err.Number & ")"
  
ExitSub:
  Set LetCodeDotIn = Nothing
  Window.HighlightElements = False

End Sub

