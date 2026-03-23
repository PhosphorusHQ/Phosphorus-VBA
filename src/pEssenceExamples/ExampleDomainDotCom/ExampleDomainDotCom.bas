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

  Window.HighlightElements = True
  
  Set ExampleDomain = New ExampleDomainDotComPage
  
  ExampleDomain.Initialize
  GoTo ExitSub

ErrorHandler:
  MsgBox Err.Description & " (Error Number #" & Err.Number & ")"
  Exit Sub
  
ExitSub:
  Set ExampleDomain = Nothing
  Window.HighlightElements = False
  
MsgBox "What next? More Edge Window elements!"

End Sub
