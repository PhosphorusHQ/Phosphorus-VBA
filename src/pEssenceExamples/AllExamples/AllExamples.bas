Attribute VB_Name = "AllExamples"
'@Folder AllExamples
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public RunningAllExamples As Boolean

Public Sub AllExamples()
  
  RunningAllExamples = True
  Window.HighlightElements = False

  Calculator.Calculator
  ExampleDomainDotCom.ExampleDomainDotCom
  LetCodeDotIn.RadioButtonsAndCheckboxes
  TheInternet.TheInternet
  
  Window.HighlightElements = False
  RunningAllExamples = False

  MsgBox "What next? Select Interaction speeds - slow, medium, fast, no delay? More pages ..."

End Sub
