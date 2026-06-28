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

Public Enum Examples
  CalculatorExample
  HelloWorldExample
  ExampleDomainDotComExample
  LetCodeDotInExample
  TheInternetExample
End Enum

Sub RunAllExamples()
  WebBrowserCommon.ForgetInternetSpeeds
  Dim i As Integer
  For i = 1 To 1
    RunAnExample Examples.CalculatorExample
    RunAnExample Examples.HelloWorldExample
    RunAnExample Examples.ExampleDomainDotComExample
    RunAnExample Examples.LetCodeDotInExample
    RunAnExample Examples.TheInternetExample
  Next i
  MsgBox "All Examples Run!"
End Sub

Public Sub RunAnExample(Example As Examples)
  
  RunningAllExamples = True
  Window.HighlightElements = False
  
  If Example = Examples.CalculatorExample Then
    Calculator.Calculator
  Else
    WebBrowserCommon.GetInternetSpeeds
    Dim WBT As WebBrowserType
    For WBT = WebBrowserType.[_First] + 1 To WebBrowserType.[_Last] - 1
      Factory.CurrentWebBrowserType = WBT
      Select Case Example
        Case Examples.HelloWorldExample
          HelloWorld.HelloWorldWideWeb
        Case Examples.ExampleDomainDotComExample
          ExampleDomainDotCom.ExampleDomainDotCom
        Case Examples.LetCodeDotInExample
          LetCodeDotIn.RadioButtonsAndCheckboxes
        Case Examples.TheInternetExample
          TheInternet.TheInternet
      End Select
    Next WBT
  End If

  Window.HighlightElements = False
  RunningAllExamples = False

End Sub


