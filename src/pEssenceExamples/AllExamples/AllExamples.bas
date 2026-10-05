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
Private RunUnstableTests As Boolean
Dim TargetWBT As WebBrowserType

Public Enum Examples
  pPathExample
  pPathExample_SubTest
  CalculatorExample
  HelloWorldExample
  ExampleDomainDotComExample
  LetCodeDotInRadioButtonsAndCheckboxesExample
  TheInternetExample
End Enum

Public Sub RunAllExamples()
  WebBrowserCommon.ForgetInternetSpeeds
  Dim i As Integer
  For i = 1 To 1
    RunUnstableTests = True
    RunAnExample Examples.pPathExample
    RunAnExample Examples.CalculatorExample
'TargetWBT = WebBrowserType.Brave
TargetWBT = 0 'All!
'    RunAnExample Examples.HelloWorldExample
'PJG Needs rebuilding on NAS WordPress    RunAnExample Examples.ExampleDomainDotComExample
'LetCodeIn Needs rebuiding on NAS WordPress?
    RunAnExample Examples.LetCodeDotInRadioButtonsAndCheckboxesExample
    RunAnExample Examples.TheInternetExample
  Next i
  MsgBox "All Examples Run!"
End Sub

Private Sub RunAnExample(Example As Examples)

  RunningAllExamples = True
  Window.HighlightElements = False
  Phosphorus.Log4PStatic.GetLogger
  
  If Example = Examples.CalculatorExample Then
    Calculator.Calculator
  ElseIf Example = Examples.pPathExample Then
    pPathExamples.RunAllpPathTestsWithMultipleTries
  Else
'    WebBrowserCommon.GetInternetSpeeds
WebBrowserCommon.SetDummyMobileDataInternetSpeeds

    Dim WBTFrom As WebBrowserType
    Dim WBTTo As WebBrowserType
    If TargetWBT = 0 Then
      WBTFrom = WebBrowserType.[_First] + 1
      WBTTo = WebBrowserType.[_Last] - 1
    Else
      WBTFrom = TargetWBT
      WBTTo = TargetWBT
    End If
    Dim WBT As WebBrowserType
    For WBT = WBTFrom To WBTTo
      Factory.CurrentWebBrowserType = WBT
      TryToRunAnExampleMultipleTimes True, Example
    Next WBT
  End If

  Window.HighlightElements = False
  RunningAllExamples = False
  Phosphorus.Log4PStatic.CloseLogger

End Sub

Public Sub TryToRunAnExampleMultipleTimes(WebBasedTest As Boolean, Example As Examples, Optional SubTestName As String)
  
  Dim NumberOfAttemptsAllowed As Integer
  Dim AttemptNumber As Integer
  Dim AttemptSucceeded As Boolean
  Dim ExampleName As String
  
  NumberOfAttemptsAllowed = 3
  AttemptNumber = 1
  AttemptSucceeded = False
    
  While (AttemptNumber <= NumberOfAttemptsAllowed) And Not AttemptSucceeded
  
    Select Case Example
      
      Case Examples.HelloWorldExample
        ExampleName = "HelloWorld"
        AttemptSucceeded = HelloWorld.HelloWorldWideWeb
      
      Case Examples.ExampleDomainDotComExample
        ExampleName = "ExampleDomainDotCom"
        AttemptSucceeded = ExampleDomainDotCom.ExampleDomainDotCom
      
      Case Examples.LetCodeDotInRadioButtonsAndCheckboxesExample
        ExampleName = "LetCodeDotInRadioButtonsAndCheckboxes"
        If (Factory.CurrentWebBrowserType <> WebBrowserType.Brave And Factory.CurrentWebBrowserType <> WebBrowserType.Chromium) Or RunUnstableTests Then
          AttemptSucceeded = LetCodeDotIn.RadioButtonsAndCheckboxes
        End If
      
      Case Examples.TheInternetExample
        ExampleName = "TheInternet"
        If (Factory.CurrentWebBrowserType <> WebBrowserType.Chrome And Factory.CurrentWebBrowserType <> WebBrowserType.Yandex) Or RunUnstableTests Then
          AttemptSucceeded = TheInternet.TheInternet
        End If
      
      Case Examples.pPathExample_SubTest
        ExampleName = SubTestName
        AttemptSucceeded = pPathExamples.RunASinglepPathTest(SubTestName)
    
      Case Else
        MsgBox "Uhandled Example #" & Example
        AttemptSucceeded = True
        
    End Select
    
    If Not AttemptSucceeded Then
      AttemptNumber = AttemptNumber + 1
    End If
  Wend
  
  If WebBasedTest Then
    If AttemptSucceeded Then
      Debug.Print Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser succeeded for '" & ExampleName & "' Example after " & AttemptNumber & " attempt(s)!"
    Else
      Debug.Print Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser failed for '" & ExampleName & "' Example after " & NumberOfAttemptsAllowed & " attempts!"
    End If
  Else
    If AttemptSucceeded Then
      Debug.Print "Test succeeded for '" & ExampleName & "' Example after " & AttemptNumber & " attempt(s)!"
    Else
      Debug.Print "Test failed for '" & ExampleName & "' Example after " & NumberOfAttemptsAllowed & " attempts!"
    End If
  End If
  
End Sub

