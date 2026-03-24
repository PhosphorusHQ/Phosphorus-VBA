VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalculatorStandardPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Calculator
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Private Type PageAttributes
  CalculatorControlsSearch As pSearch
  DisplaySearch As pSearch
  DisplayTextControl As String
  Elements As New Scripting.Dictionary
End Type

Private This As PageAttributes

Private KeystrokePositionCounter As Integer

Private Sub Class_Initialize()
  Set This.CalculatorControlsSearch = Factory.GetNewSearch
  Set This.DisplaySearch = Factory.GetNewSearch
  Set This.Elements = Nothing
  This.DisplayTextControl = "Display"
End Sub

Private Sub Class_Terminate()
  Set This.CalculatorControlsSearch = Nothing
  Set This.Elements = Nothing
End Sub

Public Sub FindAllControls(HomePage As CalculatorHomePage)

  Dim LandmarkGroupControl As IUIAutomationElement
  Set LandmarkGroupControl = HomePage.GetCurrentCalculatorLandmarkElement

  Dim CalculatorControls() As IUIAutomationElement
  With This.CalculatorControlsSearch
    .Initialise "CalculatorControls", LandmarkGroupControl, Descendants, pConditions, "AND(ControlType, ClassName)"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Button
    .Condition "ClassName", ClassName, IsTheString, "Button"
    .FindAll
  End With
  CalculatorControls = This.CalculatorControlsSearch.FoundUIAElements

  Dim i As Integer
  Dim CurrentUIAElement As IUIAutomationElement
  With This.Elements
    For i = 0 To UBound(CalculatorControls)
      Set CurrentUIAElement = CalculatorControls(i)
      Select Case CurrentUIAElement.CurrentName
        Case "Open history flyout"
          .Add "OpenHistory", CurrentUIAElement
        Case "Clear all memory"
          .Add "MC", CurrentUIAElement
        Case "Memory recall"
          .Add "MR", CurrentUIAElement
        Case "Memory add"
          .Add "M+", CurrentUIAElement
        Case "Memory subtract"
          .Add "M-", CurrentUIAElement
        Case "Memory store"
          .Add "MS", CurrentUIAElement
        Case "Open memory flyout"
          .Add "OpenMemory", CurrentUIAElement
        Case "Percent"
          .Add "%", CurrentUIAElement
        Case "Clear entry"
          .Add "CE", CurrentUIAElement
        Case "Clear"
          .Add "C", CurrentUIAElement
        Case "Backspace"
          .Add "BS", CurrentUIAElement
        Case "Reciprocal"
          .Add "1/x", CurrentUIAElement
        Case "Square"
          .Add "Sqr", CurrentUIAElement
        Case "Square root"
          .Add "SqrRt", CurrentUIAElement
        Case "Divide by"
          .Add "/", CurrentUIAElement
        Case "Multiply by"
          .Add "x", CurrentUIAElement
        Case "Minus"
          .Add "-", CurrentUIAElement
        Case "Plus"
          .Add "+", CurrentUIAElement
        Case "Positive negative"
          .Add "Negate", CurrentUIAElement
        Case "Equals"
          .Add "=", CurrentUIAElement
        Case "Zero"
          .Add "0", CurrentUIAElement
        Case "One"
          .Add "1", CurrentUIAElement
        Case "Two"
          .Add "2", CurrentUIAElement
        Case "Three"
          .Add "3", CurrentUIAElement
        Case "Four"
          .Add "4", CurrentUIAElement
        Case "Five"
          .Add "5", CurrentUIAElement
        Case "Six"
          .Add "6", CurrentUIAElement
        Case "Seven"
          .Add "7", CurrentUIAElement
        Case "Eight"
          .Add "8", CurrentUIAElement
        Case "Nine"
          .Add "9", CurrentUIAElement
        Case "Decimal separator"
          .Add ".", CurrentUIAElement
        Case Else
          MsgBox "Error - button not handed!? '" & CurrentUIAElement.CurrentName & "'"
    End Select
  Next i
  End With
  
  This.Elements.Add This.DisplayTextControl, Nothing
  With This.DisplaySearch
    .Initialise This.DisplayTextControl, LandmarkGroupControl, Descendants, pConditions, "AND(ControlType, AutomationId)"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Text
    .Condition "AutomationId", UIAProperties.AutomationId, IsTheString, "CalculatorResults"
  End With

End Sub

Public Sub Calculate(Keystrokes As String, Answer As String)
    
  If Keystrokes = "" Then
    'No Keystrokes - ignore!?
    Exit Sub
  End If
    
  Dim CountOfLeftBraces As Integer
  Dim CountOfRightBraces As Integer
  CountOfLeftBraces = Utils.CountOccurrences(Keystrokes, "{")
  CountOfRightBraces = Utils.CountOccurrences(Keystrokes, "}")
  Debug.Assert (CountOfLeftBraces = CountOfRightBraces)
  If Not (CountOfLeftBraces = CountOfRightBraces) Then
    Debug.Print "Standard Calculator Skipped: '" & Keystrokes & "'"
    Exit Sub
  End If
  
  Dim Continue As Boolean
  KeystrokePositionCounter = 1
  Continue = True
  While Continue
    Dim NextButtonName As String
    NextButtonName = GetNextButtonName(Keystrokes)
    If NextButtonName <> "" Then
      Actions.Click NextButtonName, This.Elements(NextButtonName)
    Else
      Continue = False
    End If
  Wend
  
  Debug.Assert (GetDisplay = Answer)
  
End Sub

Private Function GetNextButtonName(CurrentKeystrokes As String)
  Dim NextButton As String
  NextButton = VBA.Strings.Mid(CurrentKeystrokes, KeystrokePositionCounter, 1)
  If NextButton = "{" Then
    Dim RightBracketPosition As Integer
    RightBracketPosition = VBA.Strings.InStr(KeystrokePositionCounter + 1, CurrentKeystrokes, "}")
    NextButton = VBA.Strings.Mid(CurrentKeystrokes, KeystrokePositionCounter + 1, RightBracketPosition - KeystrokePositionCounter - 1)
    KeystrokePositionCounter = RightBracketPosition
  End If
  KeystrokePositionCounter = KeystrokePositionCounter + 1
  GetNextButtonName = NextButton
End Function

Private Function GetDisplay() As String
  This.DisplaySearch.Find
  Set This.Elements(This.DisplayTextControl) = This.DisplaySearch.FoundUIAElement
  Dim Display As IUIAutomationElement
  Set Display = This.Elements(This.DisplayTextControl)
  GetDisplay = VBA.Strings.Replace(Display.CurrentName, "Display is ", "")
End Function
