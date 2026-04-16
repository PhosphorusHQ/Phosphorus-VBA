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
  CalculatorControls As pLocator
  DisplaySearch As pLocator
  DisplayTextControl As String
  Elements As New Scripting.Dictionary
  OpenHistoryFlyout As pElement
  ClearAllMemory As pElement
  MemoryRecall As pElement
  MemoryAdd As pElement
  MemorySubtract As pElement
  MemoryStore As pElement
  OpenMemoryFlyout As pElement
  Percent As pElement
  ClearEntry As pElement
  Clear As pElement
  Backspace As pElement
  Reciprocal As pElement
  Square As pElement
  SquareRoot As pElement
  DivideBy As pElement
  MultiplyBy As pElement
  Minus As pElement
  Plus As pElement
  PositiveNegative As pElement
  Equals As pElement
  Zero As pElement
  One As pElement
  Two As pElement
  Three As pElement
  Four As pElement
  Five As pElement
  Six As pElement
  Seven As pElement
  Eight As pElement
  Nine As pElement
  DecimalSeparator As pElement
End Type

Private This As PageAttributes

Private KeystrokePositionCounter As Integer

Private Sub Class_Initialize()
  Set This.CalculatorControls = Factory.GetNewLocator
  Set This.DisplaySearch = Factory.GetNewLocator
  Set This.Elements = Nothing
  This.DisplayTextControl = "Display"
End Sub

Private Sub Class_Terminate()
  Set This.CalculatorControls = Nothing
  Set This.OpenHistoryFlyout = Nothing
  Set This.ClearAllMemory = Nothing
  Set This.MemoryRecall = Nothing
  Set This.MemoryAdd = Nothing
  Set This.MemorySubtract = Nothing
  Set This.MemoryStore = Nothing
  Set This.OpenMemoryFlyout = Nothing
  Set This.Percent = Nothing
  Set This.ClearEntry = Nothing
  Set This.Clear = Nothing
  Set This.Backspace = Nothing
  Set This.Reciprocal = Nothing
  Set This.Square = Nothing
  Set This.SquareRoot = Nothing
  Set This.DivideBy = Nothing
  Set This.MultiplyBy = Nothing
  Set This.Minus = Nothing
  Set This.Plus = Nothing
  Set This.PositiveNegative = Nothing
  Set This.Equals = Nothing
  Set This.Zero = Nothing
  Set This.One = Nothing
  Set This.Two = Nothing
  Set This.Three = Nothing
  Set This.Four = Nothing
  Set This.Five = Nothing
  Set This.Six = Nothing
  Set This.Seven = Nothing
  Set This.Eight = Nothing
  Set This.Nine = Nothing
  Set This.DecimalSeparator = Nothing
  Set This.Elements = Nothing
End Sub

Public Sub FindAllControls(HomePage As CalculatorHomePage)

  Dim CurrentCalculatorLandmarkGroupControl As pLocator
  Set CurrentCalculatorLandmarkGroupControl = HomePage.GetCurrentCalculatorLandmarkGroupControl
  CurrentCalculatorLandmarkGroupControl.Find
  
  Dim CalculatorControls() As pElement
  With This.CalculatorControls
    .Initialise "CalculatorControls", CurrentCalculatorLandmarkGroupControl, Descendants, pConditions, "AND(ControlType, ClassName)"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Button
    .Condition "ClassName", ClassName, IsTheString, "Button"
    .FindAll
  End With
  CalculatorControls = This.CalculatorControls.Elements

  Dim i As Integer
  Dim CurrentUIAElement As IUIAutomationElement
  With This.Elements
    For i = 0 To UBound(CalculatorControls)
      Set CurrentUIAElement = CalculatorControls(i).UIAElement
      Select Case CurrentUIAElement.CurrentName
        Case "Open history flyout"
          Set This.OpenHistoryFlyout = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "OpenHistory", This.OpenHistoryFlyout
        Case "Clear all memory"
          Set This.ClearAllMemory = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "MC", This.ClearAllMemory
        Case "Memory recall"
          Set This.MemoryRecall = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "MR", This.MemoryRecall
        Case "Memory add"
          Set This.MemoryAdd = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "M+", This.MemoryAdd
        Case "Memory subtract"
          Set This.MemorySubtract = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "M-", This.MemorySubtract
        Case "Memory store"
          Set This.MemoryStore = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "MS", This.MemoryStore
        Case "Open memory flyout"
          Set This.OpenMemoryFlyout = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "OpenMemory", This.OpenMemoryFlyout
        Case "Percent"
          Set This.Percent = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "%", This.Percent
        Case "Clear entry"
          Set This.ClearEntry = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "CE", This.ClearEntry
        Case "Clear"
          Set This.Clear = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "C", This.Clear
        Case "Backspace"
          Set This.Backspace = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "BS", This.Backspace
        Case "Reciprocal"
          Set This.Reciprocal = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "1/x", This.Reciprocal
        Case "Square"
          Set This.Square = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "Sqr", This.Square
        Case "Square root"
          Set This.SquareRoot = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "SqrRt", This.SquareRoot
        Case "Divide by"
          Set This.DivideBy = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "/", This.DivideBy
        Case "Multiply by"
          Set This.MultiplyBy = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "x", This.MultiplyBy
        Case "Minus"
          Set This.Minus = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "-", This.Minus
        Case "Plus"
          Set This.Plus = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "+", This.Plus
        Case "Positive negative"
          Set This.PositiveNegative = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "Negate", This.PositiveNegative
        Case "Equals"
          Set This.Equals = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "=", This.Equals
        Case "Zero"
          Set This.Zero = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "0", This.Zero
        Case "One"
          Set This.One = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "1", This.One
        Case "Two"
          Set This.Two = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "2", This.Two
        Case "Three"
          Set This.Three = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "3", This.Three
        Case "Four"
          Set This.Four = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "4", This.Four
        Case "Five"
          Set This.Five = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "5", This.Five
        Case "Six"
          Set This.Six = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "6", This.Six
        Case "Seven"
          Set This.Seven = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "7", This.Seven
        Case "Eight"
          Set This.Eight = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "8", This.Eight
        Case "Nine"
          Set This.Nine = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add "9", This.Nine
        Case "Decimal separator"
          Set This.DecimalSeparator = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          .Add ".", This.DecimalSeparator
        Case Else
          MsgBox "Error - button not handed!? '" & CurrentUIAElement.CurrentName & "'"
    End Select
  Next i
  End With
  
  This.Elements.Add This.DisplayTextControl, Nothing
  With This.DisplaySearch
    .Initialise This.DisplayTextControl, CurrentCalculatorLandmarkGroupControl, Descendants, pConditions, "AND(ControlType, AutomationId)"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Text
    .AutomationId "CalculatorResults"
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
      Dim NextButton As pElement
      Set NextButton = This.Elements(NextButtonName)
      Actions.Click NextButton
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
  Set This.Elements(This.DisplayTextControl) = This.DisplaySearch.Element.UIAElement
  Dim Display As IUIAutomationElement
  Set Display = This.Elements(This.DisplayTextControl)
  GetDisplay = VBA.Strings.Replace(Display.CurrentName, "Display is ", "")
End Function
