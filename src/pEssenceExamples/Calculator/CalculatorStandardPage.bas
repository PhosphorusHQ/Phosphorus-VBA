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
Option Explicit

Private CalculatorControlElementsSearch As pSearch
Private DisplayElementSearch As pSearch

Private Type PrivateElementNames
  LandmarkGroupControl As String
  DisplayTextControl As String
End Type

Private This As PrivateElementNames
Private PrivateElements As New Scripting.Dictionary

Private KeystrokePositionCounter As Integer

Private Sub Class_Initialize()
  Set CalculatorControlElementsSearch = UIACommon.GetNewSearch
  Set DisplayElementSearch = UIACommon.GetNewSearch
  Set PrivateElements = Nothing
  This.LandmarkGroupControl = "Landmark"
  This.DisplayTextControl = "Display"
End Sub

Private Sub Class_Terminate()
  Set CalculatorControlElementsSearch = Nothing
  Set PrivateElements = Nothing
End Sub

Public Sub FindAllControls(HomePage As CalculatorHomePage)

  PrivateElements.Add This.LandmarkGroupControl, HomePage.GetCurrentCalculatorLandmarkElement
  
  Dim CalculatorControls() As IUIAutomationElement
  With CalculatorControlElementsSearch
    .Initialise This.LandmarkGroupControl, PrivateElements(This.LandmarkGroupControl), TreeScope.Descendants
    .AddCondition "ControlTypeIsButton", UIAProperties.ControlType, UIAPropertyComparisons.Equals, UIAControlTypeIDs.Button
    .AddCondition "ClassNameIsButton", UIAProperties.ClassName, UIAPropertyComparisons.Equals, "Button"
    .Locator by.pConditions, "AND(ControlTypeIsButton, ClassNameIsButton)"
     CalculatorControls = .FindAll
  End With

  Dim i As Integer
  For i = 0 To UBound(CalculatorControls)
    Select Case CalculatorControls(i).CurrentName
      Case "Open history flyout"
        PrivateElements.Add "OpenHistory", CalculatorControls(i)
      Case "Clear all memory"
        PrivateElements.Add "MC", CalculatorControls(i)
      Case "Memory recall"
        PrivateElements.Add "MR", CalculatorControls(i)
      Case "Memory add"
        PrivateElements.Add "M+", CalculatorControls(i)
      Case "Memory subtract"
        PrivateElements.Add "M-", CalculatorControls(i)
      Case "Memory store"
        PrivateElements.Add "MS", CalculatorControls(i)
      Case "Open memory flyout"
        PrivateElements.Add "OpenMemory", CalculatorControls(i)
      Case "Percent"
        PrivateElements.Add "%", CalculatorControls(i)
      Case "Clear entry"
        PrivateElements.Add "CE", CalculatorControls(i)
      Case "Clear"
        PrivateElements.Add "C", CalculatorControls(i)
      Case "Backspace"
        PrivateElements.Add "BS", CalculatorControls(i)
      Case "Reciprocal"
        PrivateElements.Add "1/x", CalculatorControls(i)
      Case "Square"
        PrivateElements.Add "Sqr", CalculatorControls(i)
      Case "Square root"
        PrivateElements.Add "SqrRt", CalculatorControls(i)
      Case "Divide by"
        PrivateElements.Add "/", CalculatorControls(i)
      Case "Multiply by"
        PrivateElements.Add "x", CalculatorControls(i)
      Case "Minus"
        PrivateElements.Add "-", CalculatorControls(i)
      Case "Plus"
        PrivateElements.Add "+", CalculatorControls(i)
      Case "Positive negative"
        PrivateElements.Add "Negate", CalculatorControls(i)
      Case "Equals"
        PrivateElements.Add "=", CalculatorControls(i)
      Case "Zero"
        PrivateElements.Add "0", CalculatorControls(i)
      Case "One"
        PrivateElements.Add "1", CalculatorControls(i)
      Case "Two"
        PrivateElements.Add "2", CalculatorControls(i)
      Case "Three"
        PrivateElements.Add "3", CalculatorControls(i)
      Case "Four"
        PrivateElements.Add "4", CalculatorControls(i)
      Case "Five"
        PrivateElements.Add "5", CalculatorControls(i)
      Case "Six"
        PrivateElements.Add "6", CalculatorControls(i)
      Case "Seven"
        PrivateElements.Add "7", CalculatorControls(i)
      Case "Eight"
        PrivateElements.Add "8", CalculatorControls(i)
      Case "Nine"
        PrivateElements.Add "9", CalculatorControls(i)
      Case "Decimal separator"
        PrivateElements.Add ".", CalculatorControls(i)
      Case Else
        MsgBox "Error - button not handed!? '" & CalculatorControls(i).CurrentName & "'"
    End Select
  Next i

  PrivateElements.Add This.DisplayTextControl, Nothing
  With DisplayElementSearch
    .Initialise This.DisplayTextControl, PrivateElements(This.LandmarkGroupControl), TreeScope.Descendants
    .AddCondition "ControlTypeIsText", UIAProperties.ControlType, UIAPropertyComparisons.Equals, UIAControlTypeIDs.Text
    .AddCondition "AutomationIdIsCalculatorResults", UIAProperties.AutomationId, UIAPropertyComparisons.Equals, "CalculatorResults"
    .Locator by.pConditions, "AND(ControlTypeIsText, AutomationIdIsCalculatorResults)"
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
      Actions.Click NextButtonName, PrivateElements(NextButtonName)
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
  Set PrivateElements(This.DisplayTextControl) = DisplayElementSearch.Find
  Dim Display As IUIAutomationElement
  Set Display = PrivateElements(This.DisplayTextControl)
  GetDisplay = VBA.Strings.Replace(Display.CurrentName, "Display is ", "")
End Function
