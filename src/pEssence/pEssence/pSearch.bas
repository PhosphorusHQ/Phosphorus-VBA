VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder pEssence
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit
'Requires a reference Windows Scripting Runtime for Scripting.Dictionary

Private Type Criteria
  ElementName As String
  RootUIAElement As IUIAutomationElement
  TreeScope As Long
  FindBy As By
  AllSearchConditions As Scripting.Dictionary
  Locator As String
End Type

Private This As Criteria

Private Sub Class_Initialize()
  Set This.AllSearchConditions = New Scripting.Dictionary
End Sub

Public Sub Initialise( _
  ElementName As String, _
  RootUIAElement As IUIAutomationElement, _
  TreeScope As Long, _
  FindBy As By, _
  Locator As String)
  
  This.ElementName = ElementName
  Set This.RootUIAElement = RootUIAElement
  This.TreeScope = TreeScope
  Select Case FindBy
    Case By.pConditions
      This.FindBy = FindBy
      This.Locator = Locator
    Case By.AutomationId
      Condition "AutomationIdIs" & Locator, UIAProperties.AutomationId, UIAPropertyComparisons.Equals, Locator
      This.FindBy = By.pConditions
      This.Locator = "AutomationIdIs" & Locator
    Case Else
      ErrorLogging.LogError Errors.FindElementUnhandledByInLocator, "Unhanded By Locator: " & UIACommon.GetByName(FindBy)
      Exit Sub
  End Select

End Sub

Public Sub SetTreeScope(UIAElement As IUIAutomationElement)
  Set This.RootUIAElement = UIAElement
End Sub

Public Sub Condition( _
  ConditionName As String, _
  UIAProperty As UIAProperties, _
  UIAPropertyComparison As UIAPropertyComparisons, _
  UIAPropertyValue As Variant)
  
  Dim cond As New pCondition
  cond.ConditionName = ConditionName
  cond.UIAProperty = UIAProperty
  cond.UIAPropertyComparison = UIAPropertyComparison
  cond.UIAPropertyValue = UIAPropertyValue
  This.AllSearchConditions.Add ConditionName, cond
  Set cond = Nothing
  
End Sub

Public Sub ListAllConditions()
  Dim k As Variant
  For Each k In This.AllSearchConditions.Keys
    Dim c As New pCondition
    Set c = This.AllSearchConditions(k)
    Debug.Print c.ConditionName; c.UIAProperty; c.UIAPropertyComparison, c.UIAPropertyValue
  Next k
End Sub

Public Function Find(Optional TimeoutInSeconds As Long, Optional AcceptNoElements As Boolean, Optional FindElementAgain As Boolean = True) As IUIAutomationElement
'Find a single element with a timeout

  Dim FoundElements() As IUIAutomationElement
  Dim CountOfFoundElements As Integer
  
  'Calculate the end time
  Dim EndTime As Date
  EndTime = DateAdd("s", TimeoutInSeconds, Now)
    
  'Loop until element(s) found or timed out
  Dim SomeElementsFound As Boolean
  SomeElementsFound = False
  
  Dim PassedEndTime As Boolean
  PassedEndTime = False
  
  While Not (SomeElementsFound Or PassedEndTime)
    FoundElements = Findlements(True)
    CountOfFoundElements = 0
    If Not UIACommon.IsArrayEmpty(FoundElements) Then
      CountOfFoundElements = UBound(FoundElements) + 1
    End If
    SomeElementsFound = (CountOfFoundElements > 0)
    If Not SomeElementsFound Then
      PassedEndTime = (Now > EndTime)
      If Not PassedEndTime Then
        WindowsProcesses.Snooze 100
      End If
    End If
  Wend
    
  If CountOfFoundElements > 1 Then
    ErrorLogging.LogError Errors.FindElementsExpectedOneElementFoundMany, "Expected to find one element but found " & CountOfFoundElements
    Exit Function
  ElseIf CountOfFoundElements = 0 And Not AcceptNoElements Then
    ErrorLogging.LogError _
      Errors.FindElementsExpectedOneElementFoundNone, _
      "Expected to find one element but found none." & vbCrLf & vbCrLf & _
      "Looking for '" & This.ElementName & "' element, locator: '" & This.Locator & "', timeout (" & TimeoutInSeconds & " seconds)"
    Exit Function
  End If
  
  Dim FoundElement As IUIAutomationElement
  If CountOfFoundElements = 1 Then
    Set FoundElement = FoundElements(0)
    If FindElementAgain Then
      If UIAProps.HasProperty(This.ElementName, FoundElement, UIAProperties.IsScrollItemPatternAvailable) Then
        Actions.TryToScrollItemIntoView This.ElementName, FoundElement
        Set Find = Find(TimeoutInSeconds, AcceptNoElements, FindElementAgain:=False)
      End If
    End If
    Window.HighlightElement FoundElements(0), BorderColor:=&H808000 'Cyan
    Window.ReleaseHighlighting
    Set Find = FoundElements(0)
  End If

End Function

Public Function FindAll(Optional AcceptNoElements As Boolean) As IUIAutomationElement()
'Find more than 1 elements with no timeout
  FindAll = Findlements(AcceptNoElements)
End Function

Private Function Findlements(AcceptNoElements As Boolean) As IUIAutomationElement()
'Find 1 or more elements with no timeout

  If Not EvaluationLogicIsOk Then
    ErrorLogging.LogError Errors.FaultyEvaluationLogicUnspecifiedError, "The evaluation logic is faulty: '" & This.Locator & "'"
    Exit Function
  End If
  
  ' Find the child/descendant elements of the root element
  Dim AllElements As IUIAutomationElementArray
  If This.RootUIAElement Is Nothing Then
    ErrorLogging.LogError Errors.FindElementsRootElementIsNothing, "The root element is nothing for element '" & This.ElementName & "'!"
    Exit Function
  Else
    Set AllElements = This.RootUIAElement.FindAll(This.TreeScope, UIA.CreateTrueCondition)
  End If
  If AllElements.Length = 0 Then
    ErrorLogging.LogError Errors.FindElementsFindNoElementsBelowRoot, "Could not find any elements below the Root Element!"
    Exit Function
  End If

  Dim ReturnElements() As IUIAutomationElement
  Dim CountOfMatchingElements As Integer
  CountOfMatchingElements = 0

  Dim i As Long
  Dim CurrentElement As IUIAutomationElement
  For i = 0 To AllElements.Length - 1
    Set CurrentElement = AllElements.GetElement(i)
 
    Dim CurrentEvaluation  As String
    CurrentEvaluation = This.Locator
    
    Dim k As Variant
    For Each k In This.AllSearchConditions.Keys
      Dim c As New pCondition
      Set c = This.AllSearchConditions(k)
      CurrentEvaluation = VBA.Strings.Replace(CurrentEvaluation, c.ConditionName, c.Evaluate(CurrentElement))
    Next k

    Dim MatchFound As Boolean
    On Error Resume Next
    MatchFound = Excel.Evaluate(CurrentEvaluation) 'Evaluate the formula with Excel!
    'Did we get and error?
    If Err.Number = 0 Then
      'Do nothing!
    Else
      If Err.Number = 13 Then
        On Error GoTo 0
        ErrorLogging.LogError _
          Errors.FaultyEvaluationLogicUnspecifiedErrorOnEvaluation, _
          "The evaluation logic did not evaluate to TRUE or FALSE:" & vbCrLf & vbCrLf & _
          "'" & This.Locator & "' returned '" & VBA.Conversion.CStr(Excel.Evaluate(CurrentEvaluation)) & "'"
        Exit Function
      Else
        Err.Raise Err.Number, Err.Description
      End If
    End If
    On Error GoTo 0
    If MatchFound Then
      ReDim Preserve ReturnElements(CountOfMatchingElements)
      Set ReturnElements(CountOfMatchingElements) = CurrentElement
      CountOfMatchingElements = CountOfMatchingElements + 1
    End If
  Next i

  If UIACommon.IsArrayEmpty(ReturnElements) Then
    If AcceptNoElements Then
      'Find elements is nothing?
    Else
      ErrorLogging.LogError Errors.FindElementsFindNoElements, "Locator : By " & UIACommon.GetByName(This.FindBy) & " '" & This.Locator & "'"
    End If
  Else
    Findlements = ReturnElements
  End If
  
End Function

Private Function EvaluationLogicIsOk() As Boolean
  
  Dim CountOfLeftBraces As Integer
  Dim CountOfRightBraces As Integer
  CountOfLeftBraces = Utils.CountOccurrences(This.Locator, "(")
  CountOfRightBraces = Utils.CountOccurrences(This.Locator, ")")
  If Not (CountOfLeftBraces = CountOfRightBraces) Then
    EvaluationLogicIsOk = False
    ErrorLogging.LogError Errors.FaultyEvaluationLogicMismatchBracketsError, "The evaluation logic is faulty, there are mismatchng brackets: '" & This.Locator & "'"
    Exit Function
  End If
  
  Dim k As Variant
  Dim c As New pCondition
  For Each k In This.AllSearchConditions.Keys
    Set c = This.AllSearchConditions(k)
    If VBA.Strings.InStr(This.Locator, c.ConditionName) = 0 Then
      EvaluationLogicIsOk = False
      ErrorLogging.LogError Errors.FaultyEvaluationLogicConditionIsNotUsed, "The Condition '" & c.ConditionName & "' is not used in the locator '" & This.Locator & "'"
      Exit Function
    End If
  Next k
    
  Dim RedactedEvaluationLogic As String
  RedactedEvaluationLogic = This.Locator
    
  For Each k In This.AllSearchConditions.Keys
    Set c = This.AllSearchConditions(k)
    RedactedEvaluationLogic = VBA.Strings.Replace(RedactedEvaluationLogic, c.ConditionName, "")
  Next k
    
  RedactedEvaluationLogic = VBA.Strings.UCase(RedactedEvaluationLogic)
  RedactedEvaluationLogic = VBA.Strings.Replace(RedactedEvaluationLogic, " ", "")
  RedactedEvaluationLogic = VBA.Strings.Replace(RedactedEvaluationLogic, ",", "")
  
  'Loop to remove all functions until no change in length
  Dim PreviousLength As Integer
  Dim NewLength As Integer
  PreviousLength = 0
  NewLength = VBA.Strings.Len(RedactedEvaluationLogic)
  While NewLength <> PreviousLength
    PreviousLength = NewLength
    RedactedEvaluationLogic = VBA.Strings.Replace(RedactedEvaluationLogic, "AND()", "")
    RedactedEvaluationLogic = VBA.Strings.Replace(RedactedEvaluationLogic, "OR()", "")
    RedactedEvaluationLogic = VBA.Strings.Replace(RedactedEvaluationLogic, "NOT()", "")
    NewLength = VBA.Strings.Len(RedactedEvaluationLogic)
  Wend
  
  EvaluationLogicIsOk = (RedactedEvaluationLogic = "")

End Function

Public Function ElementName() As String
  ElementName = This.ElementName
End Function
