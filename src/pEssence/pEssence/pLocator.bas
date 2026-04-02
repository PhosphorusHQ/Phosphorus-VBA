VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pLocator"
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

Private Type Properties
  ElementName As String
  RootUIAElement As IUIAutomationElement
  RootUIAElementLocator As pLocator
  RootUIAElementIsDesktop As Boolean
  TreeScope As Long
  FindBy As By
  FindFirst As Boolean
  AllSearchConditions As Scripting.Dictionary
  EvaluationLogic As String
  FoundUIAElement As IUIAutomationElement
  FoundUIAElements() As IUIAutomationElement
End Type

Private This As Properties

Private Sub Class_Initialize()
  Set This.AllSearchConditions = New Scripting.Dictionary
End Sub

Public Sub Initialise( _
  ElementName As String, _
  RootUIAElementLocator As pLocator, _
  TreeScope As Long, _
  FindBy As By, _
  EvaluationLogic As String, _
  Optional FindFirst As Boolean)
  
  This.ElementName = ElementName
  If RootUIAElementLocator Is Nothing Then
    Set This.RootUIAElement = Factory.GetRootDesktopElement
    This.RootUIAElementIsDesktop = True
  Else
    Set This.RootUIAElementLocator = RootUIAElementLocator
    Set This.RootUIAElement = This.RootUIAElementLocator.FoundUIAElement
  End If
  This.TreeScope = TreeScope
  Select Case FindBy
    Case By.pConditions
      This.FindBy = FindBy
      This.EvaluationLogic = EvaluationLogic
    Case By.AutomationId
      This.FindBy = By.pConditions
      This.EvaluationLogic = "AutomationId"
      Condition This.EvaluationLogic, UIAProperties.AutomationId, UIAPropertyComparisons.IsTheString, EvaluationLogic
    Case Else
      ErrorLogging.LogError Errors.FindElementUnhandledByInLocator, "Unhanded By Locator: " & UIACommon.GetByName(FindBy)
      Exit Sub
  End Select
  This.FindFirst = FindFirst
  
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

Public Function ElementName() As String
  ElementName = This.ElementName
End Function

Public Function FoundUIAElement() As IUIAutomationElement
  Set FoundUIAElement = This.FoundUIAElement
End Function

Public Function FoundUIAElements() As IUIAutomationElement()
  FoundUIAElements = This.FoundUIAElements
End Function

Public Sub Find(Optional TimeoutInSeconds As Long, Optional AcceptNoElements As Boolean, Optional FindElementAgain As Boolean = True)
'Find a single element with a timeout

  If Not This.FoundUIAElement Is Nothing And Actions.IsElementAlive(This.ElementName, This.FoundUIAElement) And Not FindElementAgain Then
    'The element has already been found, no need to find it again
    Exit Sub
  End If
  
  'Make sure we have the root element before finding the current element
  If This.RootUIAElementIsDesktop Then
    If This.RootUIAElementLocator Is Nothing Then
      Set This.FoundUIAElement = Factory.GetRootDesktopElement
'      Exit Sub
    End If
  Else
    'Find/Re-Find the Root Element
    This.RootUIAElementLocator.Find TimeoutInSeconds
    Set This.RootUIAElement = This.RootUIAElementLocator.FoundUIAElement
  End If

  Toaster.Message "Finding " & This.ElementName, Finding

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
    Exit Sub
  ElseIf CountOfFoundElements = 0 And Not AcceptNoElements Then
    ErrorLogging.LogError _
      Errors.FindElementsExpectedOneElementFoundNone, _
      "Expected to find one element but found none." & vbCrLf & vbCrLf & _
      "Looking for '" & This.ElementName & "' element, evaluation logic: '" & This.EvaluationLogic & "', timeout (" & TimeoutInSeconds & " seconds)"
    Exit Sub
  End If
  
  Dim FoundElement As IUIAutomationElement
  If CountOfFoundElements = 1 Then
    Set This.FoundUIAElement = FoundElements(0)
    If FindElementAgain Then
      If UIAProps.HasProperty(This.ElementName, This.FoundUIAElement, UIAProperties.IsScrollItemPatternAvailable) Then
        Actions.TryToScrollItemIntoView This.ElementName, This.FoundUIAElement
        Find TimeoutInSeconds, AcceptNoElements, FindElementAgain:=False
      End If
    End If
    Window.HighlightElement FoundElements(0), BorderColor:=&H808000 'Cyan
    Window.ReleaseHighlighting
  End If

End Sub

Public Sub FindAll(Optional AcceptNoElements As Boolean)
'Find more than 1 elements with no timeout
  Toaster.Message "Finding all elements " & This.ElementName, Finding
  This.FoundUIAElements = Findlements(AcceptNoElements)
End Sub

Private Function Findlements(AcceptNoElements As Boolean) As IUIAutomationElement()
'Find 1 or more elements with no timeout

  If Not EvaluationLogicIsOk Then
    ErrorLogging.LogError Errors.FaultyEvaluationLogicUnspecifiedError, "The evaluation logic is faulty: '" & This.EvaluationLogic & "'"
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
 
    Dim CurrentEvaluationLogic  As String
    CurrentEvaluationLogic = This.EvaluationLogic
    
    Dim k As Variant
    For Each k In This.AllSearchConditions.Keys
      Dim c As New pCondition
      Set c = This.AllSearchConditions(k)
      If c.ConditionName = UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER Then
        CurrentEvaluationLogic = VBA.Strings.Replace(CurrentEvaluationLogic, c.ConditionName, ((i + 1) = c.UIAPropertyValue))
      Else
        CurrentEvaluationLogic = VBA.Strings.Replace(CurrentEvaluationLogic, c.ConditionName, c.Evaluate(CurrentElement))
      End If
    Next k

    Dim MatchFound As Boolean
    On Error Resume Next
    MatchFound = Excel.Evaluate(CurrentEvaluationLogic) 'Evaluate the formula with Excel!
    'Did we get and error?
    If Err.Number = 0 Then
      'Do nothing!
    Else
      If Err.Number = 13 Then
        On Error GoTo 0
        ErrorLogging.LogError _
          Errors.FaultyEvaluationLogicUnspecifiedErrorOnEvaluation, _
          "The evaluation logic did not evaluate to TRUE or FALSE:" & vbCrLf & vbCrLf & _
          "'" & This.EvaluationLogic & "' returned '" & VBA.Conversion.CStr(Excel.Evaluate(CurrentEvaluationLogic)) & "'"
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
      If This.FindFirst Then
        Exit For
      End If
    End If
  Next i

  If UIACommon.IsArrayEmpty(ReturnElements) Then
    If AcceptNoElements Then
      'Find elements is nothing?
    Else
      ErrorLogging.LogError Errors.FindElementsFindNoElements, "Locator : By " & UIACommon.GetByName(This.FindBy) & " '" & This.EvaluationLogic & "'"
    End If
  Else
    Findlements = ReturnElements
  End If
  
End Function

Private Function EvaluationLogicIsOk() As Boolean
  
  Dim CountOfLeftBraces As Integer
  Dim CountOfRightBraces As Integer
  CountOfLeftBraces = Utils.CountOccurrences(This.EvaluationLogic, "(")
  CountOfRightBraces = Utils.CountOccurrences(This.EvaluationLogic, ")")
  If Not (CountOfLeftBraces = CountOfRightBraces) Then
    EvaluationLogicIsOk = False
    ErrorLogging.LogError Errors.FaultyEvaluationLogicMismatchBracketsError, "The evaluation logic is faulty, there are mismatchng brackets: '" & This.EvaluationLogic & "'"
    Exit Function
  End If
  
  Dim k As Variant
  Dim c As New pCondition
  For Each k In This.AllSearchConditions.Keys
    Set c = This.AllSearchConditions(k)
    If VBA.Strings.InStr(This.EvaluationLogic, c.ConditionName) = 0 Then
      EvaluationLogicIsOk = False
      ErrorLogging.LogError Errors.FaultyEvaluationLogicConditionIsNotUsed, "The Condition '" & c.ConditionName & "' is not used in the locator '" & This.EvaluationLogic & "'"
      Exit Function
    End If
  Next k
    
  Dim RedactedEvaluationLogic As String
  RedactedEvaluationLogic = This.EvaluationLogic
    
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

Public Function ElementExists(Optional TimeoutInSeconds As Long)
  Set This.FoundUIAElement = Nothing
  Find TimeoutInSeconds, AcceptNoElements:=True, FindElementAgain:=True
  ElementExists = Not (This.FoundUIAElement Is Nothing)
End Function

Public Function ElementDoesntExist(Optional TimeoutInSeconds As Long)
  ElementDoesntExist = Not ElementExists(TimeoutInSeconds)
End Function

