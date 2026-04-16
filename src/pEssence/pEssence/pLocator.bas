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
  Initialised As Boolean
  Element As pElement
  Elements() As pElement 'IUIAutomationElement
  RootUIAElement As IUIAutomationElement
  RootUIAElementLocator As pLocator
  RootUIAElementIsDesktop As Boolean
  TreeScope As Long
  FindBy As By
  FindFirst As Boolean
  AllSearchConditions As Scripting.Dictionary
  EvaluationLogic As String
End Type

Private This As Properties

Private Sub Class_Initialize()
  Set This.AllSearchConditions = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
  Set This.Element = Nothing
End Sub

Public Sub Initialise( _
  ElementName As String, _
  RootUIAElementLocator As pLocator, _
  TreeScope As Long, _
  FindBy As By, _
  EvaluationLogic As String, _
  Optional FindFirst As Boolean)
  
  Set This.Element = New pElement
  This.Element.GivenName = ElementName
  
  If RootUIAElementLocator Is Nothing Then
    Set This.RootUIAElement = Factory.GetRootDesktopElement
    This.RootUIAElementIsDesktop = True
  Else
    Set This.RootUIAElementLocator = RootUIAElementLocator
    Set This.RootUIAElement = This.RootUIAElementLocator.Element.UIAElement
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
  
  This.Initialised = True
  
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

Public Sub AutomationId(ID As String)
  Condition "AutomationId", UIAProperties.AutomationId, UIAPropertyComparisons.IsTheString, ID
End Sub

Public Sub ListAllConditions()
  Dim k As Variant
  For Each k In This.AllSearchConditions.Keys
    Dim C As New pCondition
    Set C = This.AllSearchConditions(k)
    Debug.Print C.ConditionName; C.UIAProperty; C.UIAPropertyComparison, C.UIAPropertyValue
  Next k
End Sub

Public Function Initialised() As Boolean
  Initialised = This.Initialised
End Function

Public Function Element() As pElement
  Set Element = This.Element
End Function

Public Function Elements() As pElement()
  Elements = This.Elements
End Function

Public Sub Find(Optional TimeoutInSeconds As Long, Optional AcceptNoElements As Boolean, Optional FindElementAgain As Boolean = True)
'Find a single element with a timeout

  If Not This.Element.UIAElement Is Nothing And This.Element.IsAlive() And Not FindElementAgain Then
    'The element has already been found, no need to find it again
    Exit Sub
  End If
  
  'Make sure we have the root element before finding the current element
  If This.RootUIAElementIsDesktop Then
    If This.RootUIAElementLocator Is Nothing Then
      Set This.Element.UIAElement = Factory.GetRootDesktopElement
    End If
  Else
    'Find/Re-Find the Root Element
    This.RootUIAElementLocator.Find TimeoutInSeconds
    Set This.RootUIAElement = This.RootUIAElementLocator.Element.UIAElement
  End If

  Toaster.Message "Finding " & This.Element.GivenName, Finding

'  Dim FoundElements() As IUIAutomationElement
  Dim FoundElements() As pElement
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
    If Not pEssence.Utils.IsArrayEmpty(FoundElements) Then
      CountOfFoundElements = UBound(FoundElements) + 1
    End If
    SomeElementsFound = (CountOfFoundElements > 0)
    If Not SomeElementsFound Then
      PassedEndTime = (Now > EndTime)
      If Not PassedEndTime Then
        Snooze 100
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
      "Looking for '" & This.Element.GivenName & "' element, evaluation logic: '" & This.EvaluationLogic & "', timeout (" & TimeoutInSeconds & " seconds)"
    Exit Sub
  End If
  
  Dim FoundElement As IUIAutomationElement
  If CountOfFoundElements = 1 Then
    Set This.Element = FoundElements(0)
    If FindElementAgain Then
      If This.Element.HasProperty(UIAProperties.IsScrollItemPatternAvailable) Then
        Actions.TryToScrollItemIntoView This.Element
        Find TimeoutInSeconds, AcceptNoElements, FindElementAgain:=False
      End If
    End If
    Window.HighlightElement FoundElements(0).UIAElement, BorderColor:=&H808000 'Cyan
    Window.ReleaseHighlighting
  End If

End Sub

Public Sub FindAll(Optional AcceptNoElements As Boolean)
'Find more than 1 elements with no timeout
  Toaster.Message "Finding all elements " & This.Element.GivenName, Finding
  This.Elements = Findlements(AcceptNoElements)
End Sub

Private Function Findlements(AcceptNoElements As Boolean) As pElement() ' IUIAutomationElement()
'Find 1 or more elements with no timeout

  If Not EvaluationLogicIsOk Then
    ErrorLogging.LogError Errors.FaultyEvaluationLogicUnspecifiedError, "The evaluation logic is faulty: '" & This.EvaluationLogic & "'"
    Exit Function
  End If
  
  ' Find the child/descendant elements of the root element
  Dim AllElements As IUIAutomationElementArray
  If This.RootUIAElement Is Nothing Then
    ErrorLogging.LogError Errors.FindElementsRootElementIsNothing, "The root element is nothing for element '" & This.Element.GivenName & "'!"
    Exit Function
  Else
    Set AllElements = This.RootUIAElement.FindAll(This.TreeScope, UIA.CreateTrueCondition)
  End If
  If AllElements.Length = 0 Then
    ErrorLogging.LogError Errors.FindElementsFindNoElementsBelowRoot, "Could not find any elements below the Root Element!"
    Exit Function
  End If

  Dim ReturnElements() As pElement
  Dim CountOfMatchingElements As Integer
  CountOfMatchingElements = 0

  Dim i As Long
  Dim CurrentElement As pElement
  For i = 0 To AllElements.Length - 1
    
    Set CurrentElement = Factory.GetNewElement("Found Element #" & (i + 1), AllElements.GetElement(i))

    Dim CurrentEvaluationLogic  As String
    CurrentEvaluationLogic = This.EvaluationLogic
    
    Dim k As Variant
    For Each k In This.AllSearchConditions.Keys
      Dim C As New pCondition
      Set C = This.AllSearchConditions(k)
      If C.ConditionName = UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER Then
        CurrentEvaluationLogic = VBA.Strings.Replace(CurrentEvaluationLogic, C.ConditionName, ((i + 1) = C.UIAPropertyValue))
      Else
        CurrentEvaluationLogic = VBA.Strings.Replace(CurrentEvaluationLogic, C.ConditionName, C.Evaluate(CurrentElement))
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

  If pEssence.Utils.IsArrayEmpty(ReturnElements) Then
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
  Dim C As New pCondition
  For Each k In This.AllSearchConditions.Keys
    Set C = This.AllSearchConditions(k)
    If VBA.Strings.InStr(This.EvaluationLogic, C.ConditionName) = 0 Then
      EvaluationLogicIsOk = False
      ErrorLogging.LogError Errors.FaultyEvaluationLogicConditionIsNotUsed, "The Condition '" & C.ConditionName & "' is not used in the locator '" & This.EvaluationLogic & "'"
      Exit Function
    End If
  Next k
    
  Dim RedactedEvaluationLogic As String
  RedactedEvaluationLogic = This.EvaluationLogic
    
  For Each k In This.AllSearchConditions.Keys
    Set C = This.AllSearchConditions(k)
    RedactedEvaluationLogic = VBA.Strings.Replace(RedactedEvaluationLogic, C.ConditionName, "")
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
    RedactedEvaluationLogic = VBA.Strings.Replace(RedactedEvaluationLogic, "AND(", "")
    RedactedEvaluationLogic = VBA.Strings.Replace(RedactedEvaluationLogic, "OR(", "")
    RedactedEvaluationLogic = VBA.Strings.Replace(RedactedEvaluationLogic, "NOT(", "")
    RedactedEvaluationLogic = VBA.Strings.Replace(RedactedEvaluationLogic, ")", "")
    NewLength = VBA.Strings.Len(RedactedEvaluationLogic)
  Wend
  
  EvaluationLogicIsOk = (RedactedEvaluationLogic = "")

End Function

Public Function ElementExists(Optional TimeoutInSeconds As Long)
  Set This.Element.UIAElement = Nothing
  Find TimeoutInSeconds, AcceptNoElements:=True, FindElementAgain:=True
  ElementExists = Not (This.Element.UIAElement Is Nothing)
End Function

Public Function ElementDoesntExist(Optional TimeoutInSeconds As Long)
  ElementDoesntExist = Not ElementExists(TimeoutInSeconds)
End Function

