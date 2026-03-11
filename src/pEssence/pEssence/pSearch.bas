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
Option Explicit

'Private pConditions() As pCondition
'Requires a reference Windows Scripting Runtime for Scripting.Dictionary

Private Type Criteria
  ElementName As String
  RootUIAElement As IUIAutomationElement
  TreeScope As Long
  FindBy As pEssence.By
  AllSearchConditions As Scripting.Dictionary
  Locator As String
  FoundUIAElement As IUIAutomationElement
  FoundUIAElements() As IUIAutomationElement
End Type

Private This As Criteria

Private Sub Class_Initialize()
  Set This.AllSearchConditions = New Scripting.Dictionary
End Sub

Public Sub Initialise(ElementName As String, RootUIAElement As IUIAutomationElement, TreeScope As Long)
  This.ElementName = ElementName
  Set This.RootUIAElement = RootUIAElement
  This.TreeScope = TreeScope
End Sub

Public Sub SetTreeScope(UIAElement As IUIAutomationElement)
  Set This.RootUIAElement = UIAElement
End Sub

Public Sub AddCondition( _
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

Public Sub Locator(FindBy As pEssence.By, Locator As String)
  Select Case FindBy
    Case By.pConditions
      This.FindBy = FindBy
      This.Locator = Locator
    Case By.AutomationId
      AddCondition "AutomationIdIs" & Locator, UIAProperties.AutomationId, UIAPropertyComparisons.Equals, Locator
      This.FindBy = By.pConditions
      This.Locator = "AutomationIdIs" & Locator
    Case Else
      pEssence.ErrorLogging.LogError pEssence.Errors.FindElementUnhandledByInLocator, "Unhanded By Locator: " & UIACommon.GetByName(FindBy)
      Exit Sub
  End Select
End Sub

Public Sub Find(Optional TimeoutInMilliseconds As Long, Optional AcceptNoElements As Boolean)
  Dim FoundElements() As IUIAutomationElement
  FoundElements = Findlements(TimeoutInMilliseconds, AcceptNoElements)
  Dim CountOfFoundElements As Integer
  CountOfFoundElements = UBound(FoundElements) + 1
  If CountOfFoundElements > 1 Then
    pEssence.ErrorLogging.LogError pEssence.Errors.FindElementsExpectedOneElementFoundMany, "Expected to find one element but found " & CountOfFoundElements
    Exit Sub
  ElseIf CountOfFoundElements < 1 Then
    pEssence.ErrorLogging.LogError pEssence.Errors.FindElementsExpectedOneElementFoundNone, "Expected to find one element but found none."
    Exit Sub
  End If
  Set This.FoundUIAElement = FoundElements(0)
End Sub

Public Function FindAll(Optional TimeoutInMilliseconds As Long, Optional AcceptNoElements As Boolean) As IUIAutomationElement
  Findlements TimeoutInMilliseconds, AcceptNoElements
End Function

Private Function Findlements(TimeoutInMilliseconds As Long, AcceptNoElements As Boolean) As IUIAutomationElement()

  If Not EvaluationLogicIsOk Then
    pEssence.ErrorLogging.LogError pEssence.Errors.FaultyEvaluationLogicUnspecifiedError, "The evaluation logic is faulty: '" & This.Locator & "'"
    Exit Function
  End If
  
  ' Find the child/descendant elements of the root element
  Dim AllElements As IUIAutomationElementArray
  If This.RootUIAElement Is Nothing Then
    pEssence.ErrorLogging.LogError pEssence.Errors.FindElementsRootElementIsNothing, "The root element is nothing for element '" & This.ElementName & "'!"
    Exit Function
  Else
    Set AllElements = This.RootUIAElement.FindAll(This.TreeScope, pEssence.UIACommon.uiAutomation.CreateTrueCondition)
  End If
  If AllElements.Length = 0 Then
    Debug.Print "Could not find elements below the Root Element"
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
    'Evaluate the formula with Excel
    MatchFound = Excel.Evaluate(CurrentEvaluation)
    'Did we get and error?
    If Err.Number = 0 Then
      'DO nothing!
    Else
      If Err.Number = 13 Then
        On Error GoTo 0
        pEssence.ErrorLogging.LogError _
          pEssence.Errors.FaultyEvaluationLogicUnspecifiedErrorOnEvaluation, _
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
      pEssence.ErrorLogging.LogError pEssence.Errors.FindElementsFindNoElements, "Locator : By " & UIACommon.GetByName(This.FindBy) & " '" & This.Locator & "'"
    End If
  Else
    Findlements = ReturnElements
  End If
  
End Function

Private Function EvaluationLogicIsOk() As Boolean

  Dim RedactedEvaluationLogic As String
  RedactedEvaluationLogic = This.Locator
  
  Dim CountOfLeftBraces As Integer
  Dim CountOfRightBraces As Integer
  CountOfLeftBraces = CountOccurrences(RedactedEvaluationLogic, "(")
  CountOfRightBraces = CountOccurrences(RedactedEvaluationLogic, ")")
  If Not (CountOfLeftBraces = CountOfRightBraces) Then
    EvaluationLogicIsOk = False
    pEssence.ErrorLogging.LogError pEssence.Errors.FaultyEvaluationLogicMismatchBracketsError, "The evaluation logic is faulty, there are mismatchng brackets: '" & This.Locator & "'"
    Exit Function
  End If
  
  Dim k As Variant
  For Each k In This.AllSearchConditions.Keys
    Dim c As New pCondition
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

Private Function CountOccurrences(Text As String, SearchTerm As String) As Integer
  If Len(SearchTerm) = 0 Then
    CountOccurrences = 0
  Else
    CountOccurrences = (VBA.Strings.Len(Text) - VBA.Strings.Len(VBA.Strings.Replace(Text, SearchTerm, "")))
  End If
End Function

Public Sub WaitForPropertyValue(UIAProperty As UIAProperties, UIAPropertyValue As Variant, Optional TimeoutInMilliseconds As Long)
  Dim CurrentPropertyValue As Variant
  CurrentPropertyValue = pEssence.GetProperty(This.FoundUIAElement, UIAProperty)
  If CurrentPropertyValue = UIAPropertyValue Then
    'Success - exit here!
    Exit Sub
  Else
    MsgBox "Need to wait for the value before timeout here!"
  End If
End Sub

Public Function ElementName() As String
  ElementName = This.ElementName
End Function

Public Function FoundUIAElement() As IUIAutomationElement
  Set FoundUIAElement = This.FoundUIAElement
End Function

'Dim pattern As IUIAutomationInvokePattern
'Set pattern = GetPattern(element, UIA_InvokePatternId)

