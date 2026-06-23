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
  Elements() As pElement
  RootUIAElement As IUIAutomationElement
  RootUIAElementLocator As pLocator
  RootUIAElementIsDesktop As Boolean
  TreeScope As Long
  FindBy As By
  FindFirst As Boolean
  AllSearchConditions As Scripting.Dictionary
  EvaluationLogic As String
  PositionInMatchingSet  As Integer
End Type

Private This As Properties

Private Const POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER = "PositionOfElementInTreescopeCounter"

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
  Set This.Element.ParentLocator = Me

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
    Case By.AriaRole
      This.FindBy = By.pConditions
      This.EvaluationLogic = "AriaRole"
      Condition This.EvaluationLogic, UIAProperties.AriaRole, UIAPropertyComparisons.IsTheString, EvaluationLogic
    Case By.AutomationId
      This.FindBy = By.pConditions
      This.EvaluationLogic = "AutomationId"
      Condition This.EvaluationLogic, UIAProperties.AutomationId, UIAPropertyComparisons.IsTheString, EvaluationLogic
    Case By.ClassName
      This.FindBy = By.pConditions
      This.EvaluationLogic = "ClassName"
      Condition This.EvaluationLogic, UIAProperties.ClassName, UIAPropertyComparisons.IsTheString, EvaluationLogic
    Case By.NameIs
      This.FindBy = By.pConditions
      This.EvaluationLogic = "NameIs"
      Condition This.EvaluationLogic, UIAProperties.Name, UIAPropertyComparisons.IsTheString, EvaluationLogic
    Case By.ControlType
      This.FindBy = By.pConditions
      This.EvaluationLogic = "ControlType"
      Condition This.EvaluationLogic, UIAProperties.ControlType, UIAPropertyComparisons.EqualsNumber, EvaluationLogic
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
  UIAPropertyValue As Variant, _
  Optional TrimProperty As Boolean = False)
  
  Dim cond As New pCondition
  cond.ConditionName = ConditionName
  cond.UIAProperty = UIAProperty
  cond.UIAPropertyComparison = UIAPropertyComparison
  cond.UIAPropertyValue = UIAPropertyValue
  cond.TrimProperty = TrimProperty
  This.AllSearchConditions.Add ConditionName, cond
  Set cond = Nothing
  
End Sub

Public Sub AriaRole(Role As String, Optional ConditionNameSuffix As String)
  Dim ConditionName As String
  ConditionName = "AriaRole"
  If ConditionNameSuffix <> "" Then: ConditionName = ConditionName & ConditionNameSuffix
  Condition ConditionName, UIAProperties.AriaRole, UIAPropertyComparisons.IsTheString, Role
End Sub

Public Sub AriaRoleBanner()
  AriaRole AriaRoles.Banner, "Banner"
End Sub

Public Sub AriaRoleButton()
  AriaRole AriaRoles.Button, "Button"
End Sub

Public Sub AriaRoleCheckBox()
  AriaRole AriaRoles.CheckBox, "CheckBox"
End Sub

Public Sub AriaRoleComboBox()
  AriaRole AriaRoles.ComboBox, "ComboBox"
End Sub

Public Sub AriaRoleDescription()
  AriaRole AriaRoles.Description, "Description"
End Sub

Public Sub AriaRoleDialog()
  AriaRole AriaRoles.Dialog, "Dialog"
End Sub

Public Sub AriaRoleDocument()
  AriaRole AriaRoles.Document, "Document"
End Sub

Public Sub AriaRoleGeneric()
  AriaRole AriaRoles.Generic, "Generic"
End Sub

Public Sub AriaRoleGroup()
  AriaRole AriaRoles.Group, "Group"
End Sub

Public Sub AriaRoleHeading()
  AriaRole AriaRoles.Heading, "Heading"
End Sub

Public Sub AriaRoleLink()
  AriaRole AriaRoles.Link, "Link"
End Sub

Public Sub AriaRoleList()
  AriaRole AriaRoles.List, "List"
End Sub

Public Sub AriaRoleListItem()
  AriaRole AriaRoles.ListItem, "ListItem"
End Sub

Public Sub AriaRoleMain()
  AriaRole AriaRoles.Main, "Main"
End Sub

Public Sub AriaRoleNavigation()
  AriaRole AriaRoles.Navigation, "Navigation"
End Sub

Public Sub AriaRoleNullString()
  AriaRole AriaRoles.NullString, "NullString"
End Sub

Public Sub AriaRoleRadio()
  AriaRole AriaRoles.Radio, "Radio"
End Sub

Public Sub AriaRoleTextBox()
  AriaRole AriaRoles.TextBox, "TextBox"
End Sub

Public Sub AriaRoleTabPanel()
  AriaRole AriaRoles.TabPanel, "TabPanel"
End Sub

Public Sub AriaRoleToolbar()
  AriaRole AriaRoles.Toolbar, "Toolbar"
End Sub

Public Sub AutomationId(ID As String)
  Condition "AutomationId", UIAProperties.AutomationId, UIAPropertyComparisons.IsTheString, ID
End Sub

Public Sub ClassName(Name As String)
  Condition "ClassName", UIAProperties.ClassName, UIAPropertyComparisons.IsTheString, Name
End Sub

Public Sub ControlType(CtrlType As UIAControlTypeIDs)
  Condition "ControlType", UIAProperties.ControlType, UIAPropertyComparisons.EqualsNumber, CtrlType
End Sub

Public Sub NameIs(Name As String)
  Condition "NameIs", UIAProperties.Name, UIAPropertyComparisons.IsTheString, Name
End Sub

Public Sub NameIs_(Name As String, Optional ConditionNameSuffix As String, Optional TrimProperty As Boolean = False)
  Dim ConditionName As String
  If ConditionNameSuffix = "" Then: ConditionName = "NameIs" & VBA.Replace(Name, " ", ""): Else: ConditionName = "NameIs" & ConditionNameSuffix
  Condition ConditionName, UIAProperties.Name, UIAPropertyComparisons.IsTheString, Name, TrimProperty
End Sub

Public Sub PositionInTreescope(Position As Integer)
  Condition POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, Position
  This.EvaluationLogic = "AND(" & This.EvaluationLogic & ", " & POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
End Sub

Public Sub PositionInMatchingSet(Position As Integer)
  'Use -1 for last item
  This.PositionInMatchingSet = Position
End Sub

Public Sub WindowInteractionState(State As UIAWindowInteractionStates)
  Condition "WindowInteractionState", WindowWindowInteractionState, EqualsNumber, State
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

Public Sub Find(Optional TimeoutInSeconds As Long, Optional AcceptNoElements As Boolean, Optional FindElementAgain As Boolean = False)
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
    'Make sure the root element has been set from it's locator
    Set This.RootUIAElement = This.RootUIAElementLocator.Element.UIAElement
    'Find/Re-Find the Root Element?
    If (This.RootUIAElement Is Nothing) Or (Not This.RootUIAElementLocator.Element.IsAlive()) Then
      This.RootUIAElementLocator.Find TimeoutInSeconds, FindElementAgain:=True
      Set This.RootUIAElement = This.RootUIAElementLocator.Element.UIAElement
    End If
  End If
 
  Toaster.Message "Finding " & This.Element.GivenName, Finding

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
    Toaster.Message "Found " & This.Element.GivenName, Success
    Set This.Element.UIAElement = FoundElements(0).UIAElement
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
  If AllElements Is Nothing Then
    'Could not find any elements below the Root Element - yet!
    Exit Function
  Else
    If AllElements.Length = 0 Then
      'Could not find any elements below the Root Element - yet!
      Exit Function
    End If
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
      If C.ConditionName = POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER Then
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
     CountOfMatchingElements = CountOfMatchingElements + 1
      If This.PositionInMatchingSet = 0 Then
        ReDim Preserve ReturnElements(CountOfMatchingElements - 1)
        Set ReturnElements(CountOfMatchingElements - 1) = CurrentElement
        If This.FindFirst Then
          Exit For
        End If
      Else
        If This.PositionInMatchingSet = -1 Then
          ' Replace the return element with every match so we are left with the last one!
          ReDim Preserve ReturnElements(0)
          Set ReturnElements(0) = CurrentElement
        ElseIf CountOfMatchingElements = This.PositionInMatchingSet Then
          ReDim Preserve ReturnElements(0)
          Set ReturnElements(0) = CurrentElement
          Exit For
        End If
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

Public Sub ListAllChildren()
  ListAllDescendants ChildrenOnly:=True
End Sub
    
'dim TreeWalker As UIAutomationClient.IUIAutomationTreeWalker
'  Dim ParentElement As IUIAutomationElement
'Set ParentElement = This.TreeWalker.GetParentElement(CurrentElement)

Public Sub ListAllDescendants(Optional Level As Integer = 0, Optional ChildrenOnly As Boolean)
  Dim AllElements As IUIAutomationElementArray
  If Element.UIAElement Is Nothing Then
    ErrorLogging.LogError Errors.FindElementsRootElementIsNothing, "List All Children - The root element is nothing!"
    Exit Sub
  Else
    Set AllElements = Element.UIAElement.FindAll(TreeScope.Children, UIA.CreateTrueCondition)
  End If
  If AllElements.Length = 0 Then
    ErrorLogging.LogError Errors.FindElementsFindNoElementsBelowRoot, "Could not find any elements below the Root Element!"
    Exit Sub
  End If
  Dim i As Long
  Dim CurrentElement As IUIAutomationElement
  For i = 0 To AllElements.Length - 1
    Set CurrentElement = AllElements.GetElement(i)
    Debug.Print VBA.Strings.String(1 * 2, " ") & "Element #" & i
    Debug.Print VBA.Strings.String(1 * 2, " ") & "AriaRole: " & CurrentElement.GetCurrentPropertyValue(UIAProperties.AriaRole)
    Debug.Print VBA.Strings.String(1 * 2, " ") & "ControlType: " & CurrentElement.GetCurrentPropertyValue(UIAProperties.ControlType) & " (" & GetControlTypeName(CurrentElement.GetCurrentPropertyValue(UIAProperties.ControlType)) & ")"
    Debug.Print VBA.Strings.String(1 * 2, " ") & "ClassName: " & CurrentElement.GetCurrentPropertyValue(UIAProperties.ClassName)
    Debug.Print VBA.Strings.String(1 * 2, " ") & "Name: " & CurrentElement.GetCurrentPropertyValue(UIAProperties.Name)
    Debug.Print ""
    If Not ChildrenOnly Then
      ListAllDescendants Level
    End If

  Next i
End Sub

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

Public Sub WaitForElementExists(TimeoutInSeconds As Long)

  Dim EndTime As Date
  EndTime = DateAdd("s", TimeoutInSeconds, Now)
   
  Dim PassedEndTime As Boolean
  PassedEndTime = False
  
  While Not (ElementExists Or PassedEndTime)
    PassedEndTime = (Now > EndTime)
    If Not PassedEndTime Then
      Snooze 100
    End If
  Wend

End Sub

