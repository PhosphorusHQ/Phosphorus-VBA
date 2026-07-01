VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Core"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  '@Folder pPath
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Private lintNumberOfXPathExpressions As Integer

Private This As pPath.Common

'Interim Processing Variables
Private lstrRemainingXpath As String
Private lstrNextAxis As String
Private lstrNextNodeTest As String
Private leleContextNodes() As UIAutomationClient.IUIAutomationElement
Private lstrInitialPPaths() As String

Public Sub Initialise()
  Set This = New pPath.Common
  Erase leleContextNodes
  Erase lstrInitialPPaths
  If This.UIAutomation Is Nothing Then
    Set This.UIAutomation = New CUIAutomation
    'Set initial root element as the whole Desktop
    Set This.ApplicationRootUIElement = This.UIAutomation.GetRootElement
  End If
  Set This.TreeWalker = This.UIAutomation.ControlViewWalker
  If This.AutomationDictionaries Is Nothing Then
    Set This.AutomationDictionaries = New UIAutomationDictionaries
    This.AutomationDictionaries.Initialise
  End If
  Set This.Steps = New pPath.Steps
  Set This.Axes = New pPath.Axes
  Set This.NodeTests = New pPath.NodeTests
  Set This.Predicates = New pPath.Predicates
  This.Steps.Initialise This
  This.Axes.Initialise This
  This.NodeTests.Initialise This
  This.Predicates.Initialise This
  pPath.ExcelTopLevelFunctions.InitialiseTopLevelFunctions
End Sub

Public Sub SetApplicationRootElement(eleNewRootUIElement As UIAutomationClient.IUIAutomationElement)
  Set This.ApplicationRootUIElement = eleNewRootUIElement
End Sub

Public Sub AddContextNode(ContextNode As UIAutomationClient.IUIAutomationElement, InitialpPath As String, Optional FirstNode As Boolean)
  If FirstNode Then
    Erase leleContextNodes
    Erase lstrInitialPPaths
  End If
  Dim i As Integer
  i = Phosphorus.Utils.GetSizeOfArray(leleContextNodes()) + 1
  If i = 1 Then
    ReDim leleContextNodes(i)
    ReDim lstrInitialPPaths(i)
  Else
    ReDim Preserve leleContextNodes(i)
    ReDim Preserve lstrInitialPPaths(i)
  End If
  Set leleContextNodes(i) = ContextNode
  lstrInitialPPaths(i) = InitialpPath
End Sub

Public Sub InitialiseEvaluation(intNumberOfXPathExpressions As Integer)
  pPath.ConstantsAndStatic.InitialiseAxesTypes
  Set This.pPathReturnClass = Nothing
  Set This.pPathReturnClass = pPath.ConstantsAndStatic.GetNewPhosphorusPPathReturnClass(intNumberOfXPathExpressions)
End Sub

'See: https://www.w3schools.com/xml/xpath_syntax.asp

'XPath Set operators
'https://infocenter.sybase.com/help/index.jsp?topic=/com.sybase.dc30020_1251/html/xmlb/xmlb32.htm

Public Function Evaluate( _
  ByVal FullLocationPathExpression As String, _
  Optional ContextNode As UIAutomationClient.IUIAutomationElement, _
  Optional InitialpPath As String, _
  Optional UnitTestingMode As Boolean) As pPath.ReturnClass

  Logger.Info "Evaluating PPath: " & FullLocationPathExpression, Internal

  Logger.Trace "Checking for Context Node", Internal
  If Not ContextNode Is Nothing Then 'An InitialPPath should also be provided
    Logger.Debugging "Adding first context node: " & InitialpPath, Internal
    AddContextNode ContextNode, InitialpPath, True
  Else
    Logger.Debugging "Initial pPath: " & VBA.Interaction.IIf(InitialpPath = "", "{null}", InitialpPath), Internal
  End If

  This.UnitTestingMode = UnitTestingMode

  'Extract any top level node set functions
  Dim strTopLevelFunctionPrefix As String
  Dim strTopLevelFunctionSuffix As String
  strTopLevelFunctionPrefix = ""
  strTopLevelFunctionSuffix = ""

  Dim intTopLevelFunctionArrayCounter As Integer
  Dim boolTopLEvelFunctionSet As Boolean
  boolTopLEvelFunctionSet = True
  While boolTopLEvelFunctionSet
    boolTopLEvelFunctionSet = False

    Dim strCurrentTopLevelFunction As String
    Dim intNumberOfAddtionalParameters As Integer

    Dim TopLevelFunction As Collection
    For Each TopLevelFunction In pPath.ExcelTopLevelFunctions.TopLevelFunctions
      strCurrentTopLevelFunction = TopLevelFunction("FunctionName")
      intNumberOfAddtionalParameters = TopLevelFunction("NumberOfAddtionalParameters")
      If (VBA.Strings.InStr(1, FullLocationPathExpression, strCurrentTopLevelFunction & "(") = 1) And (VBA.Strings.Right(FullLocationPathExpression, 1) = ")") Then
        strTopLevelFunctionPrefix = strTopLevelFunctionPrefix & strCurrentTopLevelFunction & "("
        If intNumberOfAddtionalParameters = 0 Then
          strTopLevelFunctionSuffix = ")" & strTopLevelFunctionSuffix
          FullLocationPathExpression = VBA.Strings.Mid(FullLocationPathExpression, VBA.Strings.Len(strCurrentTopLevelFunction & "(") + 1, VBA.Strings.Len(FullLocationPathExpression) - VBA.Strings.Len(strCurrentTopLevelFunction & ")") - 1)
        Else
'Works for 1 additional parameter only!!
          strTopLevelFunctionSuffix = strTopLevelFunctionSuffix & VBA.Strings.Mid(FullLocationPathExpression, VBA.Strings.InStrRev(FullLocationPathExpression, ","))
          FullLocationPathExpression = VBA.Strings.Mid(FullLocationPathExpression, VBA.Strings.Len(strCurrentTopLevelFunction & "(") + 1, VBA.Strings.Len(FullLocationPathExpression) - VBA.Strings.Len(strCurrentTopLevelFunction) - VBA.Strings.Len(strTopLevelFunctionSuffix) - 1)
        End If
        boolTopLEvelFunctionSet = True
        Exit For
      End If
    Next TopLevelFunction

  Wend

  Dim strAllXPathExpressions() As String
  strAllXPathExpressions = This.Steps.GetAllXPathExpressions(FullLocationPathExpression)
  lintNumberOfXPathExpressions = UBound(strAllXPathExpressions)

  InitialiseEvaluation lintNumberOfXPathExpressions
  This.pPathReturnClass.TopLevelFunctionPrefix = strTopLevelFunctionPrefix
  This.pPathReturnClass.TopLevelFunctionSuffix = strTopLevelFunctionSuffix

  If Not PreValidateFullPPathExpression(FullLocationPathExpression) Then
    GoTo ExitFunction:
  End If

  If Not PreValidateSplitPPathExpression(strAllXPathExpressions, leleContextNodes) Then
    GoTo ExitFunction:
  End If

  Dim strCurrentLocationPathExpression As String
  Dim intLocationPathExpressionCounter As Integer

'TODO: This for loop could now be incorporated below!?
  For intLocationPathExpressionCounter = 1 To lintNumberOfXPathExpressions
    strCurrentLocationPathExpression = strAllXPathExpressions(intLocationPathExpressionCounter)
    If (strCurrentLocationPathExpression <> "") And _
       (VBA.Strings.Left(strCurrentLocationPathExpression, 2) <> "//") And _
       (VBA.Strings.Left(strCurrentLocationPathExpression, 1) <> "/") Then
       'We have a relative PPath
      If (VBA.Strings.Left(strCurrentLocationPathExpression, 2) <> "./") Then
        'Set the default relative axis for relative expressions
        strAllXPathExpressions(intLocationPathExpressionCounter) = "./" & strAllXPathExpressions(intLocationPathExpressionCounter)
      End If
    End If
  Next intLocationPathExpressionCounter

  Dim intNumberOfContextNodes As Integer
  intNumberOfContextNodes = Phosphorus.Utils.GetSizeOfArray(leleContextNodes)

  For intLocationPathExpressionCounter = 1 To lintNumberOfXPathExpressions

    This.CurrentLocationPathExpressionCounter = intLocationPathExpressionCounter
    strCurrentLocationPathExpression = strAllXPathExpressions(This.CurrentLocationPathExpressionCounter)

    Dim CurrentContextNode As UIAutomationClient.IUIAutomationElement
    Dim strInitialPPathForCurrentExpression As String
    If VBA.Strings.Left(strCurrentLocationPathExpression, 1) = "." Then

      Dim intNumberOfTargetContextNode As Integer
      If intNumberOfContextNodes = 1 Then
        intNumberOfTargetContextNode = 1
      Else
        intNumberOfTargetContextNode = intLocationPathExpressionCounter
      End If

      Set CurrentContextNode = leleContextNodes(intNumberOfTargetContextNode)
      strInitialPPathForCurrentExpression = lstrInitialPPaths(intNumberOfTargetContextNode)
      If strInitialPPathForCurrentExpression = "" Then
        This.pPathReturnClass.SetErrorMessage = pPath.ConstantsAndStatic.MISSING_CONTEXT_NODE_INITIAL_PPATH
        GoTo ExitFunction:
      End If
      strCurrentLocationPathExpression = VBA.Strings.Mid(strCurrentLocationPathExpression, 2, VBA.Strings.Len(strCurrentLocationPathExpression))
    Else
      Set CurrentContextNode = This.ApplicationRootUIElement
      strInitialPPathForCurrentExpression = ""
    End If

    'Start each Location Path Expression at the Current Context Node
    This.pPathReturnClass.AddMatchingElement _
      This.pPathReturnClass.GetWorkingCopyOfCandidateElements(This.CurrentLocationPathExpressionCounter), _
      CurrentContextNode, _
      "", _
      strInitialPPathForCurrentExpression

    Dim mySteps() As pPath.Step
    Dim intStepCounter As Integer
    intStepCounter = 0
    Dim strRemainingPPath As String
    strRemainingPPath = strCurrentLocationPathExpression
    Do
      intStepCounter = intStepCounter + 1
      If intStepCounter = 1 Then
        ReDim mySteps(intStepCounter) As pPath.Step
      Else
        ReDim Preserve mySteps(intStepCounter) As pPath.Step
      End If
      Set mySteps(intStepCounter) = This.Steps.GetNextStep(strRemainingPPath)
      If This.pPathReturnClass.GetErrorMessage <> "" Then
        GoTo ExitFunction:
      End If
      strRemainingPPath = mySteps(intStepCounter).RemainingPPath
    Loop Until strRemainingPPath = ""

    Dim intNumberOfSteps As Integer
    intNumberOfSteps = UBound(mySteps)
    For intStepCounter = 1 To intNumberOfSteps

      Dim myNextStep As pPath.Step
      Set myNextStep = mySteps(intStepCounter)

      'Check for error messages with the next step
      If myNextStep.Axis = Axes.None Then
        This.pPathReturnClass.SetErrorMessage = pPath.ConstantsAndStatic.NULL_PPATH_ERROR_MESSAGE
      ElseIf myNextStep.NodeTest = "" And myNextStep.Axis <> Axes.SelfShorthand And myNextStep.Axis <> Axes.ParentShorthand Then
        'Don't raise this error for the self:: and parent:: axes!
        This.pPathReturnClass.SetErrorMessage = pPath.ConstantsAndStatic.NO_NODETEST_PPATH_ERROR_MESSAGE
      ElseIf This.pPathReturnClass.GetErrorMessage <> "" Then
        'Some other Next Step error message return, so skip to end here!
      Else

        This.Axes.ProcessNextAxis myNextStep
        This.pPathReturnClass.MoveCandidateElementsToWorkingCopy This.CurrentLocationPathExpressionCounter

        This.NodeTests.ProcessNextNodeTest myNextStep
        If myNextStep.NumberOfPredicateSets > 0 Then
          This.pPathReturnClass.MoveCandidateElementsToWorkingCopy This.CurrentLocationPathExpressionCounter
          This.Predicates.ProcessPredicates myNextStep
        End If

        'Don't MoveCandidateElementsToWorkingCopy at the last action if we have no more PPath, ie. another step to process
        If myNextStep.RemainingPPath <> "" Then
          This.pPathReturnClass.MoveCandidateElementsToWorkingCopy This.CurrentLocationPathExpressionCounter
        End If

      End If
      If This.pPathReturnClass.GetErrorMessage <> "" Then
        Exit For
      End If

    Next intStepCounter

  Next intLocationPathExpressionCounter

  This.pPathReturnClass.PromoteCandidateElementsToMatchingElements

  'Set a return value - apply top level functions
  If This.pPathReturnClass.TopLevelFunctionPrefix = "count(" And This.pPathReturnClass.TopLevelFunctionSuffix = ")" Then
    This.pPathReturnClass.ReturnedValue = This.pPathReturnClass.GetFinalNumberOfMatchingElements
  ElseIf This.pPathReturnClass.TopLevelFunctionPrefix = "not(" And This.pPathReturnClass.TopLevelFunctionSuffix = ")" Then
    This.pPathReturnClass.ReturnedValue = (This.pPathReturnClass.GetFinalNumberOfMatchingElements > 0)
  ElseIf This.pPathReturnClass.TopLevelFunctionPrefix <> "" Then
    Dim i As Integer
    Dim c As Integer
    c = This.pPathReturnClass.GetFinalNumberOfMatchingElements
    If c = 0 Then
      This.pPathReturnClass.ReturnedValue = 0
    Else
'      Excel.Application.ScreenUpdating = False
      pPath.Workbook.OpenWB

      Dim PPathWS As Worksheet
      Set PPathWS = gPPathWB.Worksheets("Sheet1")

      gPPathWB.Worksheets("Sheet1").Cells.ClearContents
      Dim strOutputArray As String
      Dim varValue As Variant
      Dim larrTopLevelFunctionArray() As Variant
      Erase larrTopLevelFunctionArray
      ReDim larrTopLevelFunctionArray(c - 1)
      For i = 1 To c
        varValue = pPath.Utils.GetValue(This.pPathReturnClass.GetMatchingElement(i))
        larrTopLevelFunctionArray(i - 1) = varValue
      Next i
'      PPathWS.Activate
      PPathWS.Range("A1:A" & c).Value = Application.Transpose(larrTopLevelFunctionArray)
'      varValue = PPathWS.Evaluate(this.PPathReturnClass.TopLevelFunctionPrefix & "A1:A" & c & this.PPathReturnClass.TopLevelFunctionSuffix)
      PPathWS.Range("B1").Formula = "=" & This.pPathReturnClass.TopLevelFunctionPrefix & "A1:A" & c & This.pPathReturnClass.TopLevelFunctionSuffix
      PPathWS.Range("B1").Calculate
      varValue = PPathWS.Range("B1").Value
      PPathWS.Cells.ClearContents
      VBA.Interaction.DoEvents
'      Excel.Application.ScreenUpdating = True
      This.pPathReturnClass.ReturnedValue = varValue
    End If
  Else
    ' Default is the 'effective boolean value'
    This.pPathReturnClass.ReturnedValue = (This.pPathReturnClass.GetFinalNumberOfMatchingElements > 0)
  End If

ExitFunction:
    Set Evaluate = This.pPathReturnClass

End Function

Private Function PreValidateFullPPathExpression(pPathToValidate As String) As Boolean

  If pPathToValidate = "" Then
    This.pPathReturnClass.SetErrorMessage = pPath.ConstantsAndStatic.NULL_PPATH_ERROR_MESSAGE
    PreValidateFullPPathExpression = False
    Exit Function
  End If
  If GetCountOfInstancesOfACharacterInAString("(", pPathToValidate) <> GetCountOfInstancesOfACharacterInAString(")", pPathToValidate) Then
    This.pPathReturnClass.SetErrorMessage = pPath.ConstantsAndStatic.MISMATCHING_PARENTHESES_ERROR_MESSAGE
    PreValidateFullPPathExpression = False
    Exit Function
  End If
  If GetCountOfInstancesOfACharacterInAString("[", pPathToValidate) <> GetCountOfInstancesOfACharacterInAString("]", pPathToValidate) Then
    This.pPathReturnClass.SetErrorMessage = pPath.ConstantsAndStatic.MISMATCHING_SQUARE_BARCKETS_ERROR_MESSAGE
    PreValidateFullPPathExpression = False
    Exit Function
  End If
  Dim intUnexpectedBracket As Integer
  intUnexpectedBracket = FindUnexpectedBracket(pPathToValidate)
  If intUnexpectedBracket > 0 Then
    Dim strUnexpectedBracketType As String
    strUnexpectedBracketType = VBA.Strings.Mid(pPathToValidate, intUnexpectedBracket, 1)
    This.pPathReturnClass.SetErrorMessage = "Unexpected '" & strUnexpectedBracketType & "' Bracket at position " & intUnexpectedBracket
    PreValidateFullPPathExpression = False
    Exit Function
  End If
  'TODO: Expand the check for valid properties!
  If VBA.Strings.InStr(1, pPathToValidate, "@") > 0 Then
    Dim boolValueExists As Boolean
    Dim intCharacter As Integer
    Dim strCharacter As String
    Dim intStartOfProperty As Integer
    Dim intEndOfProperty As Integer
    Dim strProperty As String
    Dim boolInsideAString As Boolean
    boolInsideAString = False
    For intCharacter = 1 To VBA.Strings.Len(pPathToValidate)
      strCharacter = VBA.Strings.Mid(pPathToValidate, intCharacter, 1)
      If strCharacter = "@" Then
        intStartOfProperty = intCharacter + 1
      End If
      If (Not boolInsideAString) And (strCharacter = VBA.Strings.Chr(34)) Then
        boolInsideAString = True
      ElseIf boolInsideAString And (strCharacter = VBA.Strings.Chr(34)) Then
        boolInsideAString = False
      ElseIf (strCharacter = ",") Or (strCharacter = "<") Or (strCharacter = "=") Or (strCharacter = ">") Or (Not boolInsideAString And (strCharacter = " ")) Then
        intEndOfProperty = intCharacter
        strProperty = VBA.Strings.Mid(pPathToValidate, intStartOfProperty, intEndOfProperty - intStartOfProperty)
        boolValueExists = This.AutomationDictionaries.ValueExists(strProperty, This.AutomationDictionaries.NavigablePropertyIDs)
'        intStartOfProperty = 0
'        intEndOfProperty = 0
        'If Not boolValueExists Then
        If boolValueExists Then
          Exit For
        Else
          GoTo Fail
        End If
      End If
    Next intCharacter
    If intStartOfProperty > 0 And intEndOfProperty = 0 Then
      strProperty = VBA.Strings.Mid(pPathToValidate, intStartOfProperty, VBA.Strings.Len(pPathToValidate))
      boolValueExists = This.AutomationDictionaries.ValueExists(strProperty, This.AutomationDictionaries.NavigablePropertyIDs)
      If Not boolValueExists Then
        GoTo Fail
      End If
    End If
  End If

Pass:
  PreValidateFullPPathExpression = True
  Exit Function

Fail:
  This.pPathReturnClass.SetErrorMessage = pPath.ConstantsAndStatic.INVALID_PROPERTY_ERROR_MESSAGE
  PreValidateFullPPathExpression = False

End Function

Public Function GetCountOfInstancesOfACharacterInAString(character As String, SourceString As String)
  'https://stackoverflow.com/questions/9260982/how-to-find-number-of-occurences-of-slash-from-a-strings
  GetCountOfInstancesOfACharacterInAString = VBA.Strings.Len(SourceString) - VBA.Strings.Len(Replace(SourceString, character, ""))
End Function

Function FindUnexpectedBracket(pPath As String) As Long
  Dim i As Long
  Dim bracketStack As String
  Dim char As String

  ' Initialize stack to track opening brackets
  bracketStack = ""

  ' Iterate through each character
  For i = 1 To VBA.Strings.Len(pPath)
    char = VBA.Strings.Mid(pPath, i, 1)

    Select Case char
      Case "("
        ' Push opening round bracket to stack
        bracketStack = bracketStack & "("
      Case "["
        ' Push opening square bracket to stack
        bracketStack = bracketStack & "["
      Case ")"
        ' Check if closing round bracket matches last opening bracket
        If VBA.Strings.Len(bracketStack) = 0 Or VBA.Strings.Right(bracketStack, 1) <> "(" Then
          FindUnexpectedBracket = i
          Exit Function
        End If
        ' Pop the last opening bracket
        bracketStack = VBA.Strings.Left(bracketStack, VBA.Strings.Len(bracketStack) - 1)
      Case "]"
        ' Check if closing square bracket matches last opening bracket
        If VBA.Strings.Len(bracketStack) = 0 Or VBA.Strings.Right(bracketStack, 1) <> "[" Then
          FindUnexpectedBracket = i
          Exit Function
        End If
        ' Pop the last opening bracket
        bracketStack = VBA.Strings.Left(bracketStack, VBA.Strings.Len(bracketStack) - 1)
    End Select
  Next i

  ' Return 0 if no unexpected brackets found
  FindUnexpectedBracket = 0
End Function

Private Function PreValidateSplitPPathExpression(pPathArray() As String, ContextNodes() As UIAutomationClient.IUIAutomationElement) As Boolean
  '/& // are absolute and start at the root
  'Anything else is relative to a context node which must be set

  Dim intNumberOfContextNodes As Integer
  intNumberOfContextNodes = Phosphorus.Utils.GetSizeOfArray(ContextNodes)

  Dim intNumberOfRelativeExpressions As Integer
  intNumberOfRelativeExpressions = 0

  Dim boolContextModeUsed As Boolean
  boolContextModeUsed = False
  'Null PPath are caught elsewhere
  Dim iCounter As Integer
  Dim iCount As Integer
  iCount = UBound(pPathArray)
  For iCounter = 1 To iCount
    If pPathArray(iCounter) = "" Then
      This.pPathReturnClass.SetErrorMessage = "PPath #" & iCounter & ": " & pPath.ConstantsAndStatic.NULL_PPATH_ERROR_MESSAGE
      PreValidateSplitPPathExpression = False
      Exit Function
    End If
    If (VBA.Strings.Left(pPathArray(iCounter), 2) <> "//") And (VBA.Strings.Left(pPathArray(iCounter), 1) <> "/") Then
      ' We have a relative PPath
      If intNumberOfContextNodes = 0 Then
        This.pPathReturnClass.SetErrorMessage = pPath.ConstantsAndStatic.MISSING_RELATIVE_PPATH_CONTEXT_NODE
        PreValidateSplitPPathExpression = False
        Exit Function
      Else
        'A context mode has been referenced!
        boolContextModeUsed = True
        intNumberOfRelativeExpressions = intNumberOfRelativeExpressions + 1
      End If
    End If
  Next iCounter

  If intNumberOfContextNodes >= 1 And Not boolContextModeUsed Then
    This.pPathReturnClass.SetErrorMessage = pPath.ConstantsAndStatic.UNUSED_RELATIVE_PPATH_CONTEXT_NODE
    PreValidateSplitPPathExpression = False
    Exit Function
  End If

  If intNumberOfContextNodes > 1 And (intNumberOfContextNodes <> intNumberOfRelativeExpressions) Then
    This.pPathReturnClass.SetErrorMessage = pPath.ConstantsAndStatic.NUMBER_OF_RELATIVE_PPATHS_TO_CONTEXT_NODES_MISMATCH
    PreValidateSplitPPathExpression = False
    Exit Function
  End If

  PreValidateSplitPPathExpression = True

End Function

'============================================

'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-automation-element-propids
'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-entry-constants

'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controltypesoverview
'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controltype-ids




