VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder PPath
Option Explicit

Private lintNumberOfXPathExpressions As Integer

Private this As PPathCommon

'Interim Processing Variables
Private lstrRemainingXpath As String
Private lstrNextAxis As String
Private lstrNextNodeTest As String
Private leleContextNodes() As UIAutomationClient.IUIAutomationElement
Private lstrInitialPPaths() As String
  
Public Sub Initialise()
  Set this = New PPathCommon
  Erase leleContextNodes
  Erase lstrInitialPPaths
  this.DebugMode = False
  If this.UIAutomation Is Nothing Then
    Set this.UIAutomation = New CUIAutomation
    'Set initial root element as the whole Desktop
    Set this.ApplicationRootUIElement = this.UIAutomation.GetRootElement
  End If
  Set this.TreeWalker = this.UIAutomation.ControlViewWalker
  If this.AutomationDictionaries Is Nothing Then
    Set this.AutomationDictionaries = New UIAutomationDictionaries
    this.AutomationDictionaries.Initialise
  End If
  Set this.Steps = New PPathSteps
  Set this.Axes = New PPathAxes
  Set this.NodeTests = New PPathNodeTests
  Set this.Predicates = New PPathPredicates
  this.Steps.Initialise this
  this.Axes.Initialise this
  this.NodeTests.Initialise this
  this.Predicates.Initialise this
  PPathExceTopLevelFunctions.InitialiseTopLevelFunctions
End Sub

Public Sub SetApplicationRootElement(eleNewRootUIElement As UIAutomationClient.IUIAutomationElement)
  Set this.ApplicationRootUIElement = eleNewRootUIElement
End Sub

Public Sub AddContextNode(ContextNode As UIAutomationClient.IUIAutomationElement, InitialPPath As String, Optional FirstNode As Boolean)
  If FirstNode Then
    Erase leleContextNodes
    Erase lstrInitialPPaths
  End If
  Dim i As Integer
  i = Utils.GetSizeOfArray(leleContextNodes()) + 1
  If i = 1 Then
    ReDim leleContextNodes(i)
    ReDim lstrInitialPPaths(i)
  Else
    ReDim Preserve leleContextNodes(i)
    ReDim Preserve lstrInitialPPaths(i)
  End If
  Set leleContextNodes(i) = ContextNode
  lstrInitialPPaths(i) = InitialPPath
End Sub

Public Property Let SetDebugMode(boolDebugMode As Boolean)
  this.DebugMode = boolDebugMode
End Property

Public Property Get GetDebugMode() As Boolean
  GetDebugMode = this.DebugMode
End Property

Public Sub InitialiseEvaluation(intNumberOfXPathExpressions As Integer)
  Phosphorus.PPathConstants.InitialiseAxesTypes
  Set this.PPathReturnClass = Nothing
  Set this.PPathReturnClass = Phosphorus.GetNewPhosphorusPPathReturnClass(intNumberOfXPathExpressions)
End Sub
 
'See: https://www.w3schools.com/xml/xpath_syntax.asp

'XPath Set operators
'https://infocenter.sybase.com/help/index.jsp?topic=/com.sybase.dc30020_1251/html/xmlb/xmlb32.htm

Public Function Evaluate( _
  ByVal FullLocationPathExpression As String, _
  Optional ContextNode As UIAutomationClient.IUIAutomationElement, _
  Optional InitialPPath As String, _
  Optional UnitTestingMode As Boolean _
  ) As Phosphorus.PPathReturnClass

  If Not ContextNode Is Nothing Then 'An InitialPPath should also be provided
    'Add first context node
    AddContextNode ContextNode, InitialPPath, True
  End If

  this.UnitTestingMode = UnitTestingMode
  
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
    For Each TopLevelFunction In PPathExceTopLevelFunctions.TopLevelFunctions
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
  strAllXPathExpressions = this.Steps.GetAllXPathExpressions(FullLocationPathExpression)
  lintNumberOfXPathExpressions = UBound(strAllXPathExpressions)
  
  InitialiseEvaluation lintNumberOfXPathExpressions
  this.PPathReturnClass.TopLevelFunctionPrefix = strTopLevelFunctionPrefix
   this.PPathReturnClass.TopLevelFunctionSuffix = strTopLevelFunctionSuffix

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
  intNumberOfContextNodes = Utils.GetSizeOfArray(leleContextNodes)

  For intLocationPathExpressionCounter = 1 To lintNumberOfXPathExpressions
  
    this.CurrentLocationPathExpressionCounter = intLocationPathExpressionCounter
    strCurrentLocationPathExpression = strAllXPathExpressions(this.CurrentLocationPathExpressionCounter)
 
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
        this.PPathReturnClass.SetErrorMessage = Phosphorus.PPathConstants.MISSING_CONTEXT_NODE_INITIAL_PPATH
        GoTo ExitFunction:
      End If
      strCurrentLocationPathExpression = VBA.Strings.Mid(strCurrentLocationPathExpression, 2, VBA.Strings.Len(strCurrentLocationPathExpression))
    Else
      Set CurrentContextNode = this.ApplicationRootUIElement
      strInitialPPathForCurrentExpression = ""
    End If
    
    'Start each Location Path Expression at the Current Context Node
    this.PPathReturnClass.AddMatchingElement _
      this.PPathReturnClass.GetWorkingCopyOfCandidateElements(this.CurrentLocationPathExpressionCounter), _
      CurrentContextNode, _
      "", _
      strInitialPPathForCurrentExpression

    Dim mySteps() As Phosphorus.PPathStep
    Dim intStepCounter As Integer
    intStepCounter = 0
    Dim strRemainingPPath As String
    strRemainingPPath = strCurrentLocationPathExpression
    Do
      intStepCounter = intStepCounter + 1
      If intStepCounter = 1 Then
        ReDim mySteps(intStepCounter) As Phosphorus.PPathStep
      Else
        ReDim Preserve mySteps(intStepCounter) As Phosphorus.PPathStep
      End If
      Set mySteps(intStepCounter) = this.Steps.GetNextStep(strRemainingPPath)
      If this.PPathReturnClass.GetErrorMessage <> "" Then
        GoTo ExitFunction:
      End If
      strRemainingPPath = mySteps(intStepCounter).RemainingPPath
      
    Loop Until strRemainingPPath = ""
    
    Dim intNumberOfSteps As Integer
    intNumberOfSteps = UBound(mySteps)
    For intStepCounter = 1 To intNumberOfSteps
    
      Dim myNextStep As Phosphorus.PPathStep
      Set myNextStep = mySteps(intStepCounter)
    
      'Check for error messages with the next step
      If myNextStep.Axis = Axes.None Then
        this.PPathReturnClass.SetErrorMessage = Phosphorus.PPathConstants.NULL_PPATH_ERROR_MESSAGE
      ElseIf myNextStep.NodeTest = "" And myNextStep.Axis <> Axes.SelfShorthand And myNextStep.Axis <> Axes.ParentShorthand Then
        'Don't raise this error for the self:: and parent:: axes!
        this.PPathReturnClass.SetErrorMessage = Phosphorus.PPathConstants.NO_NODETEST_PPATH_ERROR_MESSAGE
      ElseIf this.PPathReturnClass.GetErrorMessage <> "" Then
        'Some other Next Step error message return, so skip to end here!
      Else
        this.Axes.ProcessNextAxis myNextStep
        this.PPathReturnClass.MoveCandidateElementsToWorkingCopy this.CurrentLocationPathExpressionCounter
        
        this.NodeTests.ProcessNextNodeTest myNextStep
   
        If myNextStep.NumberOfPredicateSets > 0 Then
          this.PPathReturnClass.MoveCandidateElementsToWorkingCopy this.CurrentLocationPathExpressionCounter
          this.Predicates.ProcessPredicates myNextStep
        End If
        
        'Don't MoveCandidateElementsToWorkingCopy at the last action if we have no more PPath, ie. another step to process
        If myNextStep.RemainingPPath <> "" Then
          this.PPathReturnClass.MoveCandidateElementsToWorkingCopy this.CurrentLocationPathExpressionCounter
        End If
      
      End If
      If this.PPathReturnClass.GetErrorMessage <> "" Then
        Exit For
      End If
     
    Next intStepCounter

  Next intLocationPathExpressionCounter
  
  this.PPathReturnClass.PromoteCandidateElementsToMatchingElements
    
  'Set a return value - apply top level functions
  If this.PPathReturnClass.TopLevelFunctionPrefix = "count(" And this.PPathReturnClass.TopLevelFunctionSuffix = ")" Then
    this.PPathReturnClass.ReturnedValue = this.PPathReturnClass.GetFinalNumberOfMatchingElements
  ElseIf this.PPathReturnClass.TopLevelFunctionPrefix = "not(" And this.PPathReturnClass.TopLevelFunctionSuffix = ")" Then
    this.PPathReturnClass.ReturnedValue = (this.PPathReturnClass.GetFinalNumberOfMatchingElements > 0)
  ElseIf this.PPathReturnClass.TopLevelFunctionPrefix <> "" Then
    Dim i As Integer
    Dim c As Integer
    c = this.PPathReturnClass.GetFinalNumberOfMatchingElements
    If c = 0 Then
      this.PPathReturnClass.ReturnedValue = 0
    Else
'      Excel.Application.ScreenUpdating = False
      PPathWorkbook.OpenWB
      
      Dim PPathWS As Worksheet
      Set PPathWS = gPPathWB.Worksheets("Sheet1")
      
      gPPathWB.Worksheets("Sheet1").Cells.ClearContents
      Dim strOutputArray As String
      Dim varValue As Variant
      Dim larrTopLevelFunctionArray() As Variant
      Erase larrTopLevelFunctionArray
      ReDim larrTopLevelFunctionArray(c - 1)
      For i = 1 To c
        varValue = PPathUtils.GetValue(this.PPathReturnClass.GetMatchingElement(i))
        larrTopLevelFunctionArray(i - 1) = varValue
      Next i
'      PPathWS.Activate
      PPathWS.Range("A1:A" & c).Value = Application.Transpose(larrTopLevelFunctionArray)
'      varValue = PPathWS.Evaluate(this.PPathReturnClass.TopLevelFunctionPrefix & "A1:A" & c & this.PPathReturnClass.TopLevelFunctionSuffix)
      PPathWS.Range("B1").Formula = "=" & this.PPathReturnClass.TopLevelFunctionPrefix & "A1:A" & c & this.PPathReturnClass.TopLevelFunctionSuffix
      PPathWS.Range("B1").Calculate
      varValue = PPathWS.Range("B1").Value
      PPathWS.Cells.ClearContents
      VBA.Interaction.DoEvents
'      Excel.Application.ScreenUpdating = True
      this.PPathReturnClass.ReturnedValue = varValue
    End If
  Else
    ' Default is the 'effective boolean value'
    this.PPathReturnClass.ReturnedValue = (this.PPathReturnClass.GetFinalNumberOfMatchingElements > 0)
  End If

ExitFunction:
    Set Evaluate = this.PPathReturnClass
  
End Function

Private Function PreValidateFullPPathExpression(PPath As String) As Boolean

  If PPath = "" Then
    this.PPathReturnClass.SetErrorMessage = Phosphorus.PPathConstants.NULL_PPATH_ERROR_MESSAGE
    PreValidateFullPPathExpression = False
    Exit Function
  End If
  If GetCountOfInstancesOfACharacterInAString("(", PPath) <> GetCountOfInstancesOfACharacterInAString(")", PPath) Then
    this.PPathReturnClass.SetErrorMessage = Phosphorus.PPathConstants.MISMATCHING_PARENTHESES_ERROR_MESSAGE
    PreValidateFullPPathExpression = False
    Exit Function
  End If
  If GetCountOfInstancesOfACharacterInAString("[", PPath) <> GetCountOfInstancesOfACharacterInAString("]", PPath) Then
    this.PPathReturnClass.SetErrorMessage = Phosphorus.PPathConstants.MISMATCHING_SQUARE_BARCKETS_ERROR_MESSAGE
    PreValidateFullPPathExpression = False
    Exit Function
  End If
  Dim intUnexpectedBracket As Integer
  intUnexpectedBracket = FindUnexpectedBracket(PPath)
  If intUnexpectedBracket > 0 Then
    Dim strUnexpectedBracketType As String
    strUnexpectedBracketType = VBA.Strings.Mid(PPath, intUnexpectedBracket, 1)
    this.PPathReturnClass.SetErrorMessage = "Unexpected '" & strUnexpectedBracketType & "' Bracket at position " & intUnexpectedBracket
    PreValidateFullPPathExpression = False
    Exit Function
  End If
  'TODO: Expand the check for valid properties!
  If VBA.Strings.InStr(1, PPath, "@") > 0 Then
    Dim boolValueExists As Boolean
    Dim intCharacter As Integer
    Dim strCharacter As String
    Dim intStartOfProperty As Integer
    Dim intEndOfProperty As Integer
    Dim strProperty As String
    Dim boolInsideAString As Boolean
    boolInsideAString = False
    For intCharacter = 1 To VBA.Strings.Len(PPath)
      strCharacter = VBA.Strings.Mid(PPath, intCharacter, 1)
      If strCharacter = "@" Then
        intStartOfProperty = intCharacter + 1
      End If
      If (Not boolInsideAString) And (strCharacter = Chr(34)) Then
        boolInsideAString = True
      ElseIf boolInsideAString And (strCharacter = Chr(34)) Then
        boolInsideAString = False
      ElseIf (strCharacter = "<") Or (strCharacter = "=") Or (strCharacter = ">") Or (Not boolInsideAString And (strCharacter = " ")) Then
        intEndOfProperty = intCharacter
        strProperty = VBA.Strings.Mid(PPath, intStartOfProperty, intEndOfProperty - intStartOfProperty)
        boolValueExists = this.AutomationDictionaries.ValueExists(strProperty, this.AutomationDictionaries.NavigablePropertyIDs)
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
      strProperty = VBA.Strings.Mid(PPath, intStartOfProperty, VBA.Strings.Len(PPath))
      boolValueExists = this.AutomationDictionaries.ValueExists(strProperty, this.AutomationDictionaries.NavigablePropertyIDs)
      If Not boolValueExists Then
        GoTo Fail
      End If
    End If
  End If
  
Pass:
  PreValidateFullPPathExpression = True
  Exit Function

Fail:
  this.PPathReturnClass.SetErrorMessage = Phosphorus.PPathConstants.INVALID_PROPERTY_ERROR_MESSAGE
  PreValidateFullPPathExpression = False

End Function

Public Function GetCountOfInstancesOfACharacterInAString(Character As String, SourceString As String)
  'https://stackoverflow.com/questions/9260982/how-to-find-number-of-occurences-of-slash-from-a-strings
  GetCountOfInstancesOfACharacterInAString = Len(SourceString) - Len(Replace(SourceString, Character, ""))
End Function

Function FindUnexpectedBracket(PPath As String) As Long
  Dim i As Long
  Dim bracketStack As String
  Dim char As String
    
  ' Initialize stack to track opening brackets
  bracketStack = ""
    
  ' Iterate through each character
  For i = 1 To Len(PPath)
    char = Mid(PPath, i, 1)
       
    Select Case char
      Case "("
        ' Push opening round bracket to stack
        bracketStack = bracketStack & "("
      Case "["
        ' Push opening square bracket to stack
        bracketStack = bracketStack & "["
      Case ")"
        ' Check if closing round bracket matches last opening bracket
        If Len(bracketStack) = 0 Or Right(bracketStack, 1) <> "(" Then
          FindUnexpectedBracket = i
          Exit Function
        End If
        ' Pop the last opening bracket
        bracketStack = Left(bracketStack, Len(bracketStack) - 1)
      Case "]"
        ' Check if closing square bracket matches last opening bracket
        If Len(bracketStack) = 0 Or Right(bracketStack, 1) <> "[" Then
          FindUnexpectedBracket = i
          Exit Function
        End If
        ' Pop the last opening bracket
        bracketStack = Left(bracketStack, Len(bracketStack) - 1)
    End Select
  Next i
    
  ' Return 0 if no unexpected brackets found
  FindUnexpectedBracket = 0
End Function

Private Function PreValidateSplitPPathExpression(PPath() As String, ContextNodes() As UIAutomationClient.IUIAutomationElement) As Boolean
  '/& // are absolute and start at the root
  'Anything else is relative to a context node which must be set
  
  Dim intNumberOfContextNodes As Integer
  intNumberOfContextNodes = Utils.GetSizeOfArray(ContextNodes)
  
  Dim intNumberOfRelativeExpressions As Integer
  intNumberOfRelativeExpressions = 0
  
  Dim boolContextModeUsed As Boolean
  boolContextModeUsed = False
  'Null PPath are caught elsewhere
  Dim iCounter As Integer
  Dim iCount As Integer
  iCount = UBound(PPath)
  For iCounter = 1 To iCount
    If PPath(iCounter) = "" Then
      this.PPathReturnClass.SetErrorMessage = "PPath #" & iCounter & ": " & Phosphorus.PPathConstants.NULL_PPATH_ERROR_MESSAGE
      PreValidateSplitPPathExpression = False
      Exit Function
    End If
    If (VBA.Strings.Left(PPath(iCounter), 2) <> "//") And (VBA.Strings.Left(PPath(iCounter), 1) <> "/") Then
      ' We have a relative PPath
      If intNumberOfContextNodes = 0 Then
        this.PPathReturnClass.SetErrorMessage = Phosphorus.PPathConstants.MISSING_RELATIVE_PPATH_CONTEXT_NODE
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
    this.PPathReturnClass.SetErrorMessage = Phosphorus.PPathConstants.UNUSED_RELATIVE_PPATH_CONTEXT_NODE
    PreValidateSplitPPathExpression = False
    Exit Function
  End If
  
  If intNumberOfContextNodes > 1 And (intNumberOfContextNodes <> intNumberOfRelativeExpressions) Then
    this.PPathReturnClass.SetErrorMessage = Phosphorus.PPathConstants.NUMBER_OF_RELATIVE_PPATHS_TO_CONTEXT_NODES_MISMATCH
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
