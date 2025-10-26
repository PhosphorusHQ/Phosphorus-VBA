VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPathPredicates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder PPath
Option Explicit

Private This As PPathCommon
Private lstrNestedPPath As String
Private lstrCurrentPredicateTest As String

Public Sub Initialise(ByRef sharedthis As PPathCommon)
  Set This = sharedthis
End Sub

Public Sub ProcessPredicates(myNextStep As Phosphorus.PPathStep)

  Dim intNumberOfPredicateSets As Integer
  Dim intPredicateSetCounter As Integer
  Dim strCurrentSetAttributesPPath As String
  intNumberOfPredicateSets = myNextStep.NumberOfPredicateSets
  For intPredicateSetCounter = 1 To intNumberOfPredicateSets

    strCurrentSetAttributesPPath = myNextStep.PredicateSet(intPredicateSetCounter).SourcePPath

    Dim intNumberOfPredicateGroups As Integer
    Dim intPredicateGroupCounter As Integer
    intNumberOfPredicateGroups = myNextStep.PredicateSet(intPredicateSetCounter).NumberOfPredicateGroups
    For intPredicateGroupCounter = 1 To intNumberOfPredicateGroups
      
      'Do we have any predicates to process for this group?
      Dim strCurrentGroupAttributesPPath As String
      strCurrentGroupAttributesPPath = myNextStep.PredicateSet(intPredicateSetCounter).PredicateGroup(intPredicateGroupCounter).FinalPPath
      If strCurrentGroupAttributesPPath <> "" Then
  
        Dim boolIsFirstSetPositionalPredicate As Boolean
        boolIsFirstSetPositionalPredicate = ( _
          intPredicateSetCounter = 1 And _
          myNextStep.PredicateSet(intPredicateSetCounter).PredicateGroup(intPredicateGroupCounter).IsPositionalPredicate)

        Dim intSetElementCounter As Integer
        Dim intNumberOfSetElements As Integer
        Dim eleCurrentUIElement As UIAutomationClient.IUIAutomationElement
        Dim strCurrentNavigationalPPath As String
        Dim strCurrentAttributeName As String
        Dim strCurrentNodeControlType As String
    
        'Prepare context node counter variables
        If boolIsFirstSetPositionalPredicate Then
          Dim eleCurrentParentUIElement As UIAutomationClient.IUIAutomationElement
          Dim strCurrentParentElementRuntimeID As String
          Dim intContextElementPositionCounter As Integer
          Dim intContextElementLastCounter As Integer
        End If
        
        intNumberOfSetElements = This.PPathReturnClass.GetNumberOfWorkingCopyOfCandidateElements(This.CurrentLocationPathExpressionCounter)
        For intSetElementCounter = 1 To intNumberOfSetElements
    
          Set eleCurrentUIElement = This.PPathReturnClass.GetWorkingCopyElement(This.CurrentLocationPathExpressionCounter, intSetElementCounter)
          strCurrentNavigationalPPath = This.PPathReturnClass.GetWorkingCopyNavigationalPPath(This.CurrentLocationPathExpressionCounter, intSetElementCounter)
          strCurrentAttributeName = This.PPathReturnClass.GetWorkingCopyAttributeName(This.CurrentLocationPathExpressionCounter, intSetElementCounter)
          strCurrentNodeControlType = This.AutomationDictionaries.ControlTypeIDs(eleCurrentUIElement.CurrentControlType)
  
          If boolIsFirstSetPositionalPredicate Then
            Set eleCurrentParentUIElement = This.TreeWalker.GetParentElement(eleCurrentUIElement)
            strCurrentParentElementRuntimeID = PPathRuntimeIDs.GetElementRuntimeID(eleCurrentParentUIElement)
            'Count the number of preceeding elements that have the same parent to get the current element's position
            intContextElementPositionCounter = 0
            Dim intPreviousElementCounter As Integer
            For intPreviousElementCounter = intSetElementCounter To 1 Step -1
              Dim elePreviousUIElement As UIAutomationClient.IUIAutomationElement
              Dim elePreviousParentUIElement As UIAutomationClient.IUIAutomationElement
              Dim strPreviousParentElementRuntimeID As String
              Set elePreviousUIElement = This.PPathReturnClass.GetWorkingCopyElement(This.CurrentLocationPathExpressionCounter, intPreviousElementCounter)
              Set elePreviousParentUIElement = This.TreeWalker.GetParentElement(elePreviousUIElement)
              strPreviousParentElementRuntimeID = PPathRuntimeIDs.GetElementRuntimeID(elePreviousParentUIElement)
              If strPreviousParentElementRuntimeID = strCurrentParentElementRuntimeID Then
                intContextElementPositionCounter = intContextElementPositionCounter + 1
              Else
                'Exit once we get past the current element's parent element
                Exit For
              End If
            Next intPreviousElementCounter
            'Now find the last element number, if necessary
            If VBA.Strings.InStr(1, strCurrentGroupAttributesPPath, "last()") > 0 Then
              intContextElementLastCounter = intContextElementPositionCounter
              Dim intNextElementCounter As Integer
              For intNextElementCounter = (intSetElementCounter + 1) To intNumberOfSetElements Step 1
                Dim eleNextUIElement As UIAutomationClient.IUIAutomationElement
                Dim eleNextParentUIElement As UIAutomationClient.IUIAutomationElement
                Dim strNextParentElementRuntimeID As String
                Set eleNextUIElement = This.PPathReturnClass.GetWorkingCopyElement(This.CurrentLocationPathExpressionCounter, intNextElementCounter)
                Set eleNextParentUIElement = This.TreeWalker.GetParentElement(eleNextUIElement)
                strNextParentElementRuntimeID = PPathRuntimeIDs.GetElementRuntimeID(eleNextParentUIElement)
                If strCurrentParentElementRuntimeID = strNextParentElementRuntimeID Then
                  intContextElementLastCounter = intContextElementLastCounter + 1
                Else
                  'Exit once we get past the current element's parent element
                  Exit For
                End If
              Next intNextElementCounter
            End If
          End If

          'Replace all the positional functions in the current predicate group
          lstrCurrentPredicateTest = strCurrentGroupAttributesPPath
          lstrCurrentPredicateTest = VBA.Strings.Replace(lstrCurrentPredicateTest, "first()", 1)
          If boolIsFirstSetPositionalPredicate Then
            lstrCurrentPredicateTest = VBA.Strings.Replace(lstrCurrentPredicateTest, "position()", intContextElementPositionCounter)
            lstrCurrentPredicateTest = VBA.Strings.Replace(lstrCurrentPredicateTest, "last()", intContextElementLastCounter)
          Else
            lstrCurrentPredicateTest = VBA.Strings.Replace(lstrCurrentPredicateTest, "position()", intSetElementCounter)
            lstrCurrentPredicateTest = VBA.Strings.Replace(lstrCurrentPredicateTest, "last()", intNumberOfSetElements)
          End If

          ProcessPPathFunctions eleCurrentUIElement, strCurrentNavigationalPPath & "/"
          ProcessNestedPPaths eleCurrentUIElement, strCurrentNavigationalPPath & "/"
          ProcessAttributeChecks eleCurrentUIElement, strCurrentNavigationalPPath & "/"
          ProcessTextAndValueChecks eleCurrentUIElement
          lstrCurrentPredicateTest = PPathExceUserDefinedFunctions.RenameExcelFunctions(lstrCurrentPredicateTest)
          
          Dim boolCurrentPredicateTestPasses As Boolean
          On Error Resume Next
          boolCurrentPredicateTestPasses = True
          boolCurrentPredicateTestPasses = Application.Evaluate(lstrCurrentPredicateTest)
          If Err.Number = 13 Then
            This.PPathReturnClass.SetErrorMessage = Phosphorus.PPathConstants.INVALID_PREDICATE & " [" & strCurrentGroupAttributesPPath & "] => " & lstrCurrentPredicateTest
            Exit Sub
          End If
          On Error GoTo 0
          If boolCurrentPredicateTestPasses Then
            PPathUtils.OutputElementDetails eleCurrentUIElement, strCurrentNavigationalPPath, This.DebugMode, This.AutomationDictionaries
            This.PPathReturnClass.AddMatchingElement This.PPathReturnClass.GetCandidateElements(This.CurrentLocationPathExpressionCounter), eleCurrentUIElement, strCurrentAttributeName, strCurrentNavigationalPPath
          End If

        Next intSetElementCounter

        'Reset Working Copy after each interim predicate group
        If (intPredicateGroupCounter < intNumberOfPredicateGroups) And (strCurrentGroupAttributesPPath <> "") Then
          This.PPathReturnClass.MoveCandidateElementsToWorkingCopy This.CurrentLocationPathExpressionCounter
        End If

      End If

    Next intPredicateGroupCounter
    
    'Reset Working Copy after each interim predicate set
    If (intPredicateSetCounter < intNumberOfPredicateSets) And (strCurrentSetAttributesPPath <> "") Then
      This.PPathReturnClass.MoveCandidateElementsToWorkingCopy This.CurrentLocationPathExpressionCounter
    End If
    
  Next intPredicateSetCounter

End Sub

Private Sub ProcessPPathFunctions( _
  eleCurrentContextUIElement As UIAutomationClient.IUIAutomationElement, _
  strInitialPPath As String)
  
  While SourcePPathContainedATopLevelFunction(eleCurrentContextUIElement, strInitialPPath, "count")
  Wend
  While SourcePPathContainedATopLevelFunction(eleCurrentContextUIElement, strInitialPPath, "not")
  Wend
    
End Sub

Private Function SourcePPathContainedATopLevelFunction( _
  eleCurrentContextUIElement As UIAutomationClient.IUIAutomationElement, _
  strInitialPPath As String, _
  strFunctionName As String) As Boolean
  
  SourcePPathContainedATopLevelFunction = False
  Dim intStartOfFunction As Integer
  intStartOfFunction = VBA.Strings.InStr(1, lstrCurrentPredicateTest, strFunctionName & "(")
  If intStartOfFunction > 0 Then
    SourcePPathContainedATopLevelFunction = True
    Dim intEndOfFunction As Integer
    Dim intParenthesisCounter As Integer
    intParenthesisCounter = 1
    For intEndOfFunction = (intStartOfFunction + VBA.Strings.Len(strFunctionName) + 1) To VBA.Strings.Len(lstrCurrentPredicateTest)
      Dim strCurrentCharacter As String
      strCurrentCharacter = VBA.Strings.Mid(lstrCurrentPredicateTest, intEndOfFunction, 1)
      If strCurrentCharacter = "(" Then
        intParenthesisCounter = intParenthesisCounter + 1
      Else
        If strCurrentCharacter = ")" Then
          intParenthesisCounter = intParenthesisCounter - 1
        End If
      End If
      If intParenthesisCounter = 0 Then
        lstrNestedPPath = VBA.Strings.Mid(lstrCurrentPredicateTest, intStartOfFunction, intEndOfFunction - intStartOfFunction + 1)
        Exit For
      End If
    Next intEndOfFunction
  End If
  If SourcePPathContainedATopLevelFunction Then
    Dim NestedPPath As Phosphorus.PPath
    Dim EvaluatedNestedPPath As Phosphorus.PPathReturnClass
    Dim varReturnValue As Variant
    Set NestedPPath = Phosphorus.Factory.GetNewPhosphorusPPath
    NestedPPath.Initialise
    Set EvaluatedNestedPPath = NestedPPath.Evaluate(lstrNestedPPath, eleCurrentContextUIElement, strInitialPPath)
    'Raise any error?
    This.PPathReturnClass.SetErrorMessage = EvaluatedNestedPPath.GetErrorMessage
    varReturnValue = EvaluatedNestedPPath.ReturnedValue
    lstrCurrentPredicateTest = VBA.Strings.Replace(lstrCurrentPredicateTest, lstrNestedPPath, varReturnValue)
    Set NestedPPath = Nothing
    Set EvaluatedNestedPPath = Nothing
  End If
End Function

Private Sub ProcessNestedPPaths( _
  eleCurrentContextUIElement As UIAutomationClient.IUIAutomationElement, _
  strInitialPPath As String)
  
  While SourcePPathContainedANestedPPath(eleCurrentContextUIElement, strInitialPPath)
  Wend
    
End Sub

Private Function SourcePPathContainedANestedPPath( _
  eleCurrentContextUIElement As UIAutomationClient.IUIAutomationElement, _
  strInitialPPath As String) As Boolean
  
  SourcePPathContainedANestedPPath = False
  Dim intStartOfNestedPPath As Integer
  intStartOfNestedPPath = VBA.Strings.InStr(1, lstrCurrentPredicateTest, "./")
  If intStartOfNestedPPath > 0 Then
    SourcePPathContainedANestedPPath = True
    Dim intEndOfNestedPPath As Integer
    Dim boolInsideAPredicate As Boolean
    boolInsideAPredicate = False
    For intEndOfNestedPPath = (intStartOfNestedPPath + VBA.Strings.Len("./") + 1) To VBA.Strings.Len(lstrCurrentPredicateTest)
      Dim strCurrentCharacter As String
      strCurrentCharacter = VBA.Strings.Mid(lstrCurrentPredicateTest, intEndOfNestedPPath, 1)
      If strCurrentCharacter = "[" Then
        boolInsideAPredicate = True
      Else
        If strCurrentCharacter = "]" Then
          boolInsideAPredicate = False
        End If
      End If
      If Not boolInsideAPredicate And VBA.Strings.InStr(1, " ,)", strCurrentCharacter) > 0 Then
        lstrNestedPPath = VBA.Strings.Mid(lstrCurrentPredicateTest, intStartOfNestedPPath, intEndOfNestedPPath - intStartOfNestedPPath)
        Exit For
      End If
    Next intEndOfNestedPPath
  End If
  If intEndOfNestedPPath = VBA.Strings.Len(lstrCurrentPredicateTest) + 1 Then
    'The whole predicate is a single predicate test
    lstrNestedPPath = lstrCurrentPredicateTest
  End If
  If SourcePPathContainedANestedPPath Then
    Dim NestedPPath As Phosphorus.PPath
    Dim EvaluatedNestedPPath As Phosphorus.PPathReturnClass
    Dim varReturnValue As Variant
    Set NestedPPath = Phosphorus.Factory.GetNewPhosphorusPPath
    NestedPPath.Initialise
    Set EvaluatedNestedPPath = NestedPPath.Evaluate(lstrNestedPPath, eleCurrentContextUIElement, strInitialPPath)
    'Raise any error?
    This.PPathReturnClass.SetErrorMessage = EvaluatedNestedPPath.GetErrorMessage
    varReturnValue = EvaluatedNestedPPath.ReturnedValue
    lstrCurrentPredicateTest = VBA.Strings.Replace(lstrCurrentPredicateTest, lstrNestedPPath, varReturnValue)
    Set NestedPPath = Nothing
    Set EvaluatedNestedPPath = Nothing
  End If
End Function

Private Sub ProcessAttributeChecks( _
  eleCurrentContextUIElement As UIAutomationClient.IUIAutomationElement, _
  strInitialPPath As String)
  
  While SourcePPathContainedAnAttributeCheck(eleCurrentContextUIElement, strInitialPPath)
  Wend
  lstrCurrentPredicateTest = VBA.Strings.Replace(lstrCurrentPredicateTest, "{AtSign}", "@")
End Sub

Private Function SourcePPathContainedAnAttributeCheck( _
  eleCurrentContextUIElement As UIAutomationClient.IUIAutomationElement, _
  strInitialPPath As String) As Boolean

  SourcePPathContainedAnAttributeCheck = False
  Dim strAttributeName As String
  Dim intStartOfAttributeName As Integer
  intStartOfAttributeName = VBA.Strings.InStr(1, lstrCurrentPredicateTest, "@")
  If intStartOfAttributeName > 0 Then
    SourcePPathContainedAnAttributeCheck = True
    Dim intEndOfAttributeName As Integer
    Dim boolInsideAPredicate As Boolean
    boolInsideAPredicate = False
    For intEndOfAttributeName = (intStartOfAttributeName + 1) To VBA.Strings.Len(lstrCurrentPredicateTest)
      Dim strCurrentCharacter As String
      strCurrentCharacter = VBA.Strings.Mid(lstrCurrentPredicateTest, intEndOfAttributeName, 1)
      If strCurrentCharacter = "[" Then
        boolInsideAPredicate = True
      Else
        If strCurrentCharacter = "]" Then
          boolInsideAPredicate = False
        End If
      End If
      If Not boolInsideAPredicate And VBA.Strings.InStr(1, ",=><", strCurrentCharacter) > 0 Then
        strAttributeName = VBA.Strings.Mid(lstrCurrentPredicateTest, intStartOfAttributeName + 1, intEndOfAttributeName - intStartOfAttributeName - 1)
        Exit For
      End If
    Next intEndOfAttributeName
  End If
  If SourcePPathContainedAnAttributeCheck Then
    Dim Key As Variant
    Dim strLookupAttributeName As String
    For Each Key In This.AutomationDictionaries.NavigablePropertyIDs.Keys
      strLookupAttributeName = This.AutomationDictionaries.NavigablePropertyIDs(Key)
      If strLookupAttributeName = strAttributeName Then
        Dim varPropertyValue As Variant
        If This.UnitTestingMode And ((strAttributeName = "ProcessId") Or (strAttributeName = "ProviderDescription") Or (strAttributeName = "BoundingRectangle") Or (strAttributeName = "NativeWindowHandle")) Then
          varPropertyValue = "#"
        Else
          varPropertyValue = eleCurrentContextUIElement.GetCurrentPropertyValue(Key)
          If Not IsEmpty(varPropertyValue) Then
            If varPropertyValue <> "" Then
              If VBA.Strings.InStr(1, varPropertyValue, "@") > 0 Then
                varPropertyValue = VBA.Strings.Replace(varPropertyValue, "@", "{AtSign}")
              End If
              If VBA.Strings.InStr(1, varPropertyValue, vbLf) > 0 Then
                varPropertyValue = VBA.Strings.Replace(varPropertyValue, vbLf, "\n")
              End If
            End If
          End If
          
          If IsArray(varPropertyValue) Then
            If UBound(varPropertyValue) = -1 Then
              varPropertyValue = ""
            Else
              If Not IsEmpty(varPropertyValue) Then
                Dim i As Integer
                Dim strArrayOfValues As String
                strArrayOfValues = ""
                For i = 0 To UBound(varPropertyValue)
                  If strArrayOfValues <> "" Then
                    strArrayOfValues = strArrayOfValues & ", "
                  End If
                  strArrayOfValues = strArrayOfValues & varPropertyValue(i)
                  Next i
                varPropertyValue = "{" & strArrayOfValues & "}"
              End If
            End If
          End If
        End If

If IsEmpty(varPropertyValue) Then
 ' Debug.Print varPropertyValue
  varPropertyValue = ""
ElseIf varPropertyValue = "" Then
'  Debug.Print varPropertyValue
'  DoEvents
Else
  If varPropertyValue <> 0 Then
'    Debug.Print varPropertyValue
'    DoEvents
  End If
End If

        If IsNumeric(varPropertyValue) Then
          'Do nothing
          lstrCurrentPredicateTest = VBA.Strings.Replace(lstrCurrentPredicateTest, "@" & strAttributeName, varPropertyValue)
        Else
          'Is the string too long for the evaluate function - max 255 characters
          If VBA.Strings.Len(lstrCurrentPredicateTest) - VBA.Strings.Len("@" & strAttributeName) + VBA.Strings.Len(varPropertyValue) > 255 Then
            Dim intStartOfTargetAttibuteValue As Integer
            Dim intEndOfTargetAttibuteValue As Integer
            Dim strComparisonOperator As String
            strComparisonOperator = ""
            Dim boolAttributeValueStarted As Boolean
            Dim boolAttributeValueIsString As Boolean
            Dim boolInsideAttributeString As Boolean
            intStartOfTargetAttibuteValue = VBA.Strings.InStr(1, lstrCurrentPredicateTest, "@" & strAttributeName) + VBA.Strings.Len("@" & strAttributeName)
            For intEndOfTargetAttibuteValue = intStartOfTargetAttibuteValue To VBA.Strings.Len(lstrCurrentPredicateTest)
              Dim strCurrentTargetAttibuteValueCharacter As String
              strCurrentTargetAttibuteValueCharacter = VBA.Strings.Mid(lstrCurrentPredicateTest, intEndOfTargetAttibuteValue, 1)
              If (strComparisonOperator = "") And VBA.Strings.InStr(1, "=><", strCurrentTargetAttibuteValueCharacter) > 0 Then
                strComparisonOperator = strComparisonOperator & strCurrentTargetAttibuteValueCharacter
                Dim strNextAttibuteValueCharacter As String
                strNextAttibuteValueCharacter = VBA.Strings.Mid(lstrCurrentPredicateTest, intEndOfTargetAttibuteValue + 1, 1)
                If VBA.Strings.InStr(1, "=><", strNextAttibuteValueCharacter) > 0 Then
                  strComparisonOperator = strComparisonOperator & strNextAttibuteValueCharacter
                  intEndOfTargetAttibuteValue = intEndOfTargetAttibuteValue + 1
                End If
                intStartOfTargetAttibuteValue = intStartOfTargetAttibuteValue + VBA.Strings.Len(strComparisonOperator)
              Else
                If Not boolAttributeValueStarted And (strCurrentTargetAttibuteValueCharacter = """") Then
                  'Start of a string
                  intStartOfTargetAttibuteValue = intStartOfTargetAttibuteValue + 1
                  boolAttributeValueStarted = True
                  boolAttributeValueIsString = True
                  boolInsideAttributeString = True
                ElseIf boolAttributeValueStarted And (strCurrentTargetAttibuteValueCharacter = """") Then
                  'Second/Closing quote so exit here
                  intEndOfTargetAttibuteValue = intEndOfTargetAttibuteValue - 1
                  '??Nested quotes?
                   Exit For
               ElseIf boolInsideAttributeString And (strCurrentTargetAttibuteValueCharacter = " ") Then
                  'Skip spaces inside strings
                ElseIf Not boolAttributeValueIsString And (strCurrentTargetAttibuteValueCharacter = " ") Then
                  'End of non-string value so exit here
                   Exit For
                Else
                  'Jus skip tp next character
                  boolAttributeValueStarted = True
                End If
              End If
            
            Next intEndOfTargetAttibuteValue
            Dim strAttributeValue As String
            If intEndOfTargetAttibuteValue = VBA.Strings.Len(lstrCurrentPredicateTest) Then
              strAttributeValue = VBA.Strings.Mid(lstrCurrentPredicateTest, intStartOfTargetAttibuteValue, VBA.Strings.Len(lstrCurrentPredicateTest))
            Else
              strAttributeValue = VBA.Strings.Mid(lstrCurrentPredicateTest, intStartOfTargetAttibuteValue, intEndOfTargetAttibuteValue - intStartOfTargetAttibuteValue + 1)
            End If
    
            Dim strSourcePredicateString As String
            strSourcePredicateString = "@" & strAttributeName & strComparisonOperator
            If boolAttributeValueIsString Then
              strSourcePredicateString = strSourcePredicateString & VBA.Strings.Chr(34) & strAttributeValue & VBA.Strings.Chr(34)
            Else
              strSourcePredicateString = strSourcePredicateString & strAttributeValue
            End If
    
            If varPropertyValue = strAttributeValue Then
              lstrCurrentPredicateTest = VBA.Strings.Replace(lstrCurrentPredicateTest, strSourcePredicateString, "TRUE")
            Else
              lstrCurrentPredicateTest = VBA.Strings.Replace(lstrCurrentPredicateTest, strSourcePredicateString, "FALSE")
            End If
          Else
            varPropertyValue = VBA.Strings.Chr(34) & varPropertyValue & VBA.Strings.Chr(34)
            lstrCurrentPredicateTest = VBA.Strings.Replace(lstrCurrentPredicateTest, "@" & strAttributeName, varPropertyValue)
          End If
        End If
        Exit For
      End If
    Next Key
  
  End If
  
End Function

Private Sub ProcessTextAndValueChecks(eleCurrentContextUIElement As UIAutomationClient.IUIAutomationElement)
  If VBA.Strings.InStr(1, lstrCurrentPredicateTest, "text()") > 0 Then
    Dim strCurrentText As String
    strCurrentText = This.NodeTests.GetTextValue(eleCurrentContextUIElement)
    lstrCurrentPredicateTest = VBA.Strings.Replace(lstrCurrentPredicateTest, "text()", VBA.Strings.Chr(34) & strCurrentText & VBA.Strings.Chr(34))
  End If
  If VBA.Strings.InStr(1, lstrCurrentPredicateTest, "value()") > 0 Then
    Dim varCurrentValue As Variant
    varCurrentValue = PPathUtils.GetValue(eleCurrentContextUIElement)
    lstrCurrentPredicateTest = VBA.Strings.Replace(lstrCurrentPredicateTest, "value()", varCurrentValue)
  End If
End Sub
