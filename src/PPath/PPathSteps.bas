VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPathSteps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder PPath
Option Explicit
Option Base 1

Private this As PPathCommon

Public Sub Initialise(ByRef sharedthis As PPathCommon)
  Set this = sharedthis
End Sub

Public Function GetAllXPathExpressions(ByVal strFullLocationPathExpression As String) As String()
  strFullLocationPathExpression = VBA.Strings.Trim(strFullLocationPathExpression)
  If VBA.Strings.Left(strFullLocationPathExpression, 1) = "(" And _
     VBA.Strings.Right(strFullLocationPathExpression, 1) = ")" Then
     strFullLocationPathExpression = VBA.Strings.Mid(strFullLocationPathExpression, 2, VBA.Strings.Len(strFullLocationPathExpression) - 2)
  End If
  Dim strReturn() As String
  If strFullLocationPathExpression = "" Then
    ReDim strReturn(1)
    strReturn(1) = "" 'Unnecessary, but adds clarity
  Else
    Dim strReturnSplit() As String
    strReturnSplit = VBA.Strings.Split(strFullLocationPathExpression, "|")
    ReDim strReturn(UBound(strReturnSplit) + 1)
    Dim i As Integer
    For i = 1 To UBound(strReturn)
      'Trim any spaces at start or end
      strReturn(i) = VBA.Strings.Trim(strReturnSplit(i - 1))
      'Remove any open parentheses - these can be ignored
      Dim intCharacterCounter As Integer
      intCharacterCounter = 0
      While VBA.Strings.Mid(strReturn(i), intCharacterCounter + 1, 1) = "("
        intCharacterCounter = intCharacterCounter + 1
      Wend
      If intCharacterCounter > 0 Then
        strReturn(i) = VBA.Strings.Mid(strReturn(i), intCharacterCounter + 1, VBA.Strings.Len(strReturn(i)))
      End If
    Next i
  End If
  GetAllXPathExpressions = strReturn
End Function

Public Function GetNextStep(strRemainingPPath As String) As Phosphorus.PPathStep

  Dim NextStep As New Phosphorus.PPathStep
  Dim intCharacter As Integer
  Dim boolContinue As Boolean

  'Get Next Axis
  intCharacter = intCharacter + 1
    
  If VBA.Strings.Left(strRemainingPPath, 2) = Axes.DescendantShorthand Then
    NextStep.Axis = Axes.DescendantShorthand
    strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, 3, VBA.Strings.Len(strRemainingPPath))
      
  ElseIf VBA.Strings.Left(strRemainingPPath, 1) = "/" Then
     
    'New Step - Remove the leading step character
    strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, 2)
      
    'Child
    If VBA.Strings.InStr(strRemainingPPath, Axes.Child) = 1 Then
      NextStep.Axis = Axes.ChildShorthand
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, VBA.Len(Axes.Child) + 1)
      
    'Descendants
    ElseIf VBA.Strings.InStr(strRemainingPPath, Axes.Descendant) = 1 Then
      NextStep.Axis = Axes.DescendantShorthand
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, VBA.Len(Axes.Descendant) + 1)
    ElseIf VBA.Strings.InStr(strRemainingPPath, Axes.DescendantOrSelf) = 1 Then
      NextStep.Axis = Axes.DescendantOrSelf
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, VBA.Len(Axes.DescendantOrSelf) + 1)
       
    'Parents
    'NOTE: We must check for the parent axes ".." before checking for the self axis "."
    ElseIf VBA.Strings.InStr(strRemainingPPath, Axes.Parent) = 1 Then
      NextStep.Axis = Axes.ParentShorthand
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, VBA.Len(Axes.Parent) + 1)
    ElseIf VBA.Strings.InStr(strRemainingPPath, Axes.ParentShorthand) = 1 Then
      NextStep.Axis = Axes.ParentShorthand
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, VBA.Len(Axes.ParentShorthand) + 1)
        
    'Self only
    ElseIf VBA.Strings.InStr(strRemainingPPath, Axes.SelfShorthand) = 1 Then
      NextStep.Axis = Axes.SelfShorthand
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, VBA.Len(Axes.SelfShorthand) + 1)
    ElseIf VBA.Strings.InStr(strRemainingPPath, Axes.Self) = 1 Then
      NextStep.Axis = Axes.SelfShorthand
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, VBA.Len(Axes.Self) + 1)
      
    'Ancestor Axes
    ElseIf VBA.Strings.InStr(strRemainingPPath, Axes.Ancestor) = 1 Then
      NextStep.Axis = Axes.Ancestor
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, VBA.Len(Axes.Ancestor) + 1)
    ElseIf VBA.Strings.InStr(strRemainingPPath, Axes.AncestorOrSelf) = 1 Then
      NextStep.Axis = Axes.AncestorOrSelf
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, VBA.Len(Axes.AncestorOrSelf) + 1)
      
    'Siblings
    ElseIf VBA.Strings.InStr(strRemainingPPath, Axes.PrecedingSibling) = 1 Then
      NextStep.Axis = Axes.PrecedingSibling
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, VBA.Len(Axes.PrecedingSibling) + 1)
    ElseIf VBA.Strings.InStr(strRemainingPPath, Axes.FollowingSibling) = 1 Then
      NextStep.Axis = Axes.FollowingSibling
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, VBA.Len(Axes.FollowingSibling) + 1)
      
    'Preceding
    ElseIf (VBA.Strings.InStr(strRemainingPPath, Axes.Preceding) = 1) Then
      NextStep.Axis = Axes.Preceding
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, VBA.Len(Axes.Preceding) + 1)
      
    'Following
    ElseIf (VBA.Strings.InStr(strRemainingPPath, Axes.Following) = 1) Then
      NextStep.Axis = Axes.Following
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, VBA.Len(Axes.Following) + 1)
      
    'Attribute
    ElseIf (VBA.Strings.InStr(strRemainingPPath, Axes.Attribute) = 1) Then
      NextStep.Axis = Axes.Attribute
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, VBA.Len(Axes.Following) + 1)
      
    Else
      
      'No recognised axis so this must be an implicit default child axis
      NextStep.Axis = Axes.ChildShorthand
      
    End If
  
  End If
    
  intCharacter = 0
        
  'Get Next Node Test
  If VBA.Strings.Left(strRemainingPPath, 1) = "*" Then
    NextStep.NodeTest = "*"
    strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, 2, VBA.Strings.Len(strRemainingPPath))
  ElseIf VBA.Strings.InStr(1, strRemainingPPath, "node()") = 1 Then
    NextStep.NodeTest = "node()"
    strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, 7, VBA.Strings.Len(strRemainingPPath))
  ElseIf VBA.Strings.InStr(1, strRemainingPPath, "element()") = 1 Then
    NextStep.NodeTest = "element()"
    strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, 10, VBA.Strings.Len(strRemainingPPath))
  'Element by kind
  ElseIf VBA.Strings.InStr(1, strRemainingPPath, "element(") = 1 Then
    Dim strParameters As String
    Dim intNextClosingBracket As Integer
    intNextClosingBracket = VBA.Strings.InStr(VBA.Strings.Len("element(") + 1, strRemainingPPath, ")")
    strParameters = VBA.Strings.Mid(strRemainingPPath, VBA.Strings.Len("element(") + 1, intNextClosingBracket - VBA.Strings.Len("element(") - 1)
    Dim arrParameters() As String
    arrParameters = VBA.Strings.Split(strParameters, ",")
    NextStep.NodeTest = VBA.Strings.Trim(arrParameters(0))
    NextStep.NodeTestKind = VBA.Strings.Trim(arrParameters(1))
    strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, intNextClosingBracket + 1, VBA.Strings.Len(strRemainingPPath))
  ElseIf VBA.Strings.InStr(1, strRemainingPPath, "text()") = 1 Then
    NextStep.NodeTest = "text()"
    strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, 7, VBA.Strings.Len(strRemainingPPath))
  ElseIf strRemainingPPath = "" Then
    this.PPathReturnClass.SetErrorMessage = Phosphorus.PPathConstants.NO_NODETEST_PPATH_ERROR_MESSAGE
    this.PPathReturnClass.RemoveAllMatchingElements
  Else
    For intCharacter = 1 To VBA.Strings.Len(strRemainingPPath)
      Dim intEndOfNodeTest As Integer
      Dim strCurrentCharacter As String
      strCurrentCharacter = VBA.Strings.Mid(strRemainingPPath, intCharacter, 1)
      If VBA.Strings.InStr(1, "[/)", strCurrentCharacter) > 0 Then
        intEndOfNodeTest = intCharacter
        Exit For
      Else
        'Skip this character
      End If
    Next intCharacter
    If intEndOfNodeTest = 0 Then
      NextStep.NodeTest = strRemainingPPath
      strRemainingPPath = ""
    Else
      NextStep.NodeTest = VBA.Strings.Mid(strRemainingPPath, 1, intEndOfNodeTest - 1)
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, intEndOfNodeTest, VBA.Strings.Len(strRemainingPPath))
    End If
    'Check for valid node names
    Dim key As Variant
    Dim IsAValidNodeTestType As Boolean
    If NextStep.Axis = Axes.Attribute Then
      IsAValidNodeTestType = this.AutomationDictionaries.ValueExists(NextStep.NodeTest, this.AutomationDictionaries.NavigablePropertyIDs)
    Else
      IsAValidNodeTestType = this.AutomationDictionaries.ValueExists(NextStep.NodeTest, this.AutomationDictionaries.ControlTypeIDs)
    End If
    If Not IsAValidNodeTestType Then
      this.PPathReturnClass.SetErrorMessage = Phosphorus.PPathConstants.INVALID_NODETEST_PPATH_ERROR_MESSAGE & " '" & NextStep.NodeTest & "'!"
      Exit Function
    End If
  End If
        
  'Get Predicates
  If strRemainingPPath = "" Then
    'End of step - skip to end of step
  ElseIf VBA.Strings.Left(strRemainingPPath, 1) = "/" Then
    'End of step, no predicates - skip to end of step
  ElseIf VBA.Strings.Left(strRemainingPPath, 1) = "[" Or VBA.Strings.Left(strRemainingPPath, 1) = ")" Then
    'Start of Predicates
    Dim intStartOfNextStep  As Integer
    intStartOfNextStep = 0
    Dim boolInsidePredicate As Boolean
    boolInsidePredicate = False
    For intCharacter = 1 To VBA.Strings.Len(strRemainingPPath)
      strCurrentCharacter = VBA.Strings.Mid(strRemainingPPath, intCharacter, 1)
      If Not boolInsidePredicate And strCurrentCharacter = "[" Then
        boolInsidePredicate = True
      Else
        If boolInsidePredicate And strCurrentCharacter = "]" Then
          boolInsidePredicate = False
        Else
          If boolInsidePredicate Then
            'Do nothing
          Else
            If Not boolInsidePredicate And strCurrentCharacter = "/" Then
              intStartOfNextStep = intCharacter
              Exit For
            End If
           End If
         End If
      End If
    Next intCharacter
      
    Dim strNextPredicates As String
    If intStartOfNextStep = 0 Then
      'No more steps so we expect only predicates
      strNextPredicates = strRemainingPPath
      strRemainingPPath = ""
    Else
      'Get next predicates up to start of next step
      strNextPredicates = VBA.Strings.Left(strRemainingPPath, intStartOfNextStep - 1)
      strRemainingPPath = VBA.Strings.Mid(strRemainingPPath, intStartOfNextStep, VBA.Strings.Len(strRemainingPPath))
    End If
      
    NextStep.AddPredicates strNextPredicates
      
  Else
    this.PPathReturnClass.SetErrorMessage = _
      Phosphorus.PPathConstants.ILLEGAL_START_OF_PREDICATE_ERROR_MESSAGE & _
      " '" & VBA.Strings.Left(strRemainingPPath, 1) & "'"
    this.PPathReturnClass.RemoveAllMatchingElements
  End If

  NextStep.RemainingPPath = strRemainingPPath
  
  Set GetNextStep = NextStep
 
End Function
