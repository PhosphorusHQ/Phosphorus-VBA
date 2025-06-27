VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPathAxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder PPath
Option Explicit

Private this As PPathCommon

Public Sub Initialise(ByRef sharedthis As PPathCommon)
  Set this = sharedthis
End Sub

Public Sub ProcessNextAxis( _
  myNextStep As Phosphorus.PPathStep)
      
  'The method loops through each element in the current Working Copy and call the axis method(s) relevant to each axis
  Dim intElementCounter As Integer
  Dim eleCurrentUIElement As UIAutomationClient.IUIAutomationElement
  Dim strInitialPPath As String
  Dim intNumberOfWorkingCopyElements As Integer
  intNumberOfWorkingCopyElements = this.PPathReturnClass.GetNumberOfWorkingCopyOfCandidateElements(this.CurrentLocationPathExpressionCounter)
  For intElementCounter = 1 To intNumberOfWorkingCopyElements

    Set eleCurrentUIElement = this.PPathReturnClass.GetWorkingCopyElement(this.CurrentLocationPathExpressionCounter, intElementCounter)
    strInitialPPath = this.PPathReturnClass.GetWorkingCopyNavigationalPPath(this.CurrentLocationPathExpressionCounter, intElementCounter)

    'Process any axis requiring the self node
    If (myNextStep.Axis = Axes.SelfShorthand) Or _
       (myNextStep.Axis = Axes.DescendantOrSelf) Or _
       (myNextStep.Axis = Axes.AncestorOrSelf) Then
      GetSelfNodes eleCurrentUIElement, strInitialPPath & Axes.SelfShorthand
    End If
  
    'Child
    If (myNextStep.Axis = Axes.ChildShorthand) Then
      GetChildren eleCurrentUIElement, Axes.ChildShorthand, strInitialPPath & Axes.ChildShorthand, False
  
    'Descendant
    ElseIf (myNextStep.Axis = Axes.DescendantShorthand) Or (myNextStep.Axis = Axes.DescendantOrSelf) Then
      GetChildren eleCurrentUIElement, Axes.ChildShorthand, strInitialPPath & Axes.ChildShorthand, True
 
    'Parent
    ElseIf myNextStep.Axis = Axes.ParentShorthand Then
      GetParents eleCurrentUIElement, Axes.ParentShorthand, strInitialPPath & Axes.ParentShorthand, False
  
    'Ancestors
    ElseIf (myNextStep.Axis = Axes.Ancestor) Or (myNextStep.Axis = Axes.AncestorOrSelf) Then
      GetParents eleCurrentUIElement, Axes.ParentShorthand, strInitialPPath & Axes.ParentShorthand, True
  
    'Siblings
    ElseIf (myNextStep.Axis = Axes.PrecedingSibling) Then
      GetSiblings eleCurrentUIElement, myNextStep.Axis, strInitialPPath
    ElseIf (myNextStep.Axis = Axes.FollowingSibling) Then
      GetSiblings eleCurrentUIElement, myNextStep.Axis, strInitialPPath
  
    'Preceding
    ElseIf (myNextStep.Axis = Axes.Preceding) Then
      GetPrecedingOrFollowing eleCurrentUIElement, myNextStep.Axis
  
    'Following
    ElseIf (myNextStep.Axis = Axes.Following) Then
      GetPrecedingOrFollowing eleCurrentUIElement, myNextStep.Axis
  
    'Attributes
    ElseIf (myNextStep.Axis = Axes.Attribute) Then
      GetAttributes eleCurrentUIElement, strInitialPPath
  
    End If
  
  Next
  
End Sub

Private Sub GetSelfNodes( _
  eleCurrentUIElement As UIAutomationClient.IUIAutomationElement, _
  strInitialXPath As String)
  
  Dim strCurrentElementAbsoluteXPath As String
  strCurrentElementAbsoluteXPath = strInitialXPath
  PPathUtils.OutputElementDetails eleCurrentUIElement, strCurrentElementAbsoluteXPath, this.DebugMode, this.AutomationDictionaries
  this.PPathReturnClass.AddMatchingElement this.PPathReturnClass.GetCandidateElements(this.CurrentLocationPathExpressionCounter), eleCurrentUIElement, "", strCurrentElementAbsoluteXPath
End Sub

Private Sub GetParents( _
  eleCurrentUIElement As UIAutomationClient.IUIAutomationElement, _
  strCurrentStepAxis As String, _
  strInitialPPath As String, _
  boolGetAllAncestors As Boolean, _
  Optional ByRef AllAncestorsOrSelfRuntimeIDs As collection)
  
  'The parent:: axis is a reverse axis, but as we only get one node it doesn't make any difference which order it's in!?
  'Also, this method already returns ancestor nodes in reverse document order, so there is no need to reverse.
  
  'Never go above the current application root element!
  Dim boolCurrentIsApplicationRootUIElement As Boolean
  boolCurrentIsApplicationRootUIElement = False
  If PPathRuntimeIDs.GetElementRuntimeID(eleCurrentUIElement) = PPathRuntimeIDs.GetElementRuntimeID(this.ApplicationRootUIElement) Then
    boolCurrentIsApplicationRootUIElement = True
  End If
    
  Dim eleParentUIElement As UIAutomationClient.IUIAutomationElement
  Set eleParentUIElement = this.TreeWalker.GetParentElement(eleCurrentUIElement)
    
  Dim strParentUIElementCurrentUIElementRuntimeID As String
  strParentUIElementCurrentUIElementRuntimeID = PPathRuntimeIDs.GetElementRuntimeID(eleParentUIElement)
          
  Dim strCurrentControlType As String
  Dim strCurrentElementAbsoluteXPath As String
  
  Dim boolAddElement As Boolean
  boolAddElement = False
    
  strCurrentControlType = this.AutomationDictionaries.ControlTypeIDs(eleParentUIElement.CurrentControlType)
  strCurrentElementAbsoluteXPath = strInitialPPath & "/" & strCurrentControlType & "[@Name='" & eleParentUIElement.CurrentName & "']"
    
  If strCurrentStepAxis = Axes.ParentShorthand Then
    boolAddElement = True
  End If
      
  If boolAddElement And Not boolCurrentIsApplicationRootUIElement Then
    OutputElementDetails eleParentUIElement, strCurrentElementAbsoluteXPath, this.DebugMode, this.AutomationDictionaries
    this.PPathReturnClass.AddMatchingElement this.PPathReturnClass.GetCandidateElements(this.CurrentLocationPathExpressionCounter), eleParentUIElement, "", strCurrentElementAbsoluteXPath
  End If
    
  If Not boolCurrentIsApplicationRootUIElement Then
    GetParents eleParentUIElement, strCurrentStepAxis, strCurrentElementAbsoluteXPath & "/..", boolGetAllAncestors, AllAncestorsOrSelfRuntimeIDs
  End If
       
End Sub

Private Sub GetSiblings( _
  eleCurrentUIElement As UIAutomationClient.IUIAutomationElement, _
  Axis As String, _
  strInitialPPath As String)
  
  Dim strCurrentElementRuntimeID As String
  strCurrentElementRuntimeID = PPathRuntimeIDs.GetElementRuntimeID(eleCurrentUIElement)
  
  Dim eleParentUIElement As UIAutomationClient.IUIAutomationElement
  Set eleParentUIElement = this.TreeWalker.GetParentElement(eleCurrentUIElement)
  
  Dim eleChildrenUIElementArray As UIAutomationClient.IUIAutomationElementArray
  Set eleChildrenUIElementArray = eleParentUIElement.FindAll(UIAutomationClient.TreeScope.TreeScope_Children, this.UIAutomation.CreateTrueCondition())
  
  Dim intListItemCounter As Integer
  Dim eleChildUIElement As UIAutomationClient.IUIAutomationElement
  Dim boolBeforeCurrentElement As Boolean
  Dim boolIsCurrentElement As Boolean
  Dim boolAfterCurrentElement As Boolean
  boolBeforeCurrentElement = True
  boolIsCurrentElement = False
  boolAfterCurrentElement = False
  Dim controls As Scripting.dictionary
  Dim strCurrentControlType As String
  Dim intInstanceNumber As Integer
  Dim strInstancePredicate As String
  Set controls = New Scripting.dictionary
  For intListItemCounter = 0 To eleChildrenUIElementArray.Length - 1
    
    Set eleChildUIElement = eleChildrenUIElementArray.GetElement(intListItemCounter)
    Dim strCurrentChildElementRuntimeID As String
    strCurrentChildElementRuntimeID = PPathRuntimeIDs.GetElementRuntimeID(eleChildUIElement)
    strCurrentControlType = this.AutomationDictionaries.ControlTypeIDs(eleChildUIElement.CurrentControlType)
    If controls.exists(strCurrentControlType) Then
      intInstanceNumber = controls(strCurrentControlType) + 1
      controls(strCurrentControlType) = intInstanceNumber
    Else
      intInstanceNumber = 1
      controls.Add strCurrentControlType, intInstanceNumber
    End If
    strInstancePredicate = "[" & intInstanceNumber & "]"
    Dim strCurrentElementAbsoluteXPath As String
    strCurrentElementAbsoluteXPath = strInitialPPath & "../" & strCurrentControlType & strInstancePredicate & "[@Name='" & eleChildUIElement.CurrentName & "']"
      
    boolIsCurrentElement = False
    If strCurrentChildElementRuntimeID = strCurrentElementRuntimeID Then
      boolBeforeCurrentElement = False
      boolIsCurrentElement = True
      boolAfterCurrentElement = True
    End If

    'Add preceding siblings straight to Temp Candidate Elements as they are in reverse order
    If (boolBeforeCurrentElement And Axis = Axes.PrecedingSibling) Then
      OutputElementDetails eleChildUIElement, strCurrentElementAbsoluteXPath, this.DebugMode, this.AutomationDictionaries
      this.PPathReturnClass.AddMatchingElement this.PPathReturnClass.GetTempCandidateElements(this.CurrentLocationPathExpressionCounter), eleChildUIElement, "", strCurrentElementAbsoluteXPath
    End If
      
    'Now promote preceding siblings to Candidate Elements as in reverse order
    If (boolIsCurrentElement And Axis = Axes.PrecedingSibling) Then
      this.PPathReturnClass.PromoteTempCandidateElementsToCandidateElementsInReverseOrder this.CurrentLocationPathExpressionCounter
    End If
    
    'Add following siblings straight to Candidate Elements as they are in forward order
    If (boolAfterCurrentElement And Axis = Axes.FollowingSibling And Not boolIsCurrentElement) Then
      OutputElementDetails eleChildUIElement, strCurrentElementAbsoluteXPath, this.DebugMode, this.AutomationDictionaries
      this.PPathReturnClass.AddMatchingElement this.PPathReturnClass.GetCandidateElements(this.CurrentLocationPathExpressionCounter), eleChildUIElement, "", strCurrentElementAbsoluteXPath
    End If
  
  Next intListItemCounter
  
End Sub

Private Sub GetChildren( _
  eleCurrentUIElement As UIAutomationClient.IUIAutomationElement, _
  Axis As String, _
  strInitialPPath As String, _
  boolGetGrandChildren As Boolean, _
  Optional ByRef AllAncestorsOrSelfRuntimeIDs As collection, _
  Optional ByRef boolPrecedingOrFollowingRootElementFound As Boolean)
  
  Dim eleChildrenUIElementArray As UIAutomationClient.IUIAutomationElementArray
  Dim eleArrayOfChildrenUIElements() As UIAutomationClient.IUIAutomationElement
  Set eleChildrenUIElementArray = eleCurrentUIElement.FindAll(UIAutomationClient.TreeScope.TreeScope_Children, this.UIAutomation.CreateTrueCondition())
    
  If eleChildrenUIElementArray.Length = 0 Then
'    OutputElementDetails eleCurrentUIElement, strInitialXPath
'    lclsReturnItem.AddMatchingElement eleCurrentUIElement, strInitialXPath
  Else
  
    Dim boolAddChild As Boolean
    boolAddChild = False
    
    'https://excelmacromastery.com/vba-dictionary/
    '“Microsoft Scripting Runtime”(Add using Tools->References from the VB menu
    Dim controls As Scripting.dictionary
    Dim strCurrentControlType As String
    Dim intInstanceNumber As Integer
    Dim strInstancePredicate As String
    Set controls = New Scripting.dictionary
    
    Dim intListItemCounter As Integer
    Dim eleChildUIElement As UIAutomationClient.IUIAutomationElement
    Dim strCurrentElementAbsoluteXPath As String
    
    For intListItemCounter = 0 To eleChildrenUIElementArray.Length - 1
      Set eleChildUIElement = eleChildrenUIElementArray.GetElement(intListItemCounter)
      'If intListItemCounter = 0 Then
      Dim intCurrentSizeOfArray As Integer
      intCurrentSizeOfArray = Utils.GetSizeOfArray(eleArrayOfChildrenUIElements)
      intCurrentSizeOfArray = intCurrentSizeOfArray + 1
      If intCurrentSizeOfArray = 1 Then
        ReDim eleArrayOfChildrenUIElements(intCurrentSizeOfArray)
      Else
        ReDim Preserve eleArrayOfChildrenUIElements(intCurrentSizeOfArray)
      End If
      Set eleArrayOfChildrenUIElements(intCurrentSizeOfArray) = eleChildUIElement
      strCurrentControlType = this.AutomationDictionaries.ControlTypeIDs(eleChildUIElement.CurrentControlType)
      If controls.exists(strCurrentControlType) Then
        intInstanceNumber = controls(strCurrentControlType) + 1
        controls(strCurrentControlType) = intInstanceNumber
      Else
        intInstanceNumber = 1
        controls.Add strCurrentControlType, intInstanceNumber
      End If
      
      strInstancePredicate = "[" & intInstanceNumber & "]"
      strCurrentElementAbsoluteXPath = strInitialPPath & strCurrentControlType & strInstancePredicate & "[@Name='" & eleChildUIElement.CurrentName & "']"
            
      If Axis = Axes.ChildShorthand Then
        boolAddChild = True
      End If
      
      Dim strCurrentElementRuntimeID As String
      If (Axis = Axes.Preceding) Or (Axis = Axes.Following) Then
      
        strCurrentElementRuntimeID = PPathRuntimeIDs.GetElementRuntimeID(eleChildUIElement)
        If strCurrentElementRuntimeID = AllAncestorsOrSelfRuntimeIDs(1) Then
          boolPrecedingOrFollowingRootElementFound = True
          'Now promote preceding elements to Candidate Elements as in reverse order
          If (Axis = Axes.Preceding) Then
            this.PPathReturnClass.PromoteTempCandidateElementsToCandidateElementsInReverseOrder this.CurrentLocationPathExpressionCounter
          End If
        End If
             
        Dim boolCurrentElementIsAnAncestorOrSelfOfPrecedingOrFollowingRootElement As Boolean
        boolCurrentElementIsAnAncestorOrSelfOfPrecedingOrFollowingRootElement = False
        Dim i As Integer
        For i = 1 To AllAncestorsOrSelfRuntimeIDs.Count
          If strCurrentElementRuntimeID = AllAncestorsOrSelfRuntimeIDs(i) Then
            boolCurrentElementIsAnAncestorOrSelfOfPrecedingOrFollowingRootElement = True
            Exit For
          End If
        Next i
        
        If Not boolCurrentElementIsAnAncestorOrSelfOfPrecedingOrFollowingRootElement Then
        
          If (Axis = Axes.Preceding) And Not boolPrecedingOrFollowingRootElementFound Then
            'Add preceding siblings straight to Temp Candidate Elements as they are in reverse order
            OutputElementDetails eleChildUIElement, strCurrentElementAbsoluteXPath, this.DebugMode, this.AutomationDictionaries
            this.PPathReturnClass.AddMatchingElement this.PPathReturnClass.GetTempCandidateElements(this.CurrentLocationPathExpressionCounter), eleChildUIElement, "", strCurrentElementAbsoluteXPath
          End If
          
          If (Axis = Axes.Following) And boolPrecedingOrFollowingRootElementFound Then
            boolAddChild = True
          End If
          
        End If
        
      End If
      
      If boolAddChild Then
        OutputElementDetails eleChildUIElement, strCurrentElementAbsoluteXPath, this.DebugMode, this.AutomationDictionaries
        this.PPathReturnClass.AddMatchingElement this.PPathReturnClass.GetCandidateElements(this.CurrentLocationPathExpressionCounter), eleChildUIElement, "", strCurrentElementAbsoluteXPath
      End If
      
      If boolGetGrandChildren And Not boolCurrentElementIsAnAncestorOrSelfOfPrecedingOrFollowingRootElement Then
        GetChildren eleChildUIElement, Axis, strCurrentElementAbsoluteXPath & Axes.ChildShorthand, boolGetGrandChildren, AllAncestorsOrSelfRuntimeIDs, boolPrecedingOrFollowingRootElementFound
      End If
      
      'Reset add child for next item in loop
      boolAddChild = False
      
    Next intListItemCounter
    Set controls = Nothing
  End If

End Sub

Private Sub GetPrecedingOrFollowing( _
  eleCurrentUIElement As UIAutomationClient.IUIAutomationElement, _
  Axis As String)
  
  'Create a collection of self or parent runtime ID's
  Dim colAllAncestorsOrSelfRuntimeIDs As collection
  Set colAllAncestorsOrSelfRuntimeIDs = New collection
  
  'Add the current element to the list
  colAllAncestorsOrSelfRuntimeIDs.Add PPathRuntimeIDs.GetElementRuntimeID(eleCurrentUIElement)

  'Get a list of all ancestors to be excluded from the list of elements
  GetParents eleCurrentUIElement, Axis, Axis & "*/..", True, colAllAncestorsOrSelfRuntimeIDs ', False
  
  'Get all children of the root element which preceed/follow the current element
  'the current element RuntimeID is colAllAncestorsOrSelfRuntimeIDs(1)
  GetChildren this.ApplicationRootUIElement, Axis, "/", True, colAllAncestorsOrSelfRuntimeIDs, False

End Sub

Public Sub GetAttributes( _
  eleCurrentUIElement As UIAutomationClient.IUIAutomationElement, _
  InitialPPath As String)

  If VBA.Strings.Right(InitialPPath, 1) <> "/" Then
    InitialPPath = InitialPPath & "/"
  End If

  Dim key As Variant
  Dim propID As Long
  Dim propName As String
  Dim propValue As Variant
    
  For Each key In this.AutomationDictionaries.NavigablePropertyIDs.Keys
     
    propID = key
    propName = this.AutomationDictionaries.NavigablePropertyIDs(propID)
    If this.UnitTestingMode And ((propName = "ProcessId") Or (propName = "ProviderDescription") Or (propName = "BoundingRectangle") Or (propName = "NativeWindowHandle")) Then
      propValue = "#"
    Else
      propValue = eleCurrentUIElement.GetCurrentPropertyValue(propID)
      If IsArray(propValue) Then
        If UBound(propValue) = -1 Then
          propValue = ""
        Else
          If Not isEmpty(propValue) Then
            Dim i As Integer
            Dim strArrayOfValues As String
            strArrayOfValues = ""
            For i = 0 To UBound(propValue)
              If strArrayOfValues <> "" Then
                strArrayOfValues = strArrayOfValues & ", "
              End If
              strArrayOfValues = strArrayOfValues & propValue(i)
              Next i
            propValue = "{" & strArrayOfValues & "}"
          End If
        End If
      End If
    End If
     
    If Not isEmpty(propValue) And propValue <> "" Then
      Dim strPropertyString As String
      If propValue = "True" Or propValue = "False" Or VBA.Information.IsNumeric(propValue) Then
        strPropertyString = propValue
      Else
        strPropertyString = "'" & propValue & "'"
      End If
       
      this.PPathReturnClass.AddMatchingElement _
        this.PPathReturnClass.GetCandidateElements(this.CurrentLocationPathExpressionCounter), _
        eleCurrentUIElement, _
        propName, _
        InitialPPath & "@" & propName & "=" & strPropertyString
    End If
  
  Next key

End Sub
