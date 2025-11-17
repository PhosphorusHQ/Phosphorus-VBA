VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Axes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private This As pPath.Common

Public Sub Initialise(ByRef sharedthis As pPath.Common)
  Set This = sharedthis
End Sub

Public Sub ProcessNextAxis( _
  myNextStep As pPath.Step)
      
  'The method loops through each element in the current Working Copy and call the axis method(s) relevant to each axis
  Dim intElementCounter As Integer
  Dim eleCurrentUIElement As UIAutomationClient.IUIAutomationElement
  Dim strInitialPPath As String
  Dim intNumberOfWorkingCopyElements As Integer
  intNumberOfWorkingCopyElements = This.PPathReturnClass.GetNumberOfWorkingCopyOfCandidateElements(This.CurrentLocationPathExpressionCounter)
  For intElementCounter = 1 To intNumberOfWorkingCopyElements

    Set eleCurrentUIElement = This.PPathReturnClass.GetWorkingCopyElement(This.CurrentLocationPathExpressionCounter, intElementCounter)
    strInitialPPath = This.PPathReturnClass.GetWorkingCopyNavigationalPPath(This.CurrentLocationPathExpressionCounter, intElementCounter)

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
  pPath.Utils.OutputElementDetails eleCurrentUIElement, strCurrentElementAbsoluteXPath, This.DebugMode, This.AutomationDictionaries
  This.PPathReturnClass.AddMatchingElement This.PPathReturnClass.GetCandidateElements(This.CurrentLocationPathExpressionCounter), eleCurrentUIElement, "", strCurrentElementAbsoluteXPath
End Sub

Private Sub GetParents( _
  eleCurrentUIElement As UIAutomationClient.IUIAutomationElement, _
  strCurrentStepAxis As String, _
  strInitialPPath As String, _
  boolGetAllAncestors As Boolean, _
  Optional ByRef AllAncestorsOrSelfRuntimeIDs As Collection)
  
  'The parent:: axis is a reverse axis, but as we only get one node it doesn't make any difference which order it's in!?
  'Also, this method already returns ancestor nodes in reverse document order, so there is no need to reverse.
  
  'Never go above the current application root element!
  Dim boolCurrentIsApplicationRootUIElement As Boolean
  boolCurrentIsApplicationRootUIElement = False
  Dim strCurrentUIElementRunTimeID As String
  Dim strApplicationRootUIElementRunTimeID As String
  strCurrentUIElementRunTimeID = pPath.RuntimeIDs.GetElementRuntimeID(eleCurrentUIElement)
  strApplicationRootUIElementRunTimeID = pPath.RuntimeIDs.GetElementRuntimeID(This.ApplicationRootUIElement)
  If strCurrentUIElementRunTimeID = strApplicationRootUIElementRunTimeID Then
    boolCurrentIsApplicationRootUIElement = True
  End If

'PJG Check this fix:
  If Not boolCurrentIsApplicationRootUIElement Then
  
    Dim eleParentUIElement As UIAutomationClient.IUIAutomationElement
    Set eleParentUIElement = This.TreeWalker.GetParentElement(eleCurrentUIElement)
    
    Dim strParentUIElementCurrentUIElementRuntimeID As String
    strParentUIElementCurrentUIElementRuntimeID = pPath.RuntimeIDs.GetElementRuntimeID(eleParentUIElement)
          
    Dim strCurrentControlType As String
    Dim strCurrentElementAbsoluteXPath As String
  
    Dim boolAddElement As Boolean
    boolAddElement = False
    
    strCurrentControlType = This.AutomationDictionaries.ControlTypeIDs(eleParentUIElement.CurrentControlType)
    strCurrentElementAbsoluteXPath = strInitialPPath & "/" & strCurrentControlType & "[@Name='" & eleParentUIElement.CurrentName & "']"
    
    If strCurrentStepAxis = Axes.ParentShorthand Then
      boolAddElement = True
    End If
      
'PJG Check this
'    If boolAddElement And Not boolCurrentIsApplicationRootUIElement Then
    If boolAddElement Then
      OutputElementDetails eleParentUIElement, strCurrentElementAbsoluteXPath, This.DebugMode, This.AutomationDictionaries
      This.PPathReturnClass.AddMatchingElement This.PPathReturnClass.GetCandidateElements(This.CurrentLocationPathExpressionCounter), eleParentUIElement, "", strCurrentElementAbsoluteXPath
    End If
    
    If Not boolCurrentIsApplicationRootUIElement Then
      GetParents eleParentUIElement, strCurrentStepAxis, strCurrentElementAbsoluteXPath & "/..", boolGetAllAncestors, AllAncestorsOrSelfRuntimeIDs
    End If
  
'PJG Check this
  End If
       
End Sub

Private Sub GetSiblings( _
  eleCurrentUIElement As UIAutomationClient.IUIAutomationElement, _
  Axis As String, _
  strInitialPPath As String)
  
  Dim strCurrentElementRuntimeID As String
  strCurrentElementRuntimeID = pPath.RuntimeIDs.GetElementRuntimeID(eleCurrentUIElement)
  
  Dim eleParentUIElement As UIAutomationClient.IUIAutomationElement
  Set eleParentUIElement = This.TreeWalker.GetParentElement(eleCurrentUIElement)
  
  Dim eleChildrenUIElementArray As UIAutomationClient.IUIAutomationElementArray
  Set eleChildrenUIElementArray = eleParentUIElement.FindAll(UIAutomationClient.TreeScope.TreeScope_Children, This.UIAutomation.CreateTrueCondition())
  
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
    strCurrentChildElementRuntimeID = pPath.RuntimeIDs.GetElementRuntimeID(eleChildUIElement)
    strCurrentControlType = This.AutomationDictionaries.ControlTypeIDs(eleChildUIElement.CurrentControlType)
    If controls.Exists(strCurrentControlType) Then
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
      OutputElementDetails eleChildUIElement, strCurrentElementAbsoluteXPath, This.DebugMode, This.AutomationDictionaries
      This.PPathReturnClass.AddMatchingElement This.PPathReturnClass.GetTempCandidateElements(This.CurrentLocationPathExpressionCounter), eleChildUIElement, "", strCurrentElementAbsoluteXPath
    End If
      
    'Now promote preceding siblings to Candidate Elements as in reverse order
    If (boolIsCurrentElement And Axis = Axes.PrecedingSibling) Then
      This.PPathReturnClass.PromoteTempCandidateElementsToCandidateElementsInReverseOrder This.CurrentLocationPathExpressionCounter
    End If
    
    'Add following siblings straight to Candidate Elements as they are in forward order
    If (boolAfterCurrentElement And Axis = Axes.FollowingSibling And Not boolIsCurrentElement) Then
      OutputElementDetails eleChildUIElement, strCurrentElementAbsoluteXPath, This.DebugMode, This.AutomationDictionaries
      This.PPathReturnClass.AddMatchingElement This.PPathReturnClass.GetCandidateElements(This.CurrentLocationPathExpressionCounter), eleChildUIElement, "", strCurrentElementAbsoluteXPath
    End If
  
  Next intListItemCounter
  
End Sub

Private Sub GetChildren( _
  eleCurrentUIElement As UIAutomationClient.IUIAutomationElement, _
  Axis As String, _
  strInitialPPath As String, _
  boolGetGrandChildren As Boolean, _
  Optional ByRef AllAncestorsOrSelfRuntimeIDs As Collection, _
  Optional ByRef boolPrecedingOrFollowingRootElementFound As Boolean)
  
  Dim eleChildrenUIElementArray As UIAutomationClient.IUIAutomationElementArray
  Dim eleArrayOfChildrenUIElements() As UIAutomationClient.IUIAutomationElement
  Set eleChildrenUIElementArray = eleCurrentUIElement.FindAll(UIAutomationClient.TreeScope.TreeScope_Children, This.UIAutomation.CreateTrueCondition())
    
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
      intCurrentSizeOfArray = Phosphorus.Utils.GetSizeOfArray(eleArrayOfChildrenUIElements)
      intCurrentSizeOfArray = intCurrentSizeOfArray + 1
      If intCurrentSizeOfArray = 1 Then
        ReDim eleArrayOfChildrenUIElements(intCurrentSizeOfArray)
      Else
        ReDim Preserve eleArrayOfChildrenUIElements(intCurrentSizeOfArray)
      End If
      Set eleArrayOfChildrenUIElements(intCurrentSizeOfArray) = eleChildUIElement
      strCurrentControlType = This.AutomationDictionaries.ControlTypeIDs(eleChildUIElement.CurrentControlType)
      If controls.Exists(strCurrentControlType) Then
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
      
        strCurrentElementRuntimeID = pPath.RuntimeIDs.GetElementRuntimeID(eleChildUIElement)
        If strCurrentElementRuntimeID = AllAncestorsOrSelfRuntimeIDs(1) Then
          boolPrecedingOrFollowingRootElementFound = True
          'Now promote preceding elements to Candidate Elements as in reverse order
          If (Axis = Axes.Preceding) Then
            This.PPathReturnClass.PromoteTempCandidateElementsToCandidateElementsInReverseOrder This.CurrentLocationPathExpressionCounter
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
            OutputElementDetails eleChildUIElement, strCurrentElementAbsoluteXPath, This.DebugMode, This.AutomationDictionaries
            This.PPathReturnClass.AddMatchingElement This.PPathReturnClass.GetTempCandidateElements(This.CurrentLocationPathExpressionCounter), eleChildUIElement, "", strCurrentElementAbsoluteXPath
          End If
          
          If (Axis = Axes.Following) And boolPrecedingOrFollowingRootElementFound Then
            boolAddChild = True
          End If
          
        End If
        
      End If
      
      If boolAddChild Then
        OutputElementDetails eleChildUIElement, strCurrentElementAbsoluteXPath, This.DebugMode, This.AutomationDictionaries
        This.PPathReturnClass.AddMatchingElement This.PPathReturnClass.GetCandidateElements(This.CurrentLocationPathExpressionCounter), eleChildUIElement, "", strCurrentElementAbsoluteXPath
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
  Dim colAllAncestorsOrSelfRuntimeIDs As Collection
  Set colAllAncestorsOrSelfRuntimeIDs = New Collection
  
  'Add the current element to the list
  colAllAncestorsOrSelfRuntimeIDs.Add pPath.RuntimeIDs.GetElementRuntimeID(eleCurrentUIElement)

  'Get a list of all ancestors to be excluded from the list of elements
  GetParents eleCurrentUIElement, Axis, Axis & "*/..", True, colAllAncestorsOrSelfRuntimeIDs ', False
  
  'Get all children of the root element which preceed/follow the current element
  'the current element RuntimeID is colAllAncestorsOrSelfRuntimeIDs(1)
  GetChildren This.ApplicationRootUIElement, Axis, "/", True, colAllAncestorsOrSelfRuntimeIDs, False

End Sub

Public Sub GetAttributes( _
  eleCurrentUIElement As UIAutomationClient.IUIAutomationElement, _
  InitialPPath As String)

  If VBA.Strings.Right(InitialPPath, 1) <> "/" Then
    InitialPPath = InitialPPath & "/"
  End If

  Dim Key As Variant
  Dim propID As Long
  Dim propName As String
  Dim propValue As Variant
    
  For Each Key In This.AutomationDictionaries.NavigablePropertyIDs.Keys
     
    propID = Key
    propName = This.AutomationDictionaries.NavigablePropertyIDs(propID)
    If This.UnitTestingMode And ((propName = "ProcessId") Or (propName = "ProviderDescription") Or (propName = "BoundingRectangle") Or (propName = "NativeWindowHandle")) Then
      propValue = "#"
    Else
      propValue = eleCurrentUIElement.GetCurrentPropertyValue(propID)
      If IsArray(propValue) Then
        If UBound(propValue) = -1 Then
          propValue = ""
        Else
          If Not IsEmpty(propValue) Then
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
     
    If Not IsEmpty(propValue) And propValue <> "" Then
      Dim strPropertyString As String
      If propValue = "True" Or propValue = "False" Or VBA.Information.IsNumeric(propValue) Then
        strPropertyString = propValue
      Else
        strPropertyString = "'" & propValue & "'"
      End If
       
      This.PPathReturnClass.AddMatchingElement _
        This.PPathReturnClass.GetCandidateElements(This.CurrentLocationPathExpressionCounter), _
        eleCurrentUIElement, _
        propName, _
        InitialPPath & "@" & propName & "=" & strPropertyString
    End If
  
  Next Key

End Sub




