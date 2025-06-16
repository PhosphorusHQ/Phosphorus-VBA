VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPathReturnClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder PPath
Option Explicit
Option Base 1

'For final elements returned
Private MatchingElements As PPathMatchingElements
Private lintNumberOfXPathExpressions As Integer
' Array of matching elements class 1 to {NumberOfXPatExpressions} for processing each reducing set of candidate elements for each PPath Expression
Private WorkingCopyOfCandidateElements() As PPathMatchingElements
Private CandidateElements() As PPathMatchingElements
Private TempCandidateElements() As PPathMatchingElements

Private ErrorMessage As String

Private Type Returns
  TopLevelFunctionPrefix As String
  TopLevelFunctionSuffix As String
  ReturnedValue As Variant
End Type

Private this As Returns

Public Sub Initialise(intNumberOfXPathExpressions As Integer)
  lintNumberOfXPathExpressions = intNumberOfXPathExpressions
  Set MatchingElements = New PPathMatchingElements
  ReDim WorkingCopyOfCandidateElements(lintNumberOfXPathExpressions)
  ReDim CandidateElements(lintNumberOfXPathExpressions)
  ReDim TempCandidateElements(lintNumberOfXPathExpressions)
  Dim i As Integer
  For i = 1 To lintNumberOfXPathExpressions
    Set WorkingCopyOfCandidateElements(i) = New PPathMatchingElements
    Set CandidateElements(i) = New PPathMatchingElements
    Set TempCandidateElements(i) = New PPathMatchingElements
  Next i
End Sub

Public Sub AddMatchingElement( _
  ByRef TargetMatchingElements As PPathMatchingElements, _
  ByRef MatchingElement As UIAutomationClient.IUIAutomationElement, _
  AttributeName As String, _
  PPath As String)
  
  Dim boolElementAddedAlready As Boolean
  boolElementAddedAlready = False
      
  Dim strNewUIElementRuntimeID As String
  strNewUIElementRuntimeID = PPathRuntimeIDs.GetElementRuntimeID(MatchingElement)
 
'If strNewUIElementRuntimeID = "" Then
'  Debug.Print MatchingElement.GetCachedPropertyValue(UIAutomationClient.UIA_PropertyIds.UIA_NativeWindowHandlePropertyId)
'End If

  If TargetMatchingElements.GetNumberOfMatchingElements > 0 Then
    
    'Ignore transient elements with no runtime ID - these could be MSAA elements!?
    If strNewUIElementRuntimeID <> "" Then
      Dim i As Integer
      For i = 1 To TargetMatchingElements.GetNumberOfMatchingElements
        Dim strExistingUIElementRuntimeID As String
        Dim strExistingUIElementAttributeName As String
        strExistingUIElementRuntimeID = PPathRuntimeIDs.GetElementRuntimeID(TargetMatchingElements.GetMatchingElement(i))
        strExistingUIElementAttributeName = TargetMatchingElements.GetAttributeName(i)
        If strExistingUIElementRuntimeID = "" Then
          MsgBox "No Unique Property ID for existing element"
        ElseIf (strNewUIElementRuntimeID = strExistingUIElementRuntimeID) And (AttributeName = strExistingUIElementAttributeName) Then
          boolElementAddedAlready = True
          Exit For
        End If
      Next i
    End If
  End If

  'Do not add elements that have been added already or transient elements that have no runtime ID
  If (Not boolElementAddedAlready) And (strNewUIElementRuntimeID <> "") Then
    TargetMatchingElements.AddMatchingElement TargetMatchingElements.GetNumberOfMatchingElements + 1, MatchingElement, AttributeName, PPath
  End If
  
End Sub

Property Get GetErrorMessage() As String
  GetErrorMessage = ErrorMessage
End Property

Property Let SetErrorMessage(strErrorMessage As String)
  ErrorMessage = strErrorMessage
End Property

'MatchingElements

Public Sub RemoveAllMatchingElements()
  MatchingElements.EraseAllElements
  Dim i As Integer
  For i = 1 To lintNumberOfXPathExpressions
   CandidateElements(i).EraseAllElements
  Next i
End Sub

Public Function GetNumberOfMatchingElements(ByRef TargetMatchingElements As PPathMatchingElements) As Long
  GetNumberOfMatchingElements = TargetMatchingElements.GetNumberOfMatchingElements
End Function

Public Function GetFinalNumberOfMatchingElements() As Long
  GetFinalNumberOfMatchingElements = MatchingElements.GetNumberOfMatchingElements
End Function

Property Get GetMatchingElement(ElementNumber As Integer) As UIAutomationClient.IUIAutomationElement
  Set GetMatchingElement = MatchingElements.GetMatchingElement(ElementNumber)
End Property

Property Get GetMatchingNavigationalPPaths() As String()
  GetMatchingNavigationalPPaths = MatchingElements.GetNavigationalPPaths
End Property

Property Get GetMatchingNavigationalPPath(ElementNumber As Integer) As String
  GetMatchingNavigationalPPath = MatchingElements.GetNavigationalPPath(ElementNumber)
End Property


'CandidateElements

Property Get GetCandidateElements(LocationPathExpressionCounter As Integer) As PPathMatchingElements
  Set GetCandidateElements = CandidateElements(LocationPathExpressionCounter)
End Property

Public Sub MoveCandidateElementsToWorkingCopy(intXPathExpressionNumber As Integer)
'  Dim NewMatchingElements() As UIAutomationClient.IUIAutomationElement
'  CandidateElements(intXPathExpressionNumber).GetMatchingElementsArray
'  NewMatchingElements = CandidateElements(intXPathExpressionNumber).GetMatchingElementsArray
  WorkingCopyOfCandidateElements(intXPathExpressionNumber).SetMatchingElements CandidateElements(intXPathExpressionNumber).GetMatchingElementsArray
  WorkingCopyOfCandidateElements(intXPathExpressionNumber).SetAttributeNames CandidateElements(intXPathExpressionNumber).GetAttributeNames
  WorkingCopyOfCandidateElements(intXPathExpressionNumber).SetNavigationalPPaths CandidateElements(intXPathExpressionNumber).GetNavigationalPPaths
  CandidateElements(intXPathExpressionNumber).ClearDownAllArrays
End Sub

Public Sub PromoteCandidateElementsToMatchingElements() 'intXPathExpressionNumber As Integer)
'TODO: Implement AND/OR logic for multiple PPath expressions
  Dim intXPathExpressionCounter As Integer
  Dim intCandidateElementCounter As Integer
  Dim intNumberOfCandidateElements As Integer
  For intXPathExpressionCounter = 1 To lintNumberOfXPathExpressions
    intNumberOfCandidateElements = CandidateElements(intXPathExpressionCounter).GetNumberOfMatchingElements
    For intCandidateElementCounter = 1 To intNumberOfCandidateElements
      AddMatchingElement _
        MatchingElements, _
        CandidateElements(intXPathExpressionCounter).GetMatchingElement(intCandidateElementCounter), _
        CandidateElements(intXPathExpressionCounter).GetAttributeName(intCandidateElementCounter), _
        CandidateElements(intXPathExpressionCounter).GetNavigationalPPath(intCandidateElementCounter)
    Next intCandidateElementCounter
  Next intXPathExpressionCounter
End Sub

'TempCandidateElements

Public Function GetNumberOfTempCandidateElements(LocationPathExpressionCounter As Integer)
  GetNumberOfTempCandidateElements = TempCandidateElements(LocationPathExpressionCounter).GetNumberOfMatchingElements
End Function

Property Get GetTempCandidateElements(LocationPathExpressionCounter As Integer) As PPathMatchingElements
  Set GetTempCandidateElements = TempCandidateElements(LocationPathExpressionCounter)
End Property

Public Sub PromoteTempCandidateElementsToCandidateElementsInReverseOrder(LocationPathExpressionCounter As Integer)
  Dim intNumberOfTempCandidateElements As Integer
  Dim intTempCandidateElementCounter As Integer
  intNumberOfTempCandidateElements = TempCandidateElements(LocationPathExpressionCounter).GetNumberOfMatchingElements
  For intTempCandidateElementCounter = intNumberOfTempCandidateElements To 1 Step -1
    AddMatchingElement _
      CandidateElements(LocationPathExpressionCounter), _
      TempCandidateElements(LocationPathExpressionCounter).GetMatchingElement(intTempCandidateElementCounter), _
      TempCandidateElements(LocationPathExpressionCounter).GetAttributeName(intTempCandidateElementCounter), _
      TempCandidateElements(LocationPathExpressionCounter).GetNavigationalPPath(intTempCandidateElementCounter)
  Next intTempCandidateElementCounter
  TempCandidateElements(LocationPathExpressionCounter).EraseAllElements
End Sub

'WorkingCopyOfCandidateElements

Public Function GetNumberOfWorkingCopyOfCandidateElements(LocationPathExpressionCounter As Integer)
  GetNumberOfWorkingCopyOfCandidateElements = WorkingCopyOfCandidateElements(LocationPathExpressionCounter).GetNumberOfMatchingElements
End Function

Property Get GetWorkingCopyOfCandidateElements(LocationPathExpressionCounter As Integer) As PPathMatchingElements
  Set GetWorkingCopyOfCandidateElements = WorkingCopyOfCandidateElements(LocationPathExpressionCounter)
End Property

Public Function GetWorkingCopyElement(LocationPathExpressionCounter As Integer, ElementNumber As Integer) As UIAutomationClient.IUIAutomationElement
  Set GetWorkingCopyElement = WorkingCopyOfCandidateElements(LocationPathExpressionCounter).GetMatchingElement(ElementNumber)
End Function

Public Function GetWorkingCopyNavigationalPPath(LocationPathExpressionCounter As Integer, ElementNumber As Integer) As String
  GetWorkingCopyNavigationalPPath = WorkingCopyOfCandidateElements(LocationPathExpressionCounter).GetNavigationalPPath(ElementNumber)
End Function

Public Function GetWorkingCopyAttributeName(LocationPathExpressionCounter As Integer, ElementNumber As Integer) As String
  GetWorkingCopyAttributeName = WorkingCopyOfCandidateElements(LocationPathExpressionCounter).GetAttributeName(ElementNumber)
End Function

Property Let TopLevelFunctionPrefix(FunctionPrefix As String)
  this.TopLevelFunctionPrefix = FunctionPrefix
End Property

Property Get TopLevelFunctionPrefix() As String
  TopLevelFunctionPrefix = this.TopLevelFunctionPrefix
End Property

Property Let TopLevelFunctionSuffix(FunctionSuffix As String)
  this.TopLevelFunctionSuffix = FunctionSuffix
End Property

Property Get TopLevelFunctionSuffix() As String
  TopLevelFunctionSuffix = this.TopLevelFunctionSuffix
End Property


Property Let ReturnedValue(Value As Variant)
  this.ReturnedValue = Value
End Property

Property Get ReturnedValue() As Variant
  ReturnedValue = this.ReturnedValue
End Property

