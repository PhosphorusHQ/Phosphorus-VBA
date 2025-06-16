VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPathStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder PPath
Option Explicit
Option Base 1

Private Type Step
  Axis As String
  PrincipalNodeKind As PrincipalNodeKindType
  NodeTest As String
  NodeTestKind As String
  PredicateSets() As PPathPredicateSet
  RemainingPPath As String
End Type

Private this As Step

Property Get Axis() As String
  Axis = this.Axis
End Property

Property Let Axis(strAxis As String)
  this.Axis = strAxis
  If strAxis = Axes.Attribute Then
    this.PrincipalNodeKind = PrincipalNodeKindType.Attributes
  Else
    this.PrincipalNodeKind = PrincipalNodeKindType.Elements
  End If
End Property

Property Get NodeTest() As String
  NodeTest = this.NodeTest
End Property

Property Let NodeTest(strNodeTest As String)
  this.NodeTest = strNodeTest
End Property

Property Get NodeTestKind() As String
  NodeTestKind = this.NodeTestKind
End Property

Property Let NodeTestKind(strNodeTestKind As String)
  this.NodeTestKind = strNodeTestKind
End Property

Property Get PrincipalNodeKind() As PrincipalNodeKindType
  PrincipalNodeKind = this.PrincipalNodeKind
End Property

Property Let PrincipalNodeKind(PrincipalNodeKind As PrincipalNodeKindType)
  this.PrincipalNodeKind = PrincipalNodeKind
End Property

Public Sub AddPredicates(PredicatePPath As String)
  If VBA.Strings.InStr(1, PredicatePPath, ")") = 0 Then
'DO we need this ????
    'There is only 1 set of predicates
    ReDim this.PredicateSets(1)
    Set this.PredicateSets(1) = New PPathPredicateSet
    this.PredicateSets(1).Initialise PredicatePPath
  Else
    Dim intNumberOfPredicateSets As Integer
    intNumberOfPredicateSets = 0
    
  '  strPredicatesArray = VBA.Strings.Split(PredicatePPath, ")")
'MsgBox "PJG1"
'TODO: Split arr by ")"'s that are not inside a [] pair
    Dim intLengthOfPredicatePPath As Integer
    intLengthOfPredicatePPath = VBA.Strings.Len(PredicatePPath)
    
    Dim intPredicateSetCounter As Integer
    intPredicateSetCounter = 0
    
    Dim intStartOfCurrentPredicateSet As Integer
    intStartOfCurrentPredicateSet = 1
    
    Dim intCharacterCounter As Integer
    For intCharacterCounter = 1 To intLengthOfPredicatePPath
      
      Dim boolInsideAPredicate As Boolean
      Dim strCurrentPredicateSetPPath As String
      Dim strCurrentCharacter As String
      strCurrentCharacter = VBA.Strings.Mid(PredicatePPath, intCharacterCounter, 1)
      
      If strCurrentCharacter = "[" Then
        boolInsideAPredicate = True
      Else
        Dim boolAddNewSet As Boolean
        boolAddNewSet = False
        If strCurrentCharacter = "]" Then
          boolInsideAPredicate = False
          'Is this the end of the PPath?
          If intCharacterCounter = intLengthOfPredicatePPath Then
            'Add the remaing attibutes as a set
            intPredicateSetCounter = intPredicateSetCounter + 1
            strCurrentPredicateSetPPath = VBA.Strings.Mid(PredicatePPath, intStartOfCurrentPredicateSet, intLengthOfPredicatePPath - intStartOfCurrentPredicateSet + 1)
            boolAddNewSet = True
          End If
        Else
          If Not boolInsideAPredicate Then
            If strCurrentCharacter = ")" Then
              intPredicateSetCounter = intPredicateSetCounter + 1
              If intCharacterCounter = 1 Then
                strCurrentPredicateSetPPath = ""
              Else
                strCurrentPredicateSetPPath = VBA.Strings.Mid(PredicatePPath, intStartOfCurrentPredicateSet, intCharacterCounter - intStartOfCurrentPredicateSet)
              End If
              boolAddNewSet = True
            End If
          End If
        End If
        If boolAddNewSet Then
          If intPredicateSetCounter = 1 Then
            ReDim this.PredicateSets(intPredicateSetCounter)
          Else
            ReDim Preserve this.PredicateSets(intPredicateSetCounter)
          End If
          Set this.PredicateSets(intPredicateSetCounter) = New PPathPredicateSet
          this.PredicateSets(intPredicateSetCounter).Initialise strCurrentPredicateSetPPath
          intStartOfCurrentPredicateSet = intCharacterCounter + 1
        End If
      End If
      
    Next intCharacterCounter
    
    
'    intNumberOfPredicateSets = UBound(strPredicatesArray) + 1
'    ReDim this.PredicateSets(intNumberOfPredicateSets)
'    Dim intPredicateSetCounter As Integer
'    For intPredicateSetCounter = 1 To intNumberOfPredicateSets
'      Set this.PredicateSets(intPredicateSetCounter) = New PPathPredicateSet
'      this.PredicateSets(intPredicateSetCounter).Initialise strPredicatesArray(intPredicateSetCounter - 1)
'    Next intPredicateSetCounter

'MsgBox "PJG!"
  End If
  
'  Dim intCharacterCounter
End Sub

Property Get PredicateSet(Number As Integer) As PPathPredicateSet
  Set PredicateSet = this.PredicateSets(Number)
End Property

Public Function NumberOfPredicateSets() As Integer
  NumberOfPredicateSets = Utils.GetSizeOfArray(this.PredicateSets)
End Function

Property Get RemainingPPath() As String
  RemainingPPath = this.RemainingPPath
End Property

Property Let RemainingPPath(strRemainingPPath As String)
  this.RemainingPPath = strRemainingPPath
End Property

