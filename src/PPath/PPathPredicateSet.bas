VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPathPredicateSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder PPath
Option Explicit
Option Base 1

Private Type PredicateSet
  SourcePPath  As String
  PredicateGroups() As PPathPredicateGroup
End Type

Private this As PredicateSet

Public Sub Initialise(SourcePPath As String)
  this.SourcePPath = SourcePPath
'TODO: only expecting one predicate group for now!

  Dim strNextPredicateGroup As String
  Dim intPredicateGroupCounter As Integer
  Dim intNestingLevel As Integer
  Dim strCurrentCharacter As String
  Dim iCharacterCounter As Integer
  Dim intLengthOfSourcePPath As Integer
  intLengthOfSourcePPath = VBA.Strings.Len(this.SourcePPath)
  intNestingLevel = 0
  intPredicateGroupCounter = 0
  For iCharacterCounter = 1 To intLengthOfSourcePPath
    strCurrentCharacter = VBA.Strings.Mid(this.SourcePPath, iCharacterCounter, 1)
    If strCurrentCharacter = "[" Then
      intNestingLevel = intNestingLevel + 1
      If intNestingLevel = 1 Then
        strNextPredicateGroup = strCurrentCharacter
      End If
    ElseIf strCurrentCharacter = "]" Then
      intNestingLevel = intNestingLevel - 1
      If intNestingLevel = 0 Then
        strNextPredicateGroup = strNextPredicateGroup & strCurrentCharacter
        intPredicateGroupCounter = intPredicateGroupCounter + 1
        If intPredicateGroupCounter = 1 Then
          ReDim this.PredicateGroups(intPredicateGroupCounter)
        Else
          ReDim Preserve this.PredicateGroups(intPredicateGroupCounter)
        End If
        Set this.PredicateGroups(intPredicateGroupCounter) = New PPathPredicateGroup
        this.PredicateGroups(intPredicateGroupCounter).Initialise strNextPredicateGroup
        strNextPredicateGroup = ""
      End If
    ElseIf intNestingLevel = 1 Then
      strNextPredicateGroup = strNextPredicateGroup & strCurrentCharacter
    End If
  Next iCharacterCounter
 
End Sub

Public Function SourcePPath() As String
  SourcePPath = this.SourcePPath
End Function

Property Get PredicateGroup(Number As Integer) As PPathPredicateGroup
  Set PredicateGroup = this.PredicateGroups(Number)
End Property

Public Function NumberOfPredicateGroups() As Integer
  NumberOfPredicateGroups = Utils.GetSizeOfArray(this.PredicateGroups)
End Function

