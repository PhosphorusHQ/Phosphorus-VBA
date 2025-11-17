VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredicateSet"
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
Option Base 1

Private Type PredicateSet
  SourcePPath  As String
  PredicateGroups() As pPath.PredicateGroup
End Type

Private This As PredicateSet

Public Sub Initialise(SourcePPath As String)
  This.SourcePPath = SourcePPath
'TODO: only expecting one predicate group for now!

  Dim strNextPredicateGroup As String
  Dim intPredicateGroupCounter As Integer
  Dim intNestingLevel As Integer
  Dim strCurrentCharacter As String
  Dim iCharacterCounter As Integer
  Dim intLengthOfSourcePPath As Integer
  intLengthOfSourcePPath = VBA.Strings.Len(This.SourcePPath)
  intNestingLevel = 0
  intPredicateGroupCounter = 0
  For iCharacterCounter = 1 To intLengthOfSourcePPath
    strCurrentCharacter = VBA.Strings.Mid(This.SourcePPath, iCharacterCounter, 1)
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
          ReDim This.PredicateGroups(intPredicateGroupCounter)
        Else
          ReDim Preserve This.PredicateGroups(intPredicateGroupCounter)
        End If
        Set This.PredicateGroups(intPredicateGroupCounter) = New pPath.PredicateGroup
        This.PredicateGroups(intPredicateGroupCounter).Initialise strNextPredicateGroup
        strNextPredicateGroup = ""
      End If
    ElseIf intNestingLevel = 1 Then
      strNextPredicateGroup = strNextPredicateGroup & strCurrentCharacter
    End If
  Next iCharacterCounter
 
End Sub

Public Function SourcePPath() As String
  SourcePPath = This.SourcePPath
End Function

Property Get PredicateGroup(Number As Integer) As pPath.PredicateGroup
  Set PredicateGroup = This.PredicateGroups(Number)
End Property

Public Function NumberOfPredicateGroups() As Integer
  NumberOfPredicateGroups = Phosphorus.Utils.GetSizeOfArray(This.PredicateGroups)
End Function

