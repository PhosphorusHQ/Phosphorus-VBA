VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPathPredicateGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder PPath
Option Explicit
Option Base 1

Private Type PredicateGroup
  SourcePPath  As String
  FinalPPath  As String
  IsPositionalPredicate As Boolean
'  PredicateItems() As PPathPredicateItem
End Type

Private this As PredicateGroup

Public Sub Initialise(SourcePPath As String)
  If SourcePPath = "" Then
    this.SourcePPath = SourcePPath
  Else
    this.SourcePPath = VBA.Strings.Mid(SourcePPath, 2, VBA.Strings.Len(SourcePPath) - 2)
  End If
  
'PJG 14/05/25 Items are not used!?
'  ReDim this.PredicateItems(1)
'  Set this.PredicateItems(1) = New PPathPredicateItem
'  this.PredicateItems(1).Initialise this.SourcePPath
  InitialiseFinalPPath
End Sub

Private Sub InitialiseFinalPPath()
  
  Dim boolCurrentPredicateIsProcessed As Boolean
  boolCurrentPredicateIsProcessed = False

  'Is the predicate just a position?
  If VBA.Information.IsNumeric(this.SourcePPath) Then
    Dim dblCurrentPredicate  As Double
    dblCurrentPredicate = VBA.Conversion.CDbl(this.SourcePPath)
    'Is it just an integer?
    If dblCurrentPredicate Mod 1 = 0 Then
      this.FinalPPath = "position()=" & this.SourcePPath
      boolCurrentPredicateIsProcessed = True
    End If
  Else
    If (VBA.Strings.InStr(1, this.SourcePPath, "first()") = 1) Or (VBA.Strings.InStr(1, this.SourcePPath, "last()") = 1) Then
      this.FinalPPath = "position()=" & this.SourcePPath
      boolCurrentPredicateIsProcessed = True
    End If
  End If
  
  If Not boolCurrentPredicateIsProcessed Then
    this.FinalPPath = this.SourcePPath
  End If
  
  'Is this a Positional Predicate?
  this.IsPositionalPredicate = (VBA.Strings.InStr(1, this.FinalPPath, "position()") > 0)
  
End Sub

Public Function SourcePPath() As String
  SourcePPath = this.SourcePPath
End Function

'Public Function NumberOfPredicateItems() As Integer
'  NumberOfPredicateItems = Utils.GetSizeOfArray(this.PredicateItems)
'End Function

Public Function FinalPPath() As String
  FinalPPath = this.FinalPPath
End Function

Public Function IsPositionalPredicate() As Boolean
  IsPositionalPredicate = this.IsPositionalPredicate
End Function

'Property Get PredicateItem(Number As Integer) As PPathPredicateItem
'  Set PredicateItem = this.PredicateItems(Number)
'End Property
