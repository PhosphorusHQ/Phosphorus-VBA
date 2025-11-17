VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredicateGroup"
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

Private Type PredicateGroup
  SourcePPath  As String
  FinalPPath  As String
  IsPositionalPredicate As Boolean
'  PredicateItems() As PPathPredicateItem
End Type

Private This As PredicateGroup

Public Sub Initialise(SourcePPath As String)
  If SourcePPath = "" Then
    This.SourcePPath = SourcePPath
  Else
    This.SourcePPath = VBA.Strings.Mid(SourcePPath, 2, VBA.Strings.Len(SourcePPath) - 2)
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
  If VBA.Information.IsNumeric(This.SourcePPath) Then
    Dim dblCurrentPredicate  As Double
    dblCurrentPredicate = VBA.Conversion.CDbl(This.SourcePPath)
    'Is it just an integer?
    If dblCurrentPredicate Mod 1 = 0 Then
      This.FinalPPath = "position()=" & This.SourcePPath
      boolCurrentPredicateIsProcessed = True
    End If
  Else
    If (VBA.Strings.InStr(1, This.SourcePPath, "first()") = 1) Or (VBA.Strings.InStr(1, This.SourcePPath, "last()") = 1) Then
      This.FinalPPath = "position()=" & This.SourcePPath
      boolCurrentPredicateIsProcessed = True
    End If
  End If
  
  If Not boolCurrentPredicateIsProcessed Then
    This.FinalPPath = This.SourcePPath
  End If
  
  'Is this a Positional Predicate?
  This.IsPositionalPredicate = (VBA.Strings.InStr(1, This.FinalPPath, "position()") > 0)
  
End Sub

Public Function SourcePPath() As String
  SourcePPath = This.SourcePPath
End Function

'Public Function NumberOfPredicateItems() As Integer
'  NumberOfPredicateItems = Utils.GetSizeOfArray(this.PredicateItems)
'End Function

Public Function FinalPPath() As String
  FinalPPath = This.FinalPPath
End Function

Public Function IsPositionalPredicate() As Boolean
  IsPositionalPredicate = This.IsPositionalPredicate
End Function

'Property Get PredicateItem(Number As Integer) As PPathPredicateItem
'  Set PredicateItem = this.PredicateItems(Number)
'End Property


