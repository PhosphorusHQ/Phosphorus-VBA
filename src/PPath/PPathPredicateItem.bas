VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPathPredicateItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder PPath
Option Explicit
Option Base 1

Private Type PredicateItem
  SourcePPath  As String
  IsAPPath As Boolean
End Type

Private this As PredicateItem

Public Sub Initialise(SourcePPath As String)
  this.SourcePPath = SourcePPath
'TODO Determine if is a PPAth
  this.IsAPPath = False
End Sub

Public Function SourcePPath() As String
'TODO: Split by AND and OR for multiple items & ignore ('s )'s unless they are "()" (empty)
  SourcePPath = this.SourcePPath
End Function


