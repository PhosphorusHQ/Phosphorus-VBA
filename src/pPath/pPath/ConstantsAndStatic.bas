Attribute VB_Name = "ConstantsAndStatic"
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

Public Const NULL_PPATH_ERROR_MESSAGE = "Some PPath expression must be specified!"

Public Const MISSING_RELATIVE_PPATH_CONTEXT_NODE = "PPaths is relative but no context node is provided '.'!"
Public Const UNUSED_RELATIVE_PPATH_CONTEXT_NODE = "A context node is provided but isn't used!"
Public Const MISSING_CONTEXT_NODE_INITIAL_PPATH = "A context node is provided but no initial PPath!"
Public Const NUMBER_OF_RELATIVE_PPATHS_TO_CONTEXT_NODES_MISMATCH = "The number of nodes and relative PPaths do not match!"

Public Const MISMATCHING_PARENTHESES_ERROR_MESSAGE = "Unequal number of left & right parentheses '(' & ')' in PPath!"
Public Const MISMATCHING_SQUARE_BARCKETS_ERROR_MESSAGE = "Unequal number of left & right square brackets '[' & ']' in PPath!"
Public Const INVALID_PROPERTY_ERROR_MESSAGE = "Unrecognised property in PPath!"

Public Const NO_NODETEST_PPATH_ERROR_MESSAGE = "Missing Node Test!"
Public Const INVALID_NODETEST_PPATH_ERROR_MESSAGE = "Invalid Node Test"

Public Const ILLEGAL_START_OF_PREDICATE_ERROR_MESSAGE = "Illegal start Of predicate!"
Public Const INVALID_PREDICATE = "Invalid Predicate"

Public Axes As AxisType

'https://www.w3schools.com/xml/xpath_axes.asp
'TODO: Add attribute axis later
Public Type AxisType
  None As String
  Child As String
  ChildShorthand As String
  Descendant As String
  DescendantShorthand As String
  DescendantOrSelf  As String
  Self As String
  SelfShorthand As String
  Parent As String
  ParentShorthand As String
  Ancestor As String
  AncestorOrSelf As String
  PrecedingSibling As String
  FollowingSibling As String
  Preceding As String
  Following As String
  Attribute As String
End Type

Public Enum PrincipalNodeKindType
  Attributes
  Elements
End Enum

Public Sub InitialiseAxesTypes()
  Axes.None = ""
  Axes.Child = "child::"
  Axes.ChildShorthand = "/"
  Axes.Descendant = "descendant::"
  Axes.DescendantShorthand = "//"
  Axes.DescendantOrSelf = "descendant-or-self::"
  Axes.Self = "self::"
  Axes.SelfShorthand = "."
  Axes.Parent = "parent::"
  Axes.ParentShorthand = ".."
  Axes.Ancestor = "ancestor::"
  Axes.AncestorOrSelf = "ancestor-or-self::"
  Axes.PrecedingSibling = "preceding-sibling::"
  Axes.FollowingSibling = "following-sibling::"
  Axes.Preceding = "preceding::"
  Axes.Following = "following::"
  Axes.Attribute = "attribute::"
End Sub

Public Function GetNewPhosphorusPPath() As pPath.Core
  Set GetNewPhosphorusPPath = New pPath.Core
End Function

Public Function GetNewPhosphorusPPathReturnClass(intNumberOfPPathExpressions As Integer) As pPath.ReturnClass
  Set GetNewPhosphorusPPathReturnClass = New pPath.ReturnClass
  GetNewPhosphorusPPathReturnClass.Initialise intNumberOfPPathExpressions
End Function

