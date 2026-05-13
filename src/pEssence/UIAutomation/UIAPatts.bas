Attribute VB_Name = "UIAPatts"
'@Folder UIAutomation
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controlpattern-ids 10000-10050
Public Enum UIAPatterns 'UIA_PatternIds
  Annotation = UIA_PatternIds.UIA_AnnotationPatternId
  CustomNavigation = UIA_PatternIds.UIA_CustomNavigationPatternId
  Dock = UIA_PatternIds.UIA_DockPatternId
  Drag = UIA_PatternIds.UIA_DragPatternId
  DropTarget = UIA_PatternIds.UIA_DropTargetPatternId
  ExpandCollapse = UIA_PatternIds.UIA_ExpandCollapsePatternId
  GridItem = UIA_PatternIds.UIA_GridItemPatternId
  GridPattern = UIA_PatternIds.UIA_GridPatternId
  Invoke = UIA_PatternIds.UIA_InvokePatternId
  ItemContainer = UIA_PatternIds.UIA_ItemContainerPatternId
  LegacyIAccessible = UIA_PatternIds.UIA_LegacyIAccessiblePatternId
  MultipleView = UIA_PatternIds.UIA_MultipleViewPatternId
  ObjectModel = UIA_PatternIds.UIA_ObjectModelPatternId
  RangeValue = UIA_PatternIds.UIA_RangeValuePatternId
  ScrollItemPattern = UIA_PatternIds.UIA_ScrollItemPatternId
  Scroll = UIA_PatternIds.UIA_ScrollPatternId
  SelectionItem = UIA_PatternIds.UIA_SelectionItemPatternId
  Selection = UIA_PatternIds.UIA_SelectionPatternId
  Selection2 = UIA_PatternIds.UIA_SelectionPattern2Id
  SpreadsheetItem = UIA_PatternIds.UIA_SpreadsheetItemPatternId
  Spreadsheet = UIA_PatternIds.UIA_SpreadsheetPatternId
  Styles = UIA_PatternIds.UIA_StylesPatternId
  SynchronizedInput = UIA_PatternIds.UIA_SynchronizedInputPatternId
  TableItem = UIA_PatternIds.UIA_TableItemPatternId
  Table = UIA_PatternIds.UIA_TablePatternId
  TextChild = UIA_PatternIds.UIA_TextChildPatternId
  TextEdit = UIA_PatternIds.UIA_TextEditPatternId
  Text = UIA_PatternIds.UIA_TextPatternId
  Text2 = UIA_PatternIds.UIA_TextPattern2Id
  Toggle = UIA_PatternIds.UIA_TogglePatternId
  Transform = UIA_PatternIds.UIA_TransformPatternId
  Transform2 = UIA_PatternIds.UIA_TransformPattern2Id
  Value = UIA_PatternIds.UIA_ValuePatternId
  VirtualizedItem = UIA_PatternIds.UIA_VirtualizedItemPatternId
  Window = UIA_PatternIds.UIA_WindowPatternId
End Enum

Function GetPatternName(PatternId As Long) As String
  Dim R As String
  Select Case PatternId
    Case UIA_AnnotationPatternId: R = "Annotation"
    Case UIA_AnnotationPatternId: R = "CustomNavigation"
    Case UIA_DockPatternId: R = "Dock"
    Case UIA_DragPatternId: R = "Drag"
    Case UIA_DropTargetPatternId: R = "DropTarget"
    Case UIA_ExpandCollapsePatternId: R = "ExpandCollapse"
    Case UIA_GridItemPatternId: R = "GridItem"
    Case UIA_GridPatternId: R = "GridPattern"
    Case UIA_InvokePatternId: R = "InvokePattern"
    Case UIA_ItemContainerPatternId: R = "ItemContainer"
    Case UIA_LegacyIAccessiblePatternId: R = "LegacyIAccessiblePattern"
    Case UIA_MultipleViewPatternId: R = "MultipleView"
    Case UIA_ObjectModelPatternId: R = "ObjectModel"
    Case UIA_RangeValuePatternId: R = "RangeValue"
    Case UIA_ScrollItemPatternId: R = "ScrollItemPattern"
    Case UIA_ScrollPatternId: R = "Scroll"
    Case UIA_SelectionItemPatternId: R = "SelectionItem"
    Case UIA_SelectionPatternId: R = "Selection"
    Case UIA_SelectionPattern2Id: R = "Selection2"
    Case UIA_SpreadsheetItemPatternId: R = "SpreadsheetItem"
    Case UIA_SpreadsheetPatternId: R = "Spreadsheet"
    Case UIA_StylesPatternId: R = "Styles"
    Case UIA_SynchronizedInputPatternId: R = "SynchronizedInput"
    Case UIA_TableItemPatternId: R = "TableItem"
    Case UIA_TablePatternId: R = "Table"
    Case UIA_TextChildPatternId: R = "TextChild"
    Case UIA_TextEditPatternId: R = "TextEdit"
    Case UIA_TextPatternId: R = "Text"
    Case UIA_TextPattern2Id: R = "Text2"
    Case UIA_TogglePatternId: R = "Toggle"
    Case UIA_TransformPatternId: R = "Transform"
    Case UIA_TransformPattern2Id: R = "Transform2"
    Case UIA_ValuePatternId: R = "Value"
    Case UIA_VirtualizedItemPatternId: R = "VirtualizedItem"
    Case UIA_WindowPatternId: R = "Window"
  End Select
  GetPatternName = R
End Function

Public Function HasPattern(UIAElement As IUIAutomationElement, PatternId As Long) As Boolean
  Dim Pattern As IUnknown
  On Error Resume Next
  Set Pattern = UIAElement.GetCurrentPattern(PatternId)
  On Error GoTo 0
  HasPattern = Not Pattern Is Nothing
End Function

