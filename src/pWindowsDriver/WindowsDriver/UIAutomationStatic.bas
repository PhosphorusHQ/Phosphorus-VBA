Attribute VB_Name = "UIAutomationStatic"
'@Folder WindowsDriver
Option Explicit

Public Function GetPatternName(patternId As Long) As String
  Select Case patternId
    'Members of UIAutomationClient.UIA_PatternIds
    Case UIA_AnnotationPatternId: GetPatternName = "Annotation"
    Case UIA_CustomNavigationPatternId: GetPatternName = "Custom Navigation"
    Case UIA_DockPatternId: GetPatternName = "Dock"
    Case UIA_DragPatternId: GetPatternName = "Drag"
    Case UIA_DropTargetPatternId: GetPatternName = "Drop Target"
    Case UIA_ExpandCollapsePatternId: GetPatternName = "Expand Collapse"
    Case UIA_GridItemPatternId: GetPatternName = "Grid Item"
    Case UIA_GridPatternId: GetPatternName = "Grid"
    Case UIA_InvokePatternId: GetPatternName = "Invoke"
    Case UIA_ItemContainerPatternId: GetPatternName = "Item Container"
    Case UIA_LegacyIAccessiblePatternId: GetPatternName = "Legacy IAccessible"
    Case UIA_MultipleViewPatternId: GetPatternName = "Multiple View"
    Case UIA_ObjectModelPatternId: GetPatternName = "Object Model"
    Case UIA_RangeValuePatternId: GetPatternName = "Range Value"
    Case UIA_ScrollItemPatternId: GetPatternName = "Scroll Item"
    Case UIA_ScrollPatternId: GetPatternName = "Scroll"
    Case UIA_SelectionItemPatternId: GetPatternName = "Selection Item"
    Case UIA_SelectionPatternId: GetPatternName = "Selection"
    Case UIA_SpreadsheetPatternId: GetPatternName = "Spreadsheet"
    Case UIA_SpreadsheetItemPatternId: GetPatternName = "SpreadsheetItem"
    Case UIA_StylesPatternId: GetPatternName = "Styles"
    Case UIA_SynchronizedInputPatternId: GetPatternName = "Synchronized Input"
    Case UIA_TableItemPatternId: GetPatternName = "Table Item"
    Case UIA_TablePatternId: GetPatternName = "Table"
    Case UIA_TextChildPatternId: GetPatternName = "TextChild"
    Case UIA_TextEditPatternId: GetPatternName = "TextEdit"
    Case UIA_TextPatternId: GetPatternName = "Text"
    Case UIA_TextPattern2Id: GetPatternName = "Text2"
    Case UIA_TogglePatternId: GetPatternName = "Toggle"
    Case UIA_TransformPatternId: GetPatternName = "Transform"
    Case UIA_TransformPattern2Id: GetPatternName = "Transform2"
    Case UIA_ValuePatternId: GetPatternName = "Value"
    Case UIA_VirtualizedItemPatternId: GetPatternName = "Virtualized Item"
    Case UIA_WindowPatternId: GetPatternName = "Window"
    Case Else: GetPatternName = "Unknown (ID: " & patternId & ")"
  End Select
End Function
