VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIAutomationDictionaries"
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

'https://excelmacromastery.com/vba-dictionary/

'Dictionaries of TypeIDs
Public ControlTypeIDs As Scripting.dictionary 'Requires a reference Windows Scripting Runtime
Public AllPropertyIDs As Scripting.dictionary
Public NavigablePropertyIDs As Scripting.dictionary

Public Sub Initialise()
  InitialiseControlTypeIDsDictionary
  InitialiseAllPropertyIDsDictionary
End Sub

Private Sub InitialiseControlTypeIDsDictionary()
  'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controltype-ids
  'Set ControlTypeIDs = CreateObject("Scripting.Dictionary")
  Set ControlTypeIDs = New Scripting.dictionary
  With ControlTypeIDs
    .Add UIA_AppBarControlTypeId, "AppBar"
    .Add UIA_ButtonControlTypeId, "Button"
    .Add UIA_CalendarControlTypeId, "Calendar"
    .Add UIA_CheckBoxControlTypeId, "CheckBox"
    .Add UIA_ComboBoxControlTypeId, "ComboBox"
    .Add UIA_CustomControlTypeId, "Custom"
    .Add UIA_DataGridControlTypeId, "DataGrid"
    .Add UIA_DataItemControlTypeId, "DataItem"
    .Add UIA_DocumentControlTypeId, "Document"
    .Add UIA_EditControlTypeId, "Edit"
    .Add UIA_GroupControlTypeId, "Group"
    .Add UIA_HeaderControlTypeId, "Header"
    .Add UIA_HeaderItemControlTypeId, "HeaderItem"
    .Add UIA_HyperlinkControlTypeId, "Hyperlink"
    .Add UIA_ImageControlTypeId, "Image"
    .Add UIA_ListControlTypeId, "List"
    .Add UIA_ListItemControlTypeId, "ListItem"
    .Add UIA_MenuControlTypeId, "Menu"
    .Add UIA_MenuBarControlTypeId, "MenuBar"
    .Add UIA_MenuItemControlTypeId, "MenuItem"
    .Add UIA_PaneControlTypeId, "Pane"
    .Add UIA_ProgressBarControlTypeId, "ProgressBar"
    .Add UIA_RadioButtonControlTypeId, "RadioButton"
    .Add UIA_ScrollBarControlTypeId, "ScrollBar"
    .Add UIA_SemanticZoomControlTypeId, "SemanticZoom"
    .Add UIA_SeparatorControlTypeId, "Separator"
    .Add UIA_SliderControlTypeId, "Slider"
    .Add UIA_SpinnerControlTypeId, "Spinner"
    .Add UIA_SplitButtonControlTypeId, "SplitButton"
    .Add UIA_StatusBarControlTypeId, "StatusBar"
    .Add UIA_TabControlTypeId, "Tab"
    .Add UIA_TabItemControlTypeId, "TabItem"
    .Add UIA_TableControlTypeId, "Table"
    .Add UIA_TextControlTypeId, "Text"
    .Add UIA_ThumbControlTypeId, "Thumb"
    .Add UIA_TitleBarControlTypeId, "TitleBar"
    .Add UIA_ToolBarControlTypeId, "ToolBar"
    .Add UIA_ToolTipControlTypeId, "ToolTip"
    .Add UIA_TreeControlTypeId, "Tree"
    .Add UIA_TreeItemControlTypeId, "TreeItem"
    .Add UIA_WindowControlTypeId, "Window"
  End With
End Sub

'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-automation-element-propids

Private Sub InitialiseAllPropertyIDsDictionary()
  'AllPropertyIDs are a superset of NavigablePropertyIDs
  InitialiseNavigablePropertyIDsDictionary

  Set AllPropertyIDs = CreateObject("Scripting.Dictionary")
  Set AllPropertyIDs = New Scripting.dictionary
  Dim Key As Variant
  For Each Key In NavigablePropertyIDs
    AllPropertyIDs.Add Key, NavigablePropertyIDs(Key)
  Next Key
  
  'TODO how to deal with :
  '  UIA_AnnotationObjectsPropertyId, UIA_ControllerForPropertyId, UIA_DescribedByPropertyId,
  '  UIA_FlowsFromPropertyId, UIA_FlowsToPropertyId, UIA_LabeledByPropertyId
  '? - These cause an error!
  
  With AllPropertyIDs
'    .Add UIA_AnnotationObjectsPropertyId, "AnnotationObjects"
    .Add UIA_AnnotationTargetPropertyId, "AnnotationTarget"
    .Add UIA_AnnotationTypesPropertyId, "AnnotationTypes"
'    .Add UIA_ControllerForPropertyId, "ControllerFor"
'    .Add UIA_DescribedByPropertyId, "DescribedBy"
    .Add UIA_DragGrabbedItemsPropertyId, "DragGrabbedItems"
'    .Add UIA_FlowsFromPropertyId, "FlowsFrom"
'    .Add UIA_FlowsToPropertyId, "FlowsTo"
    .Add UIA_GridItemContainingGridPropertyId, "GridItemContainingGrid"
'    .Add UIA_LabeledByPropertyId, "LabeledBy"
    .Add UIA_LegacyIAccessibleSelectionPropertyId, "LegacyIAccessibleSelection"
    .Add UIA_Selection2FirstSelectedItemPropertyId, "Selection2FirstSelectedItem"
    .Add UIA_Selection2LastSelectedItemPropertyId, "Selection2LastSelectedItem"
    .Add UIA_SelectionItemSelectionContainerPropertyId, "SelectionItemSelectionContainer"
    .Add UIA_SelectionSelectionPropertyId, "SelectionSelection"
    .Add UIA_SpreadsheetItemAnnotationObjectsPropertyId, "SpreadsheetItemAnnotationObjects"
    .Add UIA_TableColumnHeadersPropertyId, "TableColumnHeaders"
    .Add UIA_TableItemColumnHeaderItemsPropertyId, "TableItemColumnHeaderItems"
    .Add UIA_TableItemRowHeaderItemsPropertyId, "TableItemRowHeaderItems"
    .Add UIA_TableRowHeadersPropertyId, "TableRowHeaders"
  End With
    
End Sub

Private Sub InitialiseNavigablePropertyIDsDictionary()

  Set NavigablePropertyIDs = CreateObject("Scripting.Dictionary")
  Set NavigablePropertyIDs = New Scripting.dictionary
  'TODO how to deal with :
  '  UIA_AnnotationObjectsPropertyId, UIA_ControllerForPropertyId, UIA_DescribedByPropertyId,
  '  UIA_FlowsFromPropertyId, UIA_FlowsToPropertyId, UIA_LabeledByPropertyId
  '? - These cause an error!
  
  With NavigablePropertyIDs
    .Add UIA_AcceleratorKeyPropertyId, "AcceleratorKey"
    .Add UIA_AccessKeyPropertyId, "AccessKey"
'    .Add UIA_AnnotationObjectsPropertyId, "AnnotationObjects"
'    .Add UIA_AnnotationTargetPropertyId, "AnnotationTarget"
'    .Add UIA_AnnotationTypesPropertyId, "AnnotationTypes"
    .Add UIA_AriaPropertiesPropertyId, "AriaProperties"
    .Add UIA_AriaRolePropertyId, "AriaRole" 'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-ariaspecification
    .Add UIA_AutomationIdPropertyId, "AutomationId"
    .Add UIA_BoundingRectanglePropertyId, "BoundingRectangle"
    .Add UIA_CenterPointPropertyId, "CenterPoint"
    .Add UIA_ClassNamePropertyId, "ClassName"
    .Add UIA_ClickablePointPropertyId, "ClickablePoint"
'    .Add UIA_ControllerForPropertyId, "ControllerFor"
    'We don't strictly need the ControlType as the a property type but make it available anyway!
    .Add UIA_ControlTypePropertyId, "ControlType"
    .Add UIA_CulturePropertyId, "Culture"
'    .Add UIA_DescribedByPropertyId, "DescribedBy"
    .Add UIA_DockDockPositionPropertyId, "DockDockPosition"
    .Add UIA_DragDropEffectPropertyId, "DragDropEffect"
    .Add UIA_DragDropEffectsPropertyId, "DragDropEffects"
'    .Add UIA_DragGrabbedItemsPropertyId, "DragGrabbedItems"
    .Add UIA_DragIsGrabbedPropertyId, "DragIsGrabbed"
    .Add UIA_DropTargetDropTargetEffectPropertyId, "DropTargetDropTargetEffect"
    .Add UIA_DropTargetDropTargetEffectsPropertyId, "DropTargetDropTargetEffects"
    .Add UIA_ExpandCollapseExpandCollapseStatePropertyId, "ExpandCollapseExpandCollapseState"
    .Add UIA_FillColorPropertyId, "FillColor"
    .Add UIA_FillTypePropertyId, "FillType"
'    .Add UIA_FlowsFromPropertyId, "FlowsFrom"
'    .Add UIA_FlowsToPropertyId, "FlowsTo"
    .Add UIA_FrameworkIdPropertyId, "FrameworkId"
    .Add UIA_FullDescriptionPropertyId, "FullDescription"
    .Add UIA_GridColumnCountPropertyId, "GridColumnCount"
    .Add UIA_GridItemColumnPropertyId, "GridItemColumn"
    .Add UIA_GridItemColumnSpanPropertyId, "GridItemColumnSpan"
'    .Add UIA_GridItemContainingGridPropertyId, "GridItemContainingGrid"
    .Add UIA_GridItemRowPropertyId, "GridItemRow"
    .Add UIA_GridItemRowSpanPropertyId, "GridItemRowSpan"
    .Add UIA_GridRowCountPropertyId, "GridRowCount"
    .Add UIA_HasKeyboardFocusPropertyId, "HasKeyboardFocus"
    .Add UIA_HeadingLevelPropertyId, "HeadingLevel"
    .Add UIA_HelpTextPropertyId, "HelpText"
    .Add UIA_IsAnnotationPatternAvailablePropertyId, "IsAnnotationPatternAvailable"
    .Add UIA_IsContentElementPropertyId, "IsContentElement"
    .Add UIA_IsControlElementPropertyId, "IsControlElement"
    .Add UIA_IsCustomNavigationPatternAvailablePropertyId, "IsCustomNavigationPatternAvailable"
    .Add UIA_IsDataValidForFormPropertyId, "IsDataValidForForm"
    .Add UIA_IsDialogPropertyId, "IsDialog"
    .Add UIA_IsDockPatternAvailablePropertyId, "IsDockPatternAvailable"
    .Add UIA_IsDragPatternAvailablePropertyId, "IsDragPatternAvailable"
    .Add UIA_IsDropTargetPatternAvailablePropertyId, "IsDropTargetPatternAvailable"
    .Add UIA_IsEnabledPropertyId, "IsEnabled"
    .Add UIA_IsExpandCollapsePatternAvailablePropertyId, "IsExpandCollapsePatternAvailable"
    .Add UIA_IsGridItemPatternAvailablePropertyId, "IsGridItemPatternAvailable"
    .Add UIA_IsGridPatternAvailablePropertyId, "IsGridPatternAvailable"
    .Add UIA_IsInvokePatternAvailablePropertyId, "IsInvokePatternAvailable"
    .Add UIA_IsItemContainerPatternAvailablePropertyId, "IsItemContainerPatternAvailable"
    .Add UIA_IsKeyboardFocusablePropertyId, "IsKeyboardFocusable"
    .Add UIA_IsLegacyIAccessiblePatternAvailablePropertyId, "IsLegacyIAccessiblePatternAvailable"
    .Add UIA_IsMultipleViewPatternAvailablePropertyId, "IsMultipleViewPatternAvailable"
    .Add UIA_IsObjectModelPatternAvailablePropertyId, "IsObjectModelPatternAvailable"
    .Add UIA_IsOffscreenPropertyId, "IsOffscreen"
    .Add UIA_IsPasswordPropertyId, "IsPassword"
    .Add UIA_IsPeripheralPropertyId, "IsPeripheral"
    .Add UIA_IsRangeValuePatternAvailablePropertyId, "IsRangeValuePatternAvailable"
    .Add UIA_IsRequiredForFormPropertyId, "IsRequiredForForm"
    .Add UIA_IsScrollItemPatternAvailablePropertyId, "IsScrollItemPatternAvailable"
    .Add UIA_IsScrollPatternAvailablePropertyId, "IsScrollPatternAvailable"
    .Add UIA_IsSelectionItemPatternAvailablePropertyId, "IsSelectionItemPatternAvailable"
    .Add UIA_IsSelectionPatternAvailablePropertyId, "IsSelectionPatternAvailable"
    .Add UIA_IsSelectionPattern2AvailablePropertyId, "IsSelectionPattern2Available"
    .Add UIA_IsSpreadsheetItemPatternAvailablePropertyId, "IsSpreadsheetItemPatternAvailable"
    .Add UIA_IsSpreadsheetPatternAvailablePropertyId, "IsSpreadsheetPatternAvailable"
    .Add UIA_IsStylesPatternAvailablePropertyId, "IsStylesPatternAvailable"
    .Add UIA_IsSynchronizedInputPatternAvailablePropertyId, "IsSynchronizedInputPatternAvailable"
    .Add UIA_IsTableItemPatternAvailablePropertyId, "IsTableItemPatternAvailable"
    .Add UIA_IsTablePatternAvailablePropertyId, "IsTablePatternAvailable"
    .Add UIA_IsTextChildPatternAvailablePropertyId, "IsTextChildPatternAvailable"
    .Add UIA_IsTextEditPatternAvailablePropertyId, "IsTextEditPatternAvailable"
    .Add UIA_IsTextPattern2AvailablePropertyId, "IsTextPattern2Available"
    .Add UIA_IsTextPatternAvailablePropertyId, "IsTextPatternAvailable"
    .Add UIA_IsTogglePatternAvailablePropertyId, "IsTogglePatternAvailable"
    .Add UIA_IsTransformPattern2AvailablePropertyId, "IsTransformPattern2Availabl"
    .Add UIA_IsTransformPatternAvailablePropertyId, "IsTransformPatternAvailable"
    .Add UIA_IsValuePatternAvailablePropertyId, "IsValuePatternAvailable"
    .Add UIA_IsVirtualizedItemPatternAvailablePropertyId, "IsVirtualizedItemPatternAvailable"
    .Add UIA_IsWindowPatternAvailablePropertyId, "IsWindowPatternAvailable"
    .Add UIA_ItemStatusPropertyId, "ItemStatus"
    .Add UIA_ItemTypePropertyId, "ItemType"
'    .Add UIA_LabeledByPropertyId, "LabeledBy"
    .Add UIA_LandmarkTypePropertyId, "LandmarkType"
    .Add UIA_LegacyIAccessibleChildIdPropertyId, "LegacyIAccessibleChildId"
    .Add UIA_LegacyIAccessibleDefaultActionPropertyId, "LegacyIAccessibleDefaultAction"
    .Add UIA_LegacyIAccessibleDescriptionPropertyId, "LegacyIAccessibleDescription"
    .Add UIA_LegacyIAccessibleHelpPropertyId, "LegacyIAccessibleHelp"
    .Add UIA_LegacyIAccessibleKeyboardShortcutPropertyId, "LegacyIAccessibleKeyboardShortcut"
    .Add UIA_LegacyIAccessibleNamePropertyId, "LegacyIAccessibleName"
    .Add UIA_LegacyIAccessibleRolePropertyId, "LegacyIAccessibleRole"
'    .Add UIA_LegacyIAccessibleSelectionPropertyId, "LegacyIAccessibleSelection"
    .Add UIA_LegacyIAccessibleStatePropertyId, "LegacyIAccessibleState"
    .Add UIA_LegacyIAccessibleValuePropertyId, "LegacyIAccessibleValue"
    .Add UIA_LevelPropertyId, "Level"
    .Add UIA_LiveSettingPropertyId, "LiveSetting"
    .Add UIA_LocalizedControlTypePropertyId, "LocalizedControlType"
    .Add UIA_LocalizedLandmarkTypePropertyId, "LocalizedLandmarkType"
    .Add UIA_MultipleViewCurrentViewPropertyId, "MultipleViewCurrentView"
    .Add UIA_MultipleViewSupportedViewsPropertyId, "MultipleViewSupportedViews"
    .Add UIA_NamePropertyId, "Name"
    .Add UIA_NativeWindowHandlePropertyId, "NativeWindowHandle"
    .Add UIA_OptimizeForVisualContentPropertyId, "OptimizeForVisualContent"
    .Add UIA_OrientationPropertyId, "Orientation"
    .Add UIA_OutlineColorPropertyId, "OutlineColor"
    .Add UIA_OutlineThicknessPropertyId, "OutlineThickness"
    .Add UIA_PositionInSetPropertyId, "PositionInSet"
    .Add UIA_ProcessIdPropertyId, "ProcessId"
    .Add UIA_ProviderDescriptionPropertyId, "ProviderDescription"
    .Add UIA_RangeValueIsReadOnlyPropertyId, "RangeValueIsReadOnly"
    .Add UIA_RangeValueLargeChangePropertyId, "RangeValueLargeChange"
    .Add UIA_RangeValueMaximumPropertyId, "RangeValueMaximum"
    .Add UIA_RangeValueMinimumPropertyId, "RangeValueMinimum"
    .Add UIA_RangeValueSmallChangePropertyId, "RangeValueSmallChange"
    .Add UIA_RangeValueValuePropertyId, "RangeValueValue"
    .Add UIA_RotationPropertyId, "Rotation"
    .Add UIA_RuntimeIdPropertyId, "RuntimeId"
    .Add UIA_ScrollHorizontallyScrollablePropertyId, "ScrollHorizontallyScrollable"
    .Add UIA_ScrollHorizontalScrollPercentPropertyId, "ScrollHorizontalScrollPercent"
    .Add UIA_ScrollHorizontalViewSizePropertyId, "ScrollHorizontalViewSize"
    .Add UIA_ScrollVerticallyScrollablePropertyId, "ScrollVerticallyScrollable"
    .Add UIA_ScrollVerticalScrollPercentPropertyId, "ScrollVerticalScrollPercent"
    .Add UIA_ScrollVerticalViewSizePropertyId, "ScrollVerticalViewSize"
'    .Add UIA_Selection2FirstSelectedItemPropertyId, "Selection2FirstSelectedItem"
    .Add UIA_Selection2ItemCountPropertyId, "Selection2ItemCount"
'    .Add UIA_Selection2LastSelectedItemPropertyId, "Selection2LastSelectedItem"
    .Add UIA_SelectionIsSelectionRequiredPropertyId, "SelectionIsSelectionRequired"
    .Add UIA_SelectionCanSelectMultiplePropertyId, "SelectionCanSelectMultiple"
    .Add UIA_SelectionItemIsSelectedPropertyId, "SelectionItemIsSelected"
'    .Add UIA_SelectionItemSelectionContainerPropertyId, "SelectionItemSelectionContainer"
'    .Add UIA_SelectionSelectionPropertyId, "SelectionSelection"
    .Add UIA_SizePropertyId, "Size"
    .Add UIA_SizeOfSetPropertyId, "SizeOfSet"
'    .Add UIA_SpreadsheetItemAnnotationObjectsPropertyId, "SpreadsheetItemAnnotationObjects"
    .Add UIA_SpreadsheetItemAnnotationTypesPropertyId, "SpreadsheetItemAnnotationTypes"
    .Add UIA_SpreadsheetItemFormulaPropertyId, "SpreadsheetItemFormula"
    .Add UIA_StylesExtendedPropertiesPropertyId, "StylesExtendedProperties"
    .Add UIA_StylesFillColorPropertyId, "StylesFillColor"
    .Add UIA_StylesFillPatternColorPropertyId, "StylesFillPatternColor"
    .Add UIA_StylesFillPatternStylePropertyId, "StylesFillPatternStyle"
    .Add UIA_StylesShapePropertyId, "StylesShape"
    .Add UIA_StylesStyleIdPropertyId, "StylesStyleId"
    .Add UIA_StylesStyleNamePropertyId, "StylesStyleName"
'    .Add UIA_TableColumnHeadersPropertyId, "TableColumnHeaders"
'    .Add UIA_TableItemColumnHeaderItemsPropertyId, "TableItemColumnHeaderItems"
'    .Add UIA_TableItemRowHeaderItemsPropertyId, "TableItemRowHeaderItems"
'    .Add UIA_TableRowHeadersPropertyId, "TableRowHeaders"
    .Add UIA_TableRowOrColumnMajorPropertyId, "TableRowOrColumnMajor"
    .Add UIA_ToggleToggleStatePropertyId, "ToggleToggleState"
    .Add UIA_Transform2CanZoomPropertyId, "Transform2CanZoom"
    .Add UIA_Transform2ZoomLevelPropertyId, "Transform2ZoomLevel"
    .Add UIA_Transform2ZoomMaximumPropertyId, "Transform2ZoomMaximumProperty"
    .Add UIA_Transform2ZoomMinimumPropertyId, "Transform2ZoomMinimum"
    .Add UIA_TransformCanMovePropertyId, "TransformCanMove"
    .Add UIA_TransformCanResizePropertyId, "TransformCanResize"
    .Add UIA_TransformCanRotatePropertyId, "TransformCanRotate"
    .Add UIA_ValueIsReadOnlyPropertyId, "ValueIsReadOnly"
    .Add UIA_ValueValuePropertyId, "ValueValue"
    .Add UIA_VisualEffectsPropertyId, "VisualEffects"
    .Add UIA_WindowCanMaximizePropertyId, "WindowCanMaximize"
    .Add UIA_WindowCanMinimizePropertyId, "WindowCanMinimize"
    .Add UIA_WindowIsModalPropertyId, "WindowIsModal"
    .Add UIA_WindowIsTopmostPropertyId, "WindowIsTopmost"
    .Add UIA_WindowWindowInteractionStatePropertyId, "WindowWindowInteractionState"
    .Add UIA_WindowWindowVisualStatePropertyId, "WindowWindowVisualState"
  End With

End Sub

Public Function ValueExists(Value As String, dictionary As Scripting.dictionary) As Boolean
  Dim Key As Variant
  For Each Key In dictionary.Keys
    If dictionary(Key) = Value Then
      ValueExists = True
      Exit Function
    End If
  Next Key
  ValueExists = False
End Function

