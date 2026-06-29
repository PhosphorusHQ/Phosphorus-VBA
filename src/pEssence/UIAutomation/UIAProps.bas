Attribute VB_Name = "UIAProps"
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

'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-automation-element-propids
Public Enum UIAProperties 'UIA_PropertyIds
  AcceleratorKey = UIA_AcceleratorKeyPropertyId
  AccessKey = UIA_AccessKeyPropertyId
  AnnotationObjects = UIA_AnnotationObjectsPropertyId
'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-ariaspecification
  AnnotationTarget = UIA_AnnotationTargetPropertyId
  AnnotationTypes = UIA_AnnotationTypesPropertyId
  AriaProperties = UIA_AriaPropertiesPropertyId
  AriaRole = UIA_AriaRolePropertyId
  AutomationId = UIA_AutomationIdPropertyId
  BoundingRectangle = UIA_BoundingRectanglePropertyId
  CenterPoint = UIA_CenterPointPropertyId
  ClassName = UIA_ClassNamePropertyId
  ClickablePoint = UIA_ClickablePointPropertyId
  ControllerFor = UIA_ControllerForPropertyId
  ControlType = UIA_ControlTypePropertyId
  Culture = UIA_CulturePropertyId
  DescribedBy = UIA_DescribedByPropertyId
  DockDockPosition = UIA_DockDockPositionPropertyId
  DragDropEffect = UIA_DragDropEffectPropertyId
  DragDropEffects = UIA_DragDropEffectsPropertyId
  DragGrabbedItems = UIA_DragGrabbedItemsPropertyId
  DragIsGrabbed = UIA_DragIsGrabbedPropertyId
  DropTargetDropTargetEffect = UIA_DropTargetDropTargetEffectPropertyId
  DropTargetDropTargetEffects = UIA_DropTargetDropTargetEffectsPropertyId
  ExpandCollapseExpandCollapseState = UIA_ExpandCollapseExpandCollapseStatePropertyId
  FillColor = UIA_FillColorPropertyId
  FillType = UIA_FillTypePropertyId
  FlowsFrom = UIA_FlowsFromPropertyId
  FlowsTo = UIA_FlowsToPropertyId
  FrameworkID = UIA_FrameworkIdPropertyId
  FullDescription = UIA_FullDescriptionPropertyId
  GridColumnCount = UIA_GridColumnCountPropertyId
  GridItemColumn = UIA_GridItemColumnPropertyId
  GridItemColumnSpan = UIA_GridItemColumnSpanPropertyId
  GridItemContainingGrid = UIA_GridItemContainingGridPropertyId
  GridItemRow = UIA_GridItemRowPropertyId
  GridItemRowSpan = UIA_GridItemRowSpanPropertyId
  GridRow = UIA_GridRowCountPropertyId
  HasKeyboardFocus = UIA_HasKeyboardFocusPropertyId
  HeadingLevel = UIA_HeadingLevelPropertyId
  HelpText = UIA_HelpTextPropertyId
  IsAnnotationPatternAvailable = UIA_IsAnnotationPatternAvailablePropertyId
  IsContentElement = UIA_IsContentElementPropertyId
  IsControlElement = UIA_IsControlElementPropertyId
  IsCustomNavigationPatternAvailable = UIA_IsCustomNavigationPatternAvailablePropertyId
  IsDataValidForForm = UIA_IsDataValidForFormPropertyId
  IsDialog = UIA_IsDialogPropertyId
  IsDockPatternAvailable = UIA_IsDockPatternAvailablePropertyId
  IsDragPatternAvailable = UIA_IsDragPatternAvailablePropertyId
  IsDropTargetPatternAvailable = UIA_IsDropTargetPatternAvailablePropertyId
  IsEnabled = UIA_IsEnabledPropertyId
  IsExpandCollapsePatternAvailable = UIA_IsExpandCollapsePatternAvailablePropertyId
  IsGridItemPatternAvailable = UIA_IsGridItemPatternAvailablePropertyId
  IsGridPatternAvailable = UIA_IsGridPatternAvailablePropertyId
  IsInvokePatternAvailable = UIA_IsInvokePatternAvailablePropertyId
  IsItemContainerPatternAvailable = UIA_IsItemContainerPatternAvailablePropertyId
  IsKeyboardFocusable = UIA_IsKeyboardFocusablePropertyId
  IsLegacyIAccessiblePatternAvailable = UIA_IsLegacyIAccessiblePatternAvailablePropertyId
  IsMultipleViewPatternAvailable = UIA_IsMultipleViewPatternAvailablePropertyId
  IsObjectModelPatternAvailable = UIA_IsObjectModelPatternAvailablePropertyId
  IsOffscreen = UIA_IsOffscreenPropertyId
  IsPassword = UIA_IsPasswordPropertyId
  IsPeripheral = UIA_IsPeripheralPropertyId
  IsRangeValuePatternAvailable = UIA_IsRangeValuePatternAvailablePropertyId
  IsRequiredForForm = UIA_IsRequiredForFormPropertyId
  IsScrollItemPatternAvailable = UIA_IsScrollItemPatternAvailablePropertyId
  IsScrollPatternAvailable = UIA_IsScrollPatternAvailablePropertyId
  IsSelectionItemPatternAvailable = UIA_IsSelectionItemPatternAvailablePropertyId
  IsSelectionPatternAvailable = UIA_IsSelectionPatternAvailablePropertyId
  IsSelectionPattern2Available = UIA_IsSelectionPattern2AvailablePropertyId
  IsSpreadsheetItemPatternAvailable = UIA_IsSpreadsheetItemPatternAvailablePropertyId
  IsSpreadsheetPatternAvailable = UIA_IsSpreadsheetPatternAvailablePropertyId
  IsStylesPatternAvailable = UIA_IsStylesPatternAvailablePropertyId
  IsSynchronizedInputPatternAvailable = UIA_IsSynchronizedInputPatternAvailablePropertyId
  IsTableItemPatternAvailable = UIA_IsTableItemPatternAvailablePropertyId
  IsTablePatternAvailable = UIA_IsTablePatternAvailablePropertyId
  IsTextChildPatternAvailable = UIA_IsTextChildPatternAvailablePropertyId
  IsTextEditPatternAvailable = UIA_IsTextEditPatternAvailablePropertyId
  IsTextPattern2Available = UIA_IsTextPattern2AvailablePropertyId
  IsTextPatternAvailable = UIA_IsTextPatternAvailablePropertyId
  IsTogglePatternAvailable = UIA_IsTogglePatternAvailablePropertyId
  IsTransformPattern2Availabl = UIA_IsTransformPattern2AvailablePropertyId
  IsTransformPatternAvailable = UIA_IsTransformPatternAvailablePropertyId
  IsValuePatternAvailable = UIA_IsValuePatternAvailablePropertyId
  IsVirtualizedItemPatternAvailable = UIA_IsVirtualizedItemPatternAvailablePropertyId
  IsWindowPatternAvailable = UIA_IsWindowPatternAvailablePropertyId
  ItemStatus = UIA_ItemStatusPropertyId
  ItemType = UIA_ItemTypePropertyId
  LabeledBy = UIA_LabeledByPropertyId
  LandmarkType = UIA_LandmarkTypePropertyId
  LegacyIAccessibleChildId = UIA_LegacyIAccessibleChildIdPropertyId
  LegacyIAccessibleDefaultAction = UIA_LegacyIAccessibleDefaultActionPropertyId
  LegacyIAccessibleDescription = UIA_LegacyIAccessibleDescriptionPropertyId
  LegacyIAccessibleHelp = UIA_LegacyIAccessibleHelpPropertyId
  LegacyIAccessibleKeyboardShortcut = UIA_LegacyIAccessibleKeyboardShortcutPropertyId
  LegacyIAccessibleName = UIA_LegacyIAccessibleNamePropertyId
  LegacyIAccessibleRole = UIA_LegacyIAccessibleRolePropertyId
  LegacyIAccessibleSelection = UIA_LegacyIAccessibleSelectionPropertyId
  LegacyIAccessibleState = UIA_LegacyIAccessibleStatePropertyId
  LegacyIAccessibleValue = UIA_LegacyIAccessibleValuePropertyId
  Level = UIA_LevelPropertyId
  LiveSetting = UIA_LiveSettingPropertyId
  LocalizedControlType = UIA_LocalizedControlTypePropertyId
  LocalizedLandmarkType = UIA_LocalizedLandmarkTypePropertyId
  MultipleViewCurrentView = UIA_MultipleViewCurrentViewPropertyId
  MultipleViewSupportedViews = UIA_MultipleViewSupportedViewsPropertyId
  Name = UIA_NamePropertyId
  NativeWindowHandle = UIA_NativeWindowHandlePropertyId
  OptimizeForVisualContent = UIA_OptimizeForVisualContentPropertyId
  Orientation = UIA_OrientationPropertyId
  OutlineColor = UIA_OutlineColorPropertyId
  OutlineThickness = UIA_OutlineThicknessPropertyId
  PositionInSet = UIA_PositionInSetPropertyId
  ProcessId = UIA_ProcessIdPropertyId
  ProviderDescription = UIA_ProviderDescriptionPropertyId
  RangeValueIsReadOnly = UIA_RangeValueIsReadOnlyPropertyId
  RangeValueLargeChange = UIA_RangeValueLargeChangePropertyId
  RangeValueMaximum = UIA_RangeValueMaximumPropertyId
  RangeValueMinimum = UIA_RangeValueMinimumPropertyId
  RangeValueSmallChange = UIA_RangeValueSmallChangePropertyId
  RangeValueValue = UIA_RangeValueValuePropertyId
  Rotation = UIA_RotationPropertyId
  RuntimeId = UIA_RuntimeIdPropertyId
  ScrollHorizontallyScrollable = UIA_ScrollHorizontallyScrollablePropertyId
  ScrollHorizontalScrollPercent = UIA_ScrollHorizontalScrollPercentPropertyId
  ScrollHorizontalViewSize = UIA_ScrollHorizontalViewSizePropertyId
  ScrollVerticallyScrollable = UIA_ScrollVerticallyScrollablePropertyId
  ScrollVerticalScrollPercent = UIA_ScrollVerticalScrollPercentPropertyId
  ScrollVerticalViewSize = UIA_ScrollVerticalViewSizePropertyId
  Selection2FirstSelectedItem = UIA_Selection2FirstSelectedItemPropertyId
  Selection2ItemCount = UIA_Selection2ItemCountPropertyId
  Selection2LastSelectedItem = UIA_Selection2LastSelectedItemPropertyId
  SelectionIsSelectionRequired = UIA_SelectionIsSelectionRequiredPropertyId
  SelectionCanSelectMultiple = UIA_SelectionCanSelectMultiplePropertyId
  SelectionItemIsSelected = UIA_SelectionItemIsSelectedPropertyId
  SelectionItemSelectionContainer = UIA_SelectionItemSelectionContainerPropertyId
  SelectionSelection = UIA_SelectionSelectionPropertyId
  SizeOfSet = UIA_SizeOfSetPropertyId
  Size = UIA_SizePropertyId
  SpreadsheetItemAnnotationObjects = UIA_SpreadsheetItemAnnotationObjectsPropertyId
  SpreadsheetItemAnnotationTypes = UIA_SpreadsheetItemAnnotationTypesPropertyId
  SpreadsheetItemFormula = UIA_SpreadsheetItemFormulaPropertyId
  StylesExtendedProperties = UIA_StylesExtendedPropertiesPropertyId
  StylesFillColor = UIA_StylesFillColorPropertyId
  StylesFillPatternColor = UIA_StylesFillPatternColorPropertyId
  StylesFillPatternStyle = UIA_StylesFillPatternStylePropertyId
  StylesShape = UIA_StylesShapePropertyId
  StylesStyleId = UIA_StylesStyleIdPropertyId
  StylesStyleName = UIA_StylesStyleNamePropertyId
  TableColumnHeaders = UIA_TableColumnHeadersPropertyId
  TableItemColumnHeaderItems = UIA_TableItemColumnHeaderItemsPropertyId
  TableItemRowHeaderItems = UIA_TableItemRowHeaderItemsPropertyId
  TableRowHeaders = UIA_TableRowHeadersPropertyId
  TableRowOrColumnMajor = UIA_TableRowOrColumnMajorPropertyId
  ToggleToggleState = UIA_ToggleToggleStatePropertyId
  Transform2CanZoom = UIA_Transform2CanZoomPropertyId
  Transform2ZoomLevel = UIA_Transform2ZoomLevelPropertyId
  Transform2ZoomMaximumProperty = UIA_Transform2ZoomMaximumPropertyId
  Transform2ZoomMinimum = UIA_Transform2ZoomMinimumPropertyId
  TransformCanMove = UIA_TransformCanMovePropertyId
  TransformCanResize = UIA_TransformCanResizePropertyId
  TransformCanRotate = UIA_TransformCanRotatePropertyId
  ValueIsReadOnly = UIA_ValueIsReadOnlyPropertyId
  ValueValue = UIA_ValueValuePropertyId
  VisualEffects = UIA_VisualEffectsPropertyId
  WindowCanMaximize = UIA_WindowCanMaximizePropertyId
  WindowCanMinimize = UIA_WindowCanMinimizePropertyId
  WindowIsModal = UIA_WindowIsModalPropertyId
  WindowIsTopmost = UIA_WindowIsTopmostPropertyId
  WindowWindowInteractionState = UIA_WindowWindowInteractionStatePropertyId
  WindowWindowVisualState = UIA_WindowWindowVisualStatePropertyId
End Enum

Public Enum UIAControlTypeIDs
  AppBar = UIA_ControlTypeIds.UIA_AppBarControlTypeId
  Button = UIA_ControlTypeIds.UIA_ButtonControlTypeId
  Calendar = UIA_ControlTypeIds.UIA_CalendarControlTypeId
  CheckBox = UIA_ControlTypeIds.UIA_CheckBoxControlTypeId
  ComboBox = UIA_ControlTypeIds.UIA_ComboBoxControlTypeId
  Custom = UIA_ControlTypeIds.UIA_CustomControlTypeId
  DataGrid = UIA_ControlTypeIds.UIA_DataGridControlTypeId
  DataItem = UIA_ControlTypeIds.UIA_DataItemControlTypeId
  Document = UIA_ControlTypeIds.UIA_DocumentControlTypeId
  Edit = UIA_ControlTypeIds.UIA_EditControlTypeId
  Group = UIA_ControlTypeIds.UIA_GroupControlTypeId
  Header = UIA_ControlTypeIds.UIA_HeaderControlTypeId
  HeaderItem = UIA_ControlTypeIds.UIA_HeaderItemControlTypeId
  Hyperlink = UIA_ControlTypeIds.UIA_HyperlinkControlTypeId
  Image = UIA_ControlTypeIds.UIA_ImageControlTypeId
  List = UIA_ControlTypeIds.UIA_ListControlTypeId
  ListItem = UIA_ControlTypeIds.UIA_ListItemControlTypeId
  Menu = UIA_ControlTypeIds.UIA_MenuControlTypeId
  MenuBar = UIA_ControlTypeIds.UIA_MenuBarControlTypeId
  MenuItem = UIA_ControlTypeIds.UIA_MenuItemControlTypeId
  Pane = UIA_ControlTypeIds.UIA_PaneControlTypeId
  ProgressBar = UIA_ControlTypeIds.UIA_ProgressBarControlTypeId
  RadioButton = UIA_ControlTypeIds.UIA_RadioButtonControlTypeId
  ScrollBar = UIA_ControlTypeIds.UIA_ScrollBarControlTypeId
  SemanticZoom = UIA_ControlTypeIds.UIA_SemanticZoomControlTypeId
  Separator = UIA_ControlTypeIds.UIA_SeparatorControlTypeId
  Slider = UIA_ControlTypeIds.UIA_SliderControlTypeId
  Spinner = UIA_ControlTypeIds.UIA_SpinnerControlTypeId
  SplitButton = UIA_ControlTypeIds.UIA_SplitButtonControlTypeId
  StatusBar = UIA_ControlTypeIds.UIA_StatusBarControlTypeId
  TabControl = UIA_ControlTypeIds.UIA_TabControlTypeId
  TabItem = UIA_ControlTypeIds.UIA_TabItemControlTypeId
  Table = UIA_ControlTypeIds.UIA_TableControlTypeId
  Text = UIA_ControlTypeIds.UIA_TextControlTypeId
  Thumb = UIA_ControlTypeIds.UIA_ThumbControlTypeId
  TitleBar = UIA_ControlTypeIds.UIA_TitleBarControlTypeId
  Toolbar = UIA_ControlTypeIds.UIA_ToolBarControlTypeId
  ToolTip = UIA_ControlTypeIds.UIA_ToolTipControlTypeId
  TreeControl = UIA_ControlTypeIds.UIA_TreeControlTypeId
  TreeItem = UIA_ControlTypeIds.UIA_TreeItemControlTypeId
  Window = UIA_ControlTypeIds.UIA_WindowControlTypeId
End Enum

Public Enum UIAWindowInteractionStates 'UIAutomationClient.WindowInteractionState
  BlockedByModalWindow = WindowInteractionState_BlockedByModalWindow
  Closing = WindowInteractionState_Closing
  NotResponding = WindowInteractionState_NotResponding
  ReadyForUserInteraction = WindowInteractionState_ReadyForUserInteraction
  running = WindowInteractionState_Running
End Enum

Public Type Rectangle
  Left As Double
  Top As Double
  Right As Double
  Bottom As Double
  Width As Double
  Height As Double
End Type

Function GetPropertyName(PropertyId As Long) As String
  Dim R As String
  Select Case PropertyId
    Case UIA_AcceleratorKeyPropertyId: R = "AcceleratorKey"
    Case UIA_AccessKeyPropertyId: R = "AccessKey"
    Case UIA_AnnotationObjectsPropertyId: R = "AnnotationObjects"
    Case UIA_AnnotationTargetPropertyId: R = "AnnotationTarget"
    Case UIA_AnnotationTypesPropertyId: R = "AnnotationTypes"
'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-ariaspecification
    Case UIA_AriaPropertiesPropertyId: R = "AriaProperties"
    Case UIA_AriaRolePropertyId: R = "AriaRole"
    Case UIA_AutomationIdPropertyId: R = "AutomationId"
    Case UIA_BoundingRectanglePropertyId: R = "BoundingRectangle"
    Case UIA_CenterPointPropertyId: R = "CenterPoint"
    Case UIA_ClassNamePropertyId: R = "ClassName"
    Case UIA_ClickablePointPropertyId: R = "ClickablePoint"
    Case UIA_ControllerForPropertyId: R = "ControllerFor"
    Case UIA_ControlTypePropertyId: R = "ControlType"
    Case UIA_CulturePropertyId: R = "Culture"
    Case UIA_DescribedByPropertyId: R = "DescribedBy"
    Case UIA_DockDockPositionPropertyId: R = "DockDockPosition"
    Case UIA_DragDropEffectPropertyId: R = "DragDropEffect"
    Case UIA_DragDropEffectsPropertyId: R = "DragDropEffects"
    Case UIA_DragGrabbedItemsPropertyId: R = "DragGrabbedItems"
    Case UIA_DragIsGrabbedPropertyId: R = "DragIsGrabbed"
    Case UIA_DropTargetDropTargetEffectPropertyId: R = "DropTargetDropTargetEffect"
    Case UIA_DropTargetDropTargetEffectsPropertyId: R = "DropTargetDropTargetEffects"
    Case UIA_ExpandCollapseExpandCollapseStatePropertyId: R = "ExpandCollapseExpandCollapseState"
    Case UIA_FillColorPropertyId: R = "FillColor"
    Case UIA_FillTypePropertyId: R = "FillType"
    Case UIA_FlowsFromPropertyId: R = "FlowsFrom"
    Case UIA_FlowsToPropertyId: R = "FlowsTo"
    Case UIA_FrameworkIdPropertyId: R = "FrameworkID"
    Case UIA_FullDescriptionPropertyId: R = "FullDescription"
    Case UIA_GridColumnCountPropertyId: R = "GridColumnCount"
    Case UIA_GridItemColumnPropertyId: R = "GridItemColumn"
    Case UIA_GridItemColumnSpanPropertyId: R = "GridItemColumnSpan"
    Case UIA_GridItemContainingGridPropertyId: R = "GridItemContainingGrid"
    Case UIA_GridItemRowPropertyId: R = "GridItemRow"
    Case UIA_GridItemRowSpanPropertyId: R = "GridItemRowSpan"
    Case UIA_GridRowCountPropertyId: R = "GridRow"
    Case UIA_GroupControlTypeId: R = "Group"
    Case UIA_HasKeyboardFocusPropertyId: R = "HasKeyboardFocus"
    Case UIA_HeadingLevelPropertyId: R = "HeadingLevel"
    Case UIA_HelpTextPropertyId: R = "HelpText"
    Case IsAnnotationPatternAvailable: R = "IsAnnotationPatternAvailable"
    Case UIA_IsContentElementPropertyId: R = "IsContentElement"
    Case UIA_IsControlElementPropertyId: R = "IsControlElement"
    Case UIA_IsCustomNavigationPatternAvailablePropertyId: R = "IsCustomNavigationPatternAvailable"
    Case UIA_IsDataValidForFormPropertyId: R = "IsDataValidForForm"
    Case UIA_IsDialogPropertyId: R = "IsDialog"
    Case UIA_IsDockPatternAvailablePropertyId: R = "IsDockPatternAvailable"
    Case UIA_IsDragPatternAvailablePropertyId: R = "IsDragPatternAvailable"
    Case UIA_IsDropTargetPatternAvailablePropertyId: R = "IsDropTargetPatternAvailable"
    Case UIA_IsEnabledPropertyId: R = "IsEnabled"
    Case UIA_IsExpandCollapsePatternAvailablePropertyId: R = "IsExpandCollapsePatternAvailable"
    Case UIA_IsGridItemPatternAvailablePropertyId: R = "IsGridItemPatternAvailable"
    Case UIA_IsGridPatternAvailablePropertyId: R = "IsGridPatternAvailable"
    Case UIA_IsInvokePatternAvailablePropertyId: R = "IsInvokePatternAvailable"
    Case UIA_IsItemContainerPatternAvailablePropertyId: R = "IsItemContainerPatternAvailable"
    Case UIA_IsKeyboardFocusablePropertyId: R = "IsKeyboardFocusable"
    Case UIA_IsLegacyIAccessiblePatternAvailablePropertyId: R = "IsLegacyIAccessiblePatternAvailable"
    Case UIA_IsMultipleViewPatternAvailablePropertyId: R = "IsMultipleViewPatternAvailable"
    Case UIA_IsObjectModelPatternAvailablePropertyId: R = "IsObjectModelPatternAvailable"
    Case UIA_IsOffscreenPropertyId: R = "IsOffscreen"
    Case UIA_IsPasswordPropertyId: R = "IsPassword"
    Case UIA_IsPeripheralPropertyId: R = "IsPeripheral"
    Case UIA_IsRangeValuePatternAvailablePropertyId: R = "IsRangeValuePatternAvailable"
    Case UIA_IsRequiredForFormPropertyId: R = "IsRequiredForForm"
    Case UIA_IsScrollItemPatternAvailablePropertyId: R = "IsScrollItemPatternAvailable"
    Case UIA_IsScrollPatternAvailablePropertyId: R = "IsScrollPatternAvailable"
    Case UIA_IsSelectionPatternAvailablePropertyId: R = "IsSelectionPatternAvailable"
    Case UIA_IsSelectionPattern2AvailablePropertyId: R = "IsSelectionPattern2Available"
    Case UIA_IsSpreadsheetItemPatternAvailablePropertyId: R = "IsSpreadsheetItemPatternAvailable"
    Case UIA_IsSpreadsheetPatternAvailablePropertyId: R = "IsSpreadsheetPatternAvailable"
    Case UIA_IsStylesPatternAvailablePropertyId: R = "IsStylesPatternAvailable"
    Case UIA_IsSynchronizedInputPatternAvailablePropertyId: R = "IsSynchronizedInputPatternAvailable"
    Case UIA_IsTableItemPatternAvailablePropertyId: R = "IsTableItemPatternAvailable"
    Case UIA_IsTablePatternAvailablePropertyId: R = "IsTablePatternAvailable"
    Case UIA_IsTextChildPatternAvailablePropertyId: R = "IsTextChildPatternAvailable"
    Case UIA_IsTextEditPatternAvailablePropertyId: R = "IsTextEditPatternAvailable"
    Case UIA_IsTextPatternAvailablePropertyId: R = "IsTextPatternAvailable"
    Case UIA_IsTextPattern2AvailablePropertyId: R = "IsTextPattern2Available"
    Case UIA_IsTogglePatternAvailablePropertyId: R = "IsTogglePatternAvailable"
    Case UIA_IsTransformPatternAvailablePropertyId: R = "IsTransformPatternAvailable"
    Case UIA_IsTransformPattern2AvailablePropertyId: R = "IsTransformPattern2Available"
    Case UIA_IsValuePatternAvailablePropertyId: R = "IsValuePatternAvailable"
    Case UIA_IsVirtualizedItemPatternAvailablePropertyId: R = "IsVirtualizedItemPatternAvailable"
    Case UIA_IsWindowPatternAvailablePropertyId: R = "IsWindowPatternAvailable"
    Case UIA_ItemStatusPropertyId: R = "ItemStatus"
    Case UIA_ItemTypePropertyId: R = "ItemType"
    Case UIA_LabeledByPropertyId: R = "LabeledBy"
    Case UIA_LandmarkTypePropertyId: R = "LandmarkType"
    Case UIA_LegacyIAccessibleChildIdPropertyId: R = "LegacyIAccessibleChildId"
    Case UIA_LegacyIAccessibleDefaultActionPropertyId: R = "LegacyIAccessibleDefaultAction"
    Case UIA_LegacyIAccessibleDescriptionPropertyId: R = "LegacyIAccessibleDescription"
    Case UIA_LegacyIAccessibleHelpPropertyId: R = "LegacyIAccessibleHelp"
    Case UIA_LegacyIAccessibleKeyboardShortcutPropertyId: R = "LegacyIAccessibleKeyboardShortcut"
    Case UIA_LegacyIAccessibleNamePropertyId: R = "LegacyIAccessibleName"
    Case UIA_LegacyIAccessibleRolePropertyId: R = "LegacyIAccessibleRole"
    Case UIA_LegacyIAccessibleSelectionPropertyId: R = "LegacyIAccessibleSelection"
    Case UIA_LegacyIAccessibleStatePropertyId: R = "LegacyIAccessibleState"
    Case UIA_LegacyIAccessibleValuePropertyId: R = "LegacyIAccessibleValue"
    Case UIA_LevelPropertyId: R = "Level"
    Case UIA_LiveSettingPropertyId: R = "LiveSetting"
    Case UIA_LocalizedControlTypePropertyId: R = "LocalizedControlType"
    Case UIA_LocalizedLandmarkTypePropertyId: R = "LocalizedLandmarkType"
    Case UIA_MultipleViewCurrentViewPropertyId: R = "MultipleViewCurrentView"
    Case UIA_MultipleViewSupportedViewsPropertyId: R = "MultipleViewSupportedViews"
    Case UIA_NamePropertyId: R = "Name"
    Case UIA_NativeWindowHandlePropertyId: R = "NativeWindowHandle"
    Case UIA_OptimizeForVisualContentPropertyId: R = "OptimizeForVisualContent"
    Case UIA_OrientationPropertyId: R = "Orientation"
    Case UIA_OutlineColorPropertyId: R = "OutlineColor"
    Case UIA_OutlineThicknessPropertyId: R = "OutlineThickness"
    Case UIA_PositionInSetPropertyId: R = "PositionInSet"
    Case UIA_ProcessIdPropertyId: R = "ProcessId"
    Case UIA_ProviderDescriptionPropertyId: R = "ProviderDescription"
    Case UIA_RangeValueIsReadOnlyPropertyId: R = "RangeValueIsReadOnly"
    Case UIA_RangeValueLargeChangePropertyId: R = "RangeValueLargeChange"
    Case UIA_RangeValueMaximumPropertyId: R = "RangeValueMaximum"
    Case UIA_RangeValueMinimumPropertyId: R = "RangeValueMinimum"
    Case UIA_RangeValueSmallChangePropertyId: R = "RangeValueSmallChange"
    Case UIA_RangeValueValuePropertyId: R = "RangeValueValue"
    Case UIA_RotationPropertyId: R = "Rotation"
    Case UIA_RuntimeIdPropertyId: R = "RuntimeId"
    Case UIA_ScrollHorizontallyScrollablePropertyId: R = "ScrollHorizontallyScrollable"
    Case UIA_ScrollHorizontalScrollPercentPropertyId: R = "ScrollHorizontalScrollPercent"
    Case UIA_ScrollHorizontalViewSizePropertyId: R = "ScrollHorizontalViewSize"
    Case UIA_ScrollVerticallyScrollablePropertyId: R = "ScrollVerticallyScrollable"
    Case UIA_ScrollVerticalScrollPercentPropertyId: R = "ScrollVerticalScrollPercent"
    Case UIA_ScrollVerticalViewSizePropertyId: R = "ScrollVerticalViewSize"
    Case UIA_Selection2FirstSelectedItemPropertyId: R = "Selection2FirstSelectedItem"
    Case UIA_Selection2ItemCountPropertyId: R = "Selection2ItemCount"
    Case UIA_Selection2LastSelectedItemPropertyId: R = "Selection2LastSelectedItem"
    Case UIA_SelectionIsSelectionRequiredPropertyId: R = "SelectionIsSelectionRequired"
    Case UIA_SelectionCanSelectMultiplePropertyId: R = "SelectionCanSelectMultiple"
    Case UIA_SelectionItemIsSelectedPropertyId: R = "SelectionItemIsSelected"
    Case UIA_SelectionItemSelectionContainerPropertyId: R = "SelectionItemSelectionContainer"
    Case UIA_SelectionSelectionPropertyId: R = "SelectionSelection"
    Case UIA_SizeOfSetPropertyId: R = "SizeOfSet"
    Case UIA_SizeOfSetPropertyId: R = "Size"
    Case UIA_SpreadsheetItemAnnotationObjectsPropertyId: R = "SpreadsheetItemAnnotationObjects"
    Case UIA_SpreadsheetItemAnnotationTypesPropertyId: R = "SpreadsheetItemAnnotationTypes"
    Case UIA_SpreadsheetItemFormulaPropertyId: R = "SpreadsheetItemFormula"
    Case UIA_StylesExtendedPropertiesPropertyId: R = "StylesExtendedProperties"
    Case UIA_StylesFillColorPropertyId: R = "StylesFillColor"
    Case UIA_StylesFillPatternColorPropertyId: R = "StylesFillPatternColor"
    Case UIA_StylesFillPatternStylePropertyId: R = "StylesFillPatternStyle"
    Case UIA_StylesShapePropertyId: R = "StylesShape"
    Case UIA_StylesStyleIdPropertyId: R = "StylesStyleId"
    Case UIA_StylesStyleNamePropertyId: R = "StylesStyleName"
    Case UIA_TableColumnHeadersPropertyId: R = "TableColumnHeaders"
    Case UIA_TableItemColumnHeaderItemsPropertyId: R = "TableItemColumnHeaderItems"
    Case UIA_TableItemRowHeaderItemsPropertyId: R = "TableItemRowHeaderItems"
    Case UIA_TableRowHeadersPropertyId: R = "TableRowHeaders"
    Case UIA_TableRowOrColumnMajorPropertyId: R = "TableRowOrColumnMajor"
    Case UIA_ToggleToggleStatePropertyId: R = "ToggleToggleState"
    Case UIA_Transform2CanZoomPropertyId: R = "Transform2CanZoom"
    Case UIA_Transform2ZoomLevelPropertyId: R = "Transform2ZoomLevel"
    Case UIA_Transform2ZoomMaximumPropertyId: R = "Transform2ZoomMaximumProperty"
    Case UIA_Transform2ZoomMinimumPropertyId: R = "Transform2ZoomMinimum"
    Case UIA_TransformCanMovePropertyId: R = "TransformCanMove"
    Case UIA_TransformCanResizePropertyId: R = "TransformCanResize"
    Case UIA_TransformCanRotatePropertyId: R = "TransformCanRotate"
    Case UIA_ValueValuePropertyId: R = "ValueValue"
    Case UIA_VisualEffectsPropertyId: R = "VisualEffects"
    Case UIA_WindowCanMaximizePropertyId: R = "WindowCanMaximize"
    Case UIA_WindowCanMinimizePropertyId: R = "WindowCanMinimize"
    Case UIA_WindowIsModalPropertyId: R = "WindowIsModal"
    Case UIA_WindowIsTopmostPropertyId: R = "WindowIsTopmost"
    Case UIA_WindowWindowInteractionStatePropertyId: R = "WindowWindowInteractionState"
    Case UIA_WindowWindowVisualStatePropertyId: R = "WindowWindowVisualState"
  End Select
  GetPropertyName = R
End Function

Public Function GetControlTypeName(ControlTypeID As UIAControlTypeIDs) As String
  Dim R As String
  Select Case ControlTypeID
    Case AppBar: R = "AppBar"
    Case Button: R = "Button"
    Case Calendar: R = "Calendar"
    Case CheckBox: R = "CheckBox"
    Case ComboBox: R = "ComboBox"
    Case ControlType: R = "ControlType"
    Case Custom: R = "Custom"
    Case DataGrid: R = "DataGrid"
    Case DataItem: R = "DataItem"
    Case Document: R = "Document"
    Case Edit: R = "Edit"
    Case Group: R = "Group"
    Case Header: R = "Header"
    Case HeaderItem: R = "HeaderItem"
    Case Hyperlink: R = "Hyperlink"
    Case Image: R = "Image"
    Case List: R = "List"
    Case ListItem: R = "ListItem"
    Case Menu: R = "Menu"
    Case MenuBar: R = "MenuBar"
    Case MenuItem: R = "MenuItem"
    Case Pane: R = "Pane"
    Case ProgressBar: R = "ProgressBar"
    Case RadioButton: R = "RadioButton"
    Case ScrollBar: R = "ScrollBar"
    Case SemanticZoom: R = "SemanticZoom"
    Case Separator: R = "Separator"
    Case Slider: R = "Slider"
    Case Spinner: R = "Spinner"
    Case SplitButton: R = "SplitButton"
    Case StatusBar: R = "StatusBar"
    Case TabControl: R = "Tab"
    Case TabItem: R = "TabItem"
    Case Table: R = "Table"
    Case Text: R = "Text"
    Case Thumb: R = "Thumb"
    Case TitleBar: R = "TitleBar"
    Case Toolbar: R = "ToolBar"
    Case ToolTip: R = "ToolTip"
    Case TreeControl: R = "Tree"
    Case TreeItem: R = "TreeItem"
    Case Window: R = "Window"
    Case Else
      MsgBox "ControlTypeID: " & ControlTypeID & " not handled in ControlTypeName"
  End Select
  GetControlTypeName = R
End Function

Public Function GetWindowInteractionStateName(WindowInteractionState As Long) As String
  Dim R As String
  Select Case WindowInteractionState 'UIAutomationClient.WindowInteractionState
    Case WindowInteractionState_Running: R = "Running"
    Case WindowInteractionState_Closing: R = "Closing"
    Case WindowInteractionState_ReadyForUserInteraction: R = "ReadyForUserInteraction"
    Case WindowInteractionState_BlockedByModalWindow: R = "BlockedByModalWindow"
    Case WindowInteractionState_NotResponding: R = "NotResponding"
    Case Else: R = "Unknown Window Interaction State (" & WindowInteractionState & ")"
  End Select
  GetWindowInteractionStateName = R
End Function

Public Function GetWindowVisualStateName(WindowVisualState As Long) As String
'https://learn.microsoft.com/en-us/windows/win32/api/uiautomationcore/ne-uiautomationcore-windowvisualstate
  Dim R As String
  Select Case WindowVisualState
    Case WindowVisualState_Maximized: R = "Maximized"
    Case WindowVisualState_Minimized: R = "Minimized"
    Case WindowVisualState_Normal: R = "Normal"
    Case Else: R = "Unknown Window Visual State (" & WindowVisualState & ")"
  End Select
  GetWindowVisualStateName = R
End Function

Public Function GetPropertyValueAsString(UIAElement As IUIAutomationElement, PropertyId As UIAProperties) As String
  Dim PropertyStringValue As String
'  PropertyStringValue = ""
  On Error Resume Next
  PropertyStringValue = VBA.Conversion.CStr(UIAElement.GetCurrentPropertyValue(PropertyId))
  On Error GoTo 0
  Select Case PropertyId
    Case UIAProperties.BoundingRectangle
      Dim Rect As pEssence.BoundingRectangle
      Set Rect = Factory.GetNewBoundingRectangle(UIAElement)
      PropertyStringValue = "Left:=" & Rect.Left & ", Top:=" & Rect.Top & ", Right:=" & Rect.Right & ", Bottom:=" & Rect.Bottom & " (" & Rect.Width & " x " & Rect.Height & ")"
    Case UIAProperties.ControlType
      PropertyStringValue = UIAProps.GetControlTypeName(CLng(PropertyStringValue))
'?? Show all values of just the set ones!?
'???    Case UIAProperties.GridRow
'      If PropertyStringValue = 0 Then: PropertyStringValue = ""
    Case UIAProperties.RuntimeId
      PropertyStringValue = pPath.RuntimeIDs.GetElementRuntimeId(UIAElement)
  End Select
  GetPropertyValueAsString = PropertyStringValue
End Function

