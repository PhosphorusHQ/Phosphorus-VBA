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
  clickablePoint = UIA_ClickablePointPropertyId
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
  IsSelectionPattern2Available = UIA_IsSelectionPattern2AvailablePropertyId
  IsSelectionPatternAvailable = UIA_IsSelectionPatternAvailablePropertyId
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
  Running = WindowInteractionState_Running
End Enum

Public Function ControlTypeName(ControlTypeID As UIAControlTypeIDs) As String
  Dim R As String
  Select Case ControlTypeID
    Case AppBar: R = "AppBar"
    Case Button: R = "Button"
    Case Calendar: R = "Calendar"
    Case CheckBox: R = "CheckBox"
    Case ComboBox: R = "ComboBox"
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
  ControlTypeName = R
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
