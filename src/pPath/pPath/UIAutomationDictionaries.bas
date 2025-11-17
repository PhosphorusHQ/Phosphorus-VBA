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
  InitialiseNavigablePropertyIDsDictionary
End Sub

Private Sub InitialiseControlTypeIDsDictionary()
  'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controltype-ids
  'Set ControlTypeIDs = CreateObject("Scripting.Dictionary")
  Set ControlTypeIDs = New Scripting.dictionary
  ControlTypeIDs.Add 50040, "AppBar"
  ControlTypeIDs.Add 50000, "Button"
  ControlTypeIDs.Add 50001, "Calendar"
  ControlTypeIDs.Add 50002, "CheckBox"
  ControlTypeIDs.Add 50003, "ComboBox"
  ControlTypeIDs.Add 50025, "Custom"
  ControlTypeIDs.Add 50028, "DataGrid"
  ControlTypeIDs.Add 50029, "DataItem"
  ControlTypeIDs.Add 50030, "Document"
  ControlTypeIDs.Add 50004, "Edit"
  ControlTypeIDs.Add 50026, "Group"
  ControlTypeIDs.Add 50034, "Header"
  ControlTypeIDs.Add 50035, "HeaderItem"
  ControlTypeIDs.Add 50005, "Hyperlink"
  ControlTypeIDs.Add 50006, "Image"
  ControlTypeIDs.Add 50008, "List"
  ControlTypeIDs.Add 50007, "ListItem"
  ControlTypeIDs.Add 50010, "MenuBar"
  ControlTypeIDs.Add 50009, "Menu"
  ControlTypeIDs.Add 50011, "MenuItem"
  ControlTypeIDs.Add 50033, "Pane"
  ControlTypeIDs.Add 50012, "ProgressBar"
  ControlTypeIDs.Add 50013, "RadioButton"
  ControlTypeIDs.Add 50014, "ScrollBar"
  ControlTypeIDs.Add 50039, "SemanticZoom"
  ControlTypeIDs.Add 50038, "Separator"
  ControlTypeIDs.Add 50015, "Slider"
  ControlTypeIDs.Add 50016, "Spinner"
  ControlTypeIDs.Add 50031, "SplitButton"
  ControlTypeIDs.Add 50017, "StatusBar"
  ControlTypeIDs.Add 50018, "Tab"
  ControlTypeIDs.Add 50019, "TabItem"
  ControlTypeIDs.Add 50036, "Table"
  ControlTypeIDs.Add 50020, "Text"
  ControlTypeIDs.Add 50027, "Thumb"
  ControlTypeIDs.Add 50037, "TitleBar"
  ControlTypeIDs.Add 50021, "ToolBar"
  ControlTypeIDs.Add 50022, "ToolTip"
  ControlTypeIDs.Add 50023, "Tree"
  ControlTypeIDs.Add 50024, "TreeItem"
  ControlTypeIDs.Add 50032, "Window"
End Sub

'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-automation-element-propids

Private Sub InitialiseAllPropertyIDsDictionary()
  Set AllPropertyIDs = CreateObject("Scripting.Dictionary")
  Set AllPropertyIDs = New Scripting.dictionary
  'TODO how to deal with :
  '  UIA_AnnotationObjectsPropertyId, UIA_ControllerForPropertyId, UIA_DescribedByPropertyId,
  '  UIA_FlowsFromPropertyId, UIA_FlowsToPropertyId, UIA_LabeledByPropertyId
  '? - These cause an error!

  AllPropertyIDs.Add 30006, "AcceleratorKey"
  AllPropertyIDs.Add 30007, "AccessKey"
  AllPropertyIDs.Add 30156, "AnnotationObjects"
  AllPropertyIDs.Add 30155, "AnnotationTypes"
  AllPropertyIDs.Add 30102, "AriaProperties"
  AllPropertyIDs.Add 30101, "AriaRole" 'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-ariaspecification
  AllPropertyIDs.Add 30011, "AutomationId"
  AllPropertyIDs.Add 30001, "BoundingRectangle"
  AllPropertyIDs.Add 30165, "CenterPoint"
  AllPropertyIDs.Add 30012, "ClassName"
  AllPropertyIDs.Add 30014, "ClickablePoint"
  AllPropertyIDs.Add 30104, "ControllerFor"
  AllPropertyIDs.Add 30003, "ControlType"
  AllPropertyIDs.Add 30015, "Culture"
  AllPropertyIDs.Add 30105, "DescribedBy"
  AllPropertyIDs.Add 30160, "FillColor"
  AllPropertyIDs.Add 30162, "FillType"
  AllPropertyIDs.Add 30148, "FlowsFrom"
  AllPropertyIDs.Add 30106, "FlowsTo"
  AllPropertyIDs.Add 30024, "FrameworkId"
  AllPropertyIDs.Add 30159, "FullDescription"
  AllPropertyIDs.Add 30008, "HasKeyboardFocus"
  AllPropertyIDs.Add 30173, "HeadingLevel"
  AllPropertyIDs.Add 30013, "HelpText"
  AllPropertyIDs.Add 30017, "IsContentElement"
  AllPropertyIDs.Add 30016, "IsControlElement"
  AllPropertyIDs.Add 30103, "IsDataValidForForm"
  AllPropertyIDs.Add 30174, "IsDialog"
  AllPropertyIDs.Add 30010, "IsEnabled"
  AllPropertyIDs.Add 30009, "IsKeyboardFocusable"
  AllPropertyIDs.Add 30022, "IsOffscreen"
  AllPropertyIDs.Add 30019, "IsPassword"
  AllPropertyIDs.Add 30150, "IsPeripheral"
  AllPropertyIDs.Add 30025, "IsRequiredForForm"
  AllPropertyIDs.Add 30026, "ItemStatus"
  AllPropertyIDs.Add 30021, "ItemType"
  AllPropertyIDs.Add 30018, "LabeledBy"
  AllPropertyIDs.Add 30157, "LandmarkType"
  AllPropertyIDs.Add 30154, "Level"
  AllPropertyIDs.Add 30135, "LiveSetting"
  AllPropertyIDs.Add 30004, "LocalizedControlType"
  AllPropertyIDs.Add 30158, "LocalizedLandmarkType"
  AllPropertyIDs.Add 30005, "Name"
  AllPropertyIDs.Add 30020, "NativeWindowHandle"
  AllPropertyIDs.Add 30111, "OptimizeForVisualContent"
  AllPropertyIDs.Add 30023, "Orientation"
  AllPropertyIDs.Add 30161, "OutlineColor"
  AllPropertyIDs.Add 30164, "OutlineThickness"
  AllPropertyIDs.Add 30152, "PositionInSet"
  AllPropertyIDs.Add 30002, "ProcessId"
  AllPropertyIDs.Add 30107, "ProviderDescription"
  AllPropertyIDs.Add 30166, "Rotation"
  AllPropertyIDs.Add 30000, "RuntimeId"
  AllPropertyIDs.Add 30167, "Size"
  AllPropertyIDs.Add 30153, "SizeOfSet"
  AllPropertyIDs.Add 30163, "VisualEffects"
End Sub

Private Sub InitialiseNavigablePropertyIDsDictionary()
  Set NavigablePropertyIDs = CreateObject("Scripting.Dictionary")
  Set NavigablePropertyIDs = New Scripting.dictionary
  'TODO how to deal with :
  '  UIA_AnnotationObjectsPropertyId, UIA_ControllerForPropertyId, UIA_DescribedByPropertyId,
  '  UIA_FlowsFromPropertyId, UIA_FlowsToPropertyId, UIA_LabeledByPropertyId
  '? - These cause an error!
  
  NavigablePropertyIDs.Add 30006, "AcceleratorKey"
  NavigablePropertyIDs.Add 30007, "AccessKey"
'  PropertyIDs.Add 30156, "AnnotationObjects"
'  NavigablePropertyIDs.Add 30155, "AnnotationTypes"
  NavigablePropertyIDs.Add 30102, "AriaProperties"
  NavigablePropertyIDs.Add 30101, "AriaRole" 'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-ariaspecification
  NavigablePropertyIDs.Add 30011, "AutomationId"
  NavigablePropertyIDs.Add 30001, "BoundingRectangle"
  NavigablePropertyIDs.Add 30165, "CenterPoint"
  NavigablePropertyIDs.Add 30012, "ClassName"
  NavigablePropertyIDs.Add 30014, "ClickablePoint"
'  PropertyIDs.Add 30104, "ControllerFor"
  'We don't need the ControlType as the a property type!
  'PropertyIDs.Add 30003, "ControlType"
  NavigablePropertyIDs.Add 30015, "Culture"
'  PropertyIDs.Add 30105, "DescribedBy"
  NavigablePropertyIDs.Add 30160, "FillColor"
  NavigablePropertyIDs.Add 30162, "FillType"
'  PropertyIDs.Add 30148, "FlowsFrom"
'  PropertyIDs.Add 30106, "FlowsTo"
  NavigablePropertyIDs.Add 30024, "FrameworkId"
  NavigablePropertyIDs.Add 30159, "FullDescription"
  NavigablePropertyIDs.Add 30008, "HasKeyboardFocus"
  NavigablePropertyIDs.Add 30173, "HeadingLevel"
  NavigablePropertyIDs.Add 30013, "HelpText"
  NavigablePropertyIDs.Add 30017, "IsContentElement"
  NavigablePropertyIDs.Add 30016, "IsControlElement"
  NavigablePropertyIDs.Add 30103, "IsDataValidForForm"
  NavigablePropertyIDs.Add 30174, "IsDialog"
  NavigablePropertyIDs.Add 30010, "IsEnabled"
  NavigablePropertyIDs.Add 30009, "IsKeyboardFocusable"
  NavigablePropertyIDs.Add 30022, "IsOffscreen"
  NavigablePropertyIDs.Add 30019, "IsPassword"
  NavigablePropertyIDs.Add 30150, "IsPeripheral"
  NavigablePropertyIDs.Add 30025, "IsRequiredForForm"
  NavigablePropertyIDs.Add 30026, "ItemStatus"
  NavigablePropertyIDs.Add 30021, "ItemType"
'  PropertyIDs.Add 30018, "LabeledBy"
  NavigablePropertyIDs.Add 30157, "LandmarkType"
  NavigablePropertyIDs.Add 30154, "Level"
  NavigablePropertyIDs.Add 30135, "LiveSetting"
  NavigablePropertyIDs.Add 30004, "LocalizedControlType"
  NavigablePropertyIDs.Add 30158, "LocalizedLandmarkType"
  NavigablePropertyIDs.Add 30005, "Name"
  NavigablePropertyIDs.Add 30020, "NativeWindowHandle"
  NavigablePropertyIDs.Add 30111, "OptimizeForVisualContent"
  NavigablePropertyIDs.Add 30023, "Orientation"
  'NavigablePropertyIDs.Add 30161, "OutlineColor"
  NavigablePropertyIDs.Add 30164, "OutlineThickness"
  NavigablePropertyIDs.Add 30152, "PositionInSet"
  NavigablePropertyIDs.Add 30002, "ProcessId"
  NavigablePropertyIDs.Add 30107, "ProviderDescription"
  NavigablePropertyIDs.Add 30166, "Rotation"
'Don't include RuntimeID in exposed properties
'  PropertyIDs.Add 30000, "RuntimeId"
  NavigablePropertyIDs.Add 30167, "Size"
  NavigablePropertyIDs.Add 30153, "SizeOfSet"
  NavigablePropertyIDs.Add 30163, "VisualEffects"
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

