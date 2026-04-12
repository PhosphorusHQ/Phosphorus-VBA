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

Public Enum UIAPatterns
  InvokePattern = UIA_PatternIds.UIA_InvokePatternId
  LegacyIAccessiblePattern = UIA_LegacyIAccessiblePatternId
  ScrollItemPattern = UIA_PatternIds.UIA_ScrollItemPatternId
  SelectionItemPattern = UIA_PatternIds.UIA_SelectionItemPatternId
  TogglePattern = UIA_PatternIds.UIA_TogglePatternId
  ValuePattern = UIA_PatternIds.UIA_ValuePatternId
End Enum

