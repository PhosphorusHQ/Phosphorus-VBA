VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder pEssence
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public RootDesktopUIAElement As IUIAutomationElement

Public Function GetRootDesktopElement() As IUIAutomationElement
  If RootDesktopUIAElement Is Nothing Then
    Set RootDesktopUIAElement = UIA.GetRootElement
  Else
    Set GetRootDesktopElement = RootDesktopUIAElement
  End If
End Function

Public Function GetNewSearch() As pSearch
  Dim Element As pSearch
  Set Element = New pSearch
  Set GetNewSearch = Element
End Function

Public Function GetNewEdgeWebBrowser() As EdgeWebBrowser
  Dim Edge As EdgeWebBrowser
  Set Edge = New EdgeWebBrowser
  Set GetNewEdgeWebBrowser = Edge
End Function


