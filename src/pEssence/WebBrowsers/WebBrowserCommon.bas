Attribute VB_Name = "WebBrowserCommon"
'@Folder WebBrowsers
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public Sub ClickEnabledItem(ClickItem As pElement, NonClickItem As pElement)
  If ClickItem.IsEnabled Then
    ClickItem.Click
  Else
    If Not NonClickItem Is Nothing Then
      NonClickItem.Click
    End If
  End If
End Sub
