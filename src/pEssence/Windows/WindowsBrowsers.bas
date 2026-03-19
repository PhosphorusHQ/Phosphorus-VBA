Attribute VB_Name = "WindowsBrowsers"
'@Folder Windows
Option Explicit
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================Option Explicit

Public Sub StartEdge(WebAppName As String, URL As String)
  LaunchCommandByProtocol WebAppName, "microsoft-edge:", URL, WindowStyle.Maximized
End Sub
