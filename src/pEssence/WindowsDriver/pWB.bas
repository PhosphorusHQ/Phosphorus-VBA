VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder WindowsDriver
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

'https://bettersolutions.com/vba/class-modules/implements.htm
'An interface contains only method & function signatures, not any properties

Public Sub Start(WebAppName As String, URL As String, WebAppPageTitle As String)
End Sub

Public Function GetRootWebArea() As pLocator
End Function

Public Function GetCurrentURL() As String
End Function

Public Sub NavigateBack()
End Sub

