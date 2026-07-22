Attribute VB_Name = "OfficeRibbon"
'@Folder OfficeRibbon
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Private Rib As IRibbonUI

Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
  Set Rib = ribbon
  ActivateTab OfficeRibbonTabNames.Home
End Sub

'https://learn.microsoft.com/en-us/office/dev/add-ins/develop/built-in-ui-ids
Public Sub ActivateTab(TabName As String)
  If Not Rib Is Nothing Then
    Rib.ActivateTabMso TabName
  Else
    MsgBox "Ribbon reference not initialized!", vbExclamation
  End If
End Sub

Public Sub RunMacro1(control As IRibbonControl)
  ActivateTab OfficeRibbonTabNames.Home
  MsgBox "Macro 1 executed!", vbInformation
End Sub

Public Sub RunMacro2(control As IRibbonControl)
  ActivateTab OfficeRibbonTabNames.Developer
  MsgBox "Macro 2 executed!", vbInformation
End Sub



