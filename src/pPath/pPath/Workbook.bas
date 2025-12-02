Attribute VB_Name = "Workbook"
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

Public gPPathWB As Excel.Workbook

Public Sub OpenWB()
  If gPPathWB Is Nothing Then
    Set gPPathWB = Excel.Workbooks.Add
    'gPPathWB.Windows(1).Visible = False : this seems to make the tests fail??
  End If
  Excel.Application.ScreenUpdating = False
End Sub

Public Sub CloseWB()
  'Automatically closed on Module Cleanup?
  On Error Resume Next
  gPPathWB.Close savechanges:=False
  Set gPPathWB = Nothing
  On Error GoTo 0
End Sub

