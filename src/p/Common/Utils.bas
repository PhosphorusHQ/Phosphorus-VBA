Attribute VB_Name = "Utils"
'@Folder Common
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public Function GetSizeOfArray(arr As Variant) As Long
  GetSizeOfArray = 0
  On Error Resume Next
  GetSizeOfArray = UBound(arr)
  On Error GoTo 0
End Function

Public Sub CloseAllOtherWorkbooks()
 Dim wb As Workbook
 For Each wb In Application.Workbooks
   If wb.Name <> ThisWorkbook.Name Then
     wb.Close savechanges:=False
   End If
 Next
End Sub
