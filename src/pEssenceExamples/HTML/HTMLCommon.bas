Attribute VB_Name = "HTMLCommon"
'@Folder HTML
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public Function GetULREncodedFilepath(SourceFilePath As String) As String
  SourceFilePath = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, SourceFilePath)
  SourceFilePath = VBA.Strings.Replace(SourceFilePath, "\", "/")
  SourceFilePath = VBA.Strings.Replace(SourceFilePath, " ", "%20")
  SourceFilePath = "file:///" & SourceFilePath
  GetULREncodedFilepath = SourceFilePath
End Function
