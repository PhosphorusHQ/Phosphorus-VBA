Attribute VB_Name = "Factory"
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

Public Function GetNewLogger() As Phosphorus.Log4P
  Set GetNewLogger = New Phosphorus.Log4P
End Function

Public Function GetNewPhosphorusLog4P() As Phosphorus.Log4P
  Set GetNewPhosphorusLog4P = New Phosphorus.Log4P
End Function
