Attribute VB_Name = "pUnitErrorStatic"
'@Folder pUnit
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

'Custom error onject - callign Applicaiton.Run xxxx does not pass back errors as normally expected
Public gpUnitError As Phosphorus.pUnitError

Public Function TrapError(Number As Long, Description As String)
 Set gpUnitError = New pUnitError
 gpUnitError.Number = Number
 gpUnitError.Description = Description
End Function

Public Function ClearAllErrors()
  Set gpUnitError = Nothing
End Function
