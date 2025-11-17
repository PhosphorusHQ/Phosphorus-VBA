Attribute VB_Name = "AssertionsStatic"
'@Folder Assertions
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public pAssert As Assertions

'TODO: Can't make this available until we remove Rubberduck
'Public Assert As Assertions
'
'Public Sub GetAssert()
'  If Assert Is Nothing Then
'    Set Assert = New Phosphorus.Assertions
'  End If
'End Sub

Public Sub GetAssert()
  If pAssert Is Nothing Then
    Set pAssert = New Phosphorus.Assertions
  End If
End Sub

Public Sub CloseAssert()
  Set pAssert = Nothing
End Sub
