VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder VersionControl
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Private Sub Workbook_Open()
  
  Dim strPhosphorusWBFullName As String
  strPhosphorusWBFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "Phosphorus.xlam")
  pUnitTests_pUnit.References.AddReferenceToWorkbook strPhosphorusWBFullName
  
End Sub

'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'  UnitTestingExternalProject.References.RemoveAllAddedReferences
'End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  pUnitTests_pUnit.References.RemoveAllAddedReferences
  'Always Save Code Changes on Closing Workbootk
  If VBA.Interaction.Environ$("COMPUTERNAME") = "LYNNSHPENVY" Then
    ExportPhosphorusSourceCode
    ThisWorkbook.Save
  End If
End Sub

'Private Sub SetModulesToKeep()
'  Dim strArray(0) As String
'  strArray(0) = "ModuleManagement"
'  ModuleManagement.SetModulesToKeep strArray()
'End Sub

Private Sub ExportPhosphorusSourceCode()
  Phosphorus.ModuleManagement.ExportModulesWithFolders SubFolderForExport:="\src\pUnitTests\pUnit", projectName:="pUnitTests_pUnit"
End Sub

Private Sub RemoveAllPhosphorusSourceCode()
'  SetModulesToKeep
  Phosphorus.ModuleManagement.RemoveAllComponentsExcept ThisWorkbook.VBProject
End Sub

Private Sub ImportAllPhosphorusSourceCode()
'  SetModulesToKeep
  Phosphorus.ModuleManagement.ImportModulesFromFolder "src\pUnitTests\pUnit", ""
End Sub

