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

Private Const ThisVBProjectName = "pUnitTests_pUnit"

Private Sub Workbook_Open()
  
  'Trust Settings: Ensure "Trust access to the VBA project object model" is enabled in Excel's Trust Center (File > Options > Trust Center > Trust Center Settings > Macro Settings).

  'Microsoft Visual Basic for Applications Extensilbility 5.3 - needed for VBProject
  'This is not built in but needs adding manually for the References module
  
  Dim strPhosphorusWBFullName As String
  strPhosphorusWBFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "Phosphorus.xlam")
  pUnitTests_pUnit.References.AddReferenceToWorkbookOrLibrary strPhosphorusWBFullName
  
End Sub

Private Sub Test_ListAllReferencesInProject()
  pUnitTests_pUnit.References.ListAllReferencesInAProject ThisVBProjectName
End Sub

Private Sub Test_RemoveAllNonBuiltInReferencesInProject()
  pUnitTests_pUnit.References.RemoveAllNonBuiltInReferencesFromAProject ThisVBProjectName
End Sub

'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'  UnitTestingExternalProject.References.RemoveAllAddedReferences
'End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

  If (VBA.Interaction.Environ$("COMPUTERNAME") = "LYNNSHPENVY") Or (VBA.Interaction.Environ$("COMPUTERNAME") = "ASPIRE16") Then
    ExportPhosphorusSourceCode
  End If
  
  pUnitTests_pUnit.References.RemoveAllNonBuiltInReferencesFromAProject ThisVBProjectName
  
  'Always Save Code Changes on Closing Workbootk
  If (VBA.Interaction.Environ$("COMPUTERNAME") = "LYNNSHPENVY") Or (VBA.Interaction.Environ$("COMPUTERNAME") = "ASPIRE16") Then
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

