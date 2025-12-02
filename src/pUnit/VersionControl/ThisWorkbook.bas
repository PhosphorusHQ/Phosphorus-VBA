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

Private Const ThisVBProjectName = "pUnit"

Private Sub Workbook_Open()

  'Trust Settings: Ensure "Trust access to the VBA project object model" is enabled in Excel's Trust Center (File > Options > Trust Center > Trust Center Settings > Macro Settings).

  'Microsoft Visual Basic for Applications Extensilbility 5.3 - needed for VBProject
  'This is not built in but needs adding manually for the References module
  
  'Microsoft Scripting Runtime - needed for Scripting.dictionary
  pUnit.References.AddReferenceToWorkbookOrLibrary "C:\Windows\System32\scrrun.dll"
    
  'Windows Script Host Object Model - needed for wshShell
  pUnit.References.AddReferenceToWorkbookOrLibrary "C:\Windows\System32\wshom.ocx"

  Dim strPhosphorusWBFullName As String
  strPhosphorusWBFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "Phosphorus.xlam")
  pUnit.References.AddReferenceToWorkbookOrLibrary strPhosphorusWBFullName

End Sub

Private Sub Test_ListAllReferencesInProject()
  pUnit.References.ListAllReferencesInAProject ThisVBProjectName
End Sub

Private Sub Test_RemoveAllNonBuiltInReferencesInProject()
  pUnit.References.RemoveAllNonBuiltInReferencesFromAProject ThisVBProjectName
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  
  If VBA.Interaction.Environ$("COMPUTERNAME") = "LYNNSHPENVY" Then
    ExportPhosphorusSourceCode
  End If
  
  pUnit.References.RemoveAllNonBuiltInReferencesFromAProject ThisVBProjectName
  
  'Always Save Code Changes on Closing Workbook
  If VBA.Interaction.Environ$("COMPUTERNAME") = "LYNNSHPENVY" Then
    ThisWorkbook.Save
  End If

End Sub

'Private Sub SetModulesToKeep()
'  Dim strArray(0) As String
'  strArray(0) = "ModuleManagement"
'  ModuleManagement.SetModulesToKeep strArray()
'End Sub

Private Sub ExportPhosphorusSourceCode()
  Phosphorus.ModuleManagement.ExportModulesWithFolders SubFolderForExport:="\src\pUnit", projectName:=ThisVBProjectName
End Sub

Private Sub RemoveAllPhosphorusSourceCode()
'  SetModulesToKeep
  Phosphorus.ModuleManagement.RemoveAllComponentsExcept ThisWorkbook.VBProject
End Sub

Private Sub ImportAllPhosphorusSourceCode()
'  SetModulesToKeep
  Phosphorus.ModuleManagement.ImportModulesFromFolder "src\pUnit", ""
End Sub

