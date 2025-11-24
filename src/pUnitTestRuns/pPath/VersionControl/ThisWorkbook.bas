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

  'This creates a dynamic reference to the Phosphorus project
  'End user's will probably hardcode this reference
  Dim strPhosphorusWBFullName As String
  strPhosphorusWBFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "Phosphorus.xlam")
  pUnitTestRuns_pPathTests.References.AddReferenceToWorkbook strPhosphorusWBFullName
  
  'This creates a dynamic reference to the pUnit project
  'End user's will probably hardcode this reference
  Dim strPUnitWBFullName As String
  strPUnitWBFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "pUnit.xlam")
  pUnitTestRuns_pPathTests.References.AddReferenceToWorkbook strPUnitWBFullName
  
  'This creates a dynamic reference in pUnit to the target test project
  'End user's will need to replicate this dynamic reference in their test run modules
  Dim strTestWorkbookFilepath As String
  strTestWorkbookFilepath = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "pUnit Tests - pPath.xlam")
  pUnitTestRuns_pPathTests.References.AddReferenceToWorkbook strTestWorkbookFilepath, "pUnit"

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  pUnitTestRuns_pPathTests.References.RemoveAllAddedReferences
  'Always Save Code Changes on Closing Workbootk
  If VBA.Interaction.Environ$("COMPUTERNAME") = "LYNNSHPENVY" Then
    ExportPhosphorusSourceCode
    ThisWorkbook.Save
  End If
End Sub

Private Sub ExportPhosphorusSourceCode()
  Phosphorus.ModuleManagement.ExportModulesWithFolders SubFolderForExport:="\src\pUnitTestRuns\pPath", projectName:="pUnitTestRuns_pPathTests"
End Sub

Private Sub RemoveAllPhosphorusSourceCode()
  Phosphorus.ModuleManagement.RemoveAllComponentsExcept ThisWorkbook.VBProject
End Sub

Private Sub ImportAllPhosphorusSourceCode()
  Phosphorus.ModuleManagement.ImportModulesFromFolder "srcPUnitTestRuns", ""
End Sub

