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

Private Const ThisVBProjectName = "pUnitTestRuns_pUnit"

Private Sub Workbook_Open()

  'Trust Settings: Ensure "Trust access to the VBA project object model" is enabled in Excel's Trust Center (File > Options > Trust Center > Trust Center Settings > Macro Settings).

  'Microsoft Visual Basic for Applications Extensilbility 5.3 - needed for VBProject
  'This is not built in but needs adding manually for the References module
    
  'Windows Script Host Object Model - needed for wshShell
  pUnitTestRuns_pUnit.References.AddReferenceToWorkbookOrLibrary "C:\Windows\System32\wshom.ocx"

  'This creates a dynamic reference to the Phosphorus project
  'End user's will probably hardcode this reference
  Dim strPhosphorusWBFullName As String
  strPhosphorusWBFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "Phosphorus.xlam")
  pUnitTestRuns_pUnit.References.AddReferenceToWorkbookOrLibrary strPhosphorusWBFullName
  
  'This creates a dynamic reference to the pUnit project
  'End user's will probably hardcode this reference
  Dim strPUnitWBFullName As String
  strPUnitWBFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "pUnit.xlam")
  pUnitTestRuns_pUnit.References.AddReferenceToWorkbookOrLibrary strPUnitWBFullName
  
  'This creates a dynamic reference in pUnit to the target test project
  'End user's will need to replicate this dynamic reference in their test run modules
  Dim strTestWorkbookFilepath As String
  strTestWorkbookFilepath = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "pUnit Tests - pUnit.xlam")
  pUnitTestRuns_pUnit.References.AddReferenceToWorkbookOrLibrary strTestWorkbookFilepath, "pUnit"

End Sub

Private Sub Test_ListAllReferencesInProject()
  pUnitTestRuns_pUnit.References.ListAllReferencesInAProject ThisVBProjectName
End Sub

Private Sub Test_RemoveAllNonBuiltInReferencesInProject()
  pUnitTestRuns_pUnit.References.RemoveAllNonBuiltInReferencesFromAProject ThisVBProjectName
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

  'Always Save Code Changes on Closing Workbootk
  ExportPhosphorusSourceCode
  pUnitTestRuns_pUnit.References.RemoveAllNonBuiltInReferencesFromAProject ThisVBProjectName
  If Not ThisWorkbook.ReadOnly Then
    ThisWorkbook.Save
  End If

End Sub

Private Sub ExportPhosphorusSourceCode()
  On Error Resume Next
  VBA.FileSystem.MkDir ThisWorkbook.Path & "\src"
  VBA.FileSystem.MkDir ThisWorkbook.Path & "\src\pUnitTestRuns"
  VBA.FileSystem.MkDir ThisWorkbook.Path & "\src\pUnitTestRuns\pUnit"
  On Error GoTo 0
  Phosphorus.ModuleManagement.ExportModulesWithFolders SubFolderForExport:="\src\pUnitTestRuns\pUnit", projectName:="pUnitTestRuns_pUnit"
End Sub

Private Sub RemoveAllPhosphorusSourceCode()
  Phosphorus.ModuleManagement.RemoveAllComponentsExcept ThisWorkbook.VBProject
End Sub

Private Sub ImportAllPhosphorusSourceCode()
  Phosphorus.ModuleManagement.ImportModulesFromFolder "srcPUnitTestRuns", ""
End Sub

