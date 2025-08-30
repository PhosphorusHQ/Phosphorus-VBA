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
Option Explicit

Private Sub Workbook_Open()
  Dim strPhosphorusWBFullName As String
  strPhosphorusWBFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "Phosphorus.xlam")
  pUnitTests.References.AddReferenceToWorkbook strPhosphorusWBFullName
End Sub

'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'  UnitTestingExternalProject.References.RemoveAllAddedReferences
'End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  pUnitTests.References.RemoveAllAddedReferences
  'Always Save Code Changes on Closing Workbootk
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
  Phosphorus.ModuleManagement.ExportModulesWithFolders SubFolderForExport:="\srcPUnitTests", projectName:="pUnitTests"
End Sub

Private Sub RemoveAllPhosphorusSourceCode()
'  SetModulesToKeep
  Phosphorus.ModuleManagement.RemoveAllComponentsExcept ThisWorkbook.VBProject
End Sub

Private Sub ImportAllPhosphorusSourceCode()
'  SetModulesToKeep
  Phosphorus.ModuleManagement.ImportModulesFromFolder "srcPUnitTests", ""
End Sub

