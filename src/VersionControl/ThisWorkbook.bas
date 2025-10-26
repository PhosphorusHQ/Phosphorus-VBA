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

'Always Save Code Changes on Closing Workbootk
Private Sub Workbook_BeforeClose(Cancel As Boolean)
  If VBA.Interaction.Environ$("COMPUTERNAME") = "LYNNSHPENVY" Then
    ThisWorkbook.Save
  End If
End Sub

Private Sub SetModulesToKeep()
  Dim strArray(0) As String
  strArray(0) = "ModuleManagement"
  ModuleManagement.SetModulesToKeep strArray()
End Sub

Private Sub ExportPhosphorusSourceCode()
  ModuleManagement.ExportModulesWithFolders SubFolderForExport:="\src"
End Sub

Private Sub RemoveAllPhosphorusSourceCode()
  SetModulesToKeep
  ModuleManagement.RemoveAllComponentsExcept ThisWorkbook.VBProject
End Sub

Private Sub ImportAllPhosphorusSourceCode()
  SetModulesToKeep
  ModuleManagement.ImportModulesFromFolder "src", ""
End Sub

Private Sub Test2()
'  ' Export from specific project
'  ExportModulesWithFolders ProjectName:="MyProjectName"
End Sub

Private Sub Test3()
'  ' Import to current project
'  ImportModulesFromFolder
End Sub

Private Sub Test4()
'  ' Import to specific project
'  ImportModulesFromFolder "MyProjectName"
End Sub
