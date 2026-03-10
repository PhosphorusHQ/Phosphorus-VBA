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

Private Const ThisVBProjectName = "pEssenceExamples"

Private Sub Workbook_Open()
  Dim strPEssenceWBFullName As String
  strPEssenceWBFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "pEssence.xlam")
  pEssenceExamples.References.AddReferenceToWorkbookOrLibrary strPEssenceWBFullName
End Sub

Private Sub ExportPhosphorusSourceCode()
  ModuleManagement.ExportModulesWithFolders SubFolderForExport:="\src\pEssenceExamples"
End Sub

'Always Save Code Changes on Closing Workbootk
Private Sub Workbook_BeforeClose(Cancel As Boolean)
  If VBA.Interaction.Environ$("COMPUTERNAME") = "LYNNSHPENVY" Then
    ExportPhosphorusSourceCode
  End If
  pEssenceExamples.References.RemoveAllNonBuiltInReferencesFromAProject ThisVBProjectName
End Sub
