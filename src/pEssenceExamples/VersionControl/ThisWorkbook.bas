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
  strPEssenceWBFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "PhosphorEssence.xlam")
  pEssenceExamples.References.AddReferenceToWorkbookOrLibrary strPEssenceWBFullName
  
  'UIAutomationClient
  pEssenceExamples.References.AddReferenceToWorkbookOrLibrary "C:\Windows\System32\UIAutomationCore.dll"

  'Microsoft Scripting Runtime - needed for Scripting.dictionary/File System Object
  pEssenceExamples.References.AddReferenceToWorkbookOrLibrary "C:\Windows\System32\scrrun.dll"

End Sub

Private Sub ExportPhosphorusSourceCode()
  ModuleManagement.ExportModulesWithFolders SubFolderForExport:="\src\pEssenceExamples"
End Sub

'Always Save Code Changes on Closing Workbootk
Private Sub Workbook_BeforeClose(Cancel As Boolean)
  If (VBA.Interaction.Environ$("COMPUTERNAME") = "LYNNSHPENVY") Or (VBA.Interaction.Environ$("COMPUTERNAME") = "ASPIRE16") Then
    ExportPhosphorusSourceCode
  End If
  pEssenceExamples.References.RemoveAllNonBuiltInReferencesFromAProject ThisVBProjectName
  'Always Save Code Changes on Closing Workbootk
  If (VBA.Interaction.Environ$("COMPUTERNAME") = "LYNNSHPENVY") Or (VBA.Interaction.Environ$("COMPUTERNAME") = "ASPIRE16") Then
    ThisWorkbook.Save
  End If
End Sub
