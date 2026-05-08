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

Private Const ThisVBProjectName = "pEssence"

Private Sub Workbook_Open()
  'Microsoft Visual Basic for Applications Extensilbility 5.3 - needed for VBProject
  'This is not built in but needs adding manually for the References module
        
  'Microsoft Scripting Runtime - needed for Scripting.dictionary/File System Object
  pEssence.References.AddReferenceToWorkbookOrLibrary "C:\Windows\System32\scrrun.dll"
    
  'OLE Automation - needed for IUnknown type
  pEssence.References.AddReferenceToWorkbookOrLibrary "C:\Windows\System32\stdole2.tlb"
    
  'UIAutomationClient
  pEssence.References.AddReferenceToWorkbookOrLibrary "C:\Windows\System32\UIAutomationCore.dll"

  'MS Forms 2.0 Object Library - needed for JPK Treeview Control
  pEssence.References.AddReferenceToWorkbookOrLibrary "C:\Windows\System32\FM20.dll"
  
  Dim strPhosphorusWBFullName As String
  strPhosphorusWBFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "Phosphorus.xlam")
  pEssence.References.AddReferenceToWorkbookOrLibrary strPhosphorusWBFullName

  Dim strPhosphorusExpernalsFullName As String
  strPhosphorusExpernalsFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "Phosphorus Externals.xlam")
  pEssence.References.AddReferenceToWorkbookOrLibrary strPhosphorusExpernalsFullName

End Sub

Private Sub ExportPhosphorusSourceCode()
  ModuleManagement.ExportModulesWithFolders SubFolderForExport:="\src\pEssence"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

  If (VBA.Interaction.Environ$("COMPUTERNAME") = "LYNNSHPENVY") Or (VBA.Interaction.Environ$("COMPUTERNAME") = "ASPIRE16") Then
    ExportPhosphorusSourceCode
  End If
  
  pEssence.References.RemoveAllNonBuiltInReferencesFromAProject ThisVBProjectName
  
  'Always Save Code Changes on Closing Workbootk
  If (VBA.Interaction.Environ$("COMPUTERNAME") = "LYNNSHPENVY") Or (VBA.Interaction.Environ$("COMPUTERNAME") = "ASPIRE16") Then
    ThisWorkbook.Save
  End If

End Sub
