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

Private Const ThisVBProjectName = "pWinDriver"

Private Sub Workbook_Open()

  'Trust Settings: Ensure "Trust access to the VBA project object model" is enabled in Excel's Trust Center (File > Options > Trust Center > Trust Center Settings > Macro Settings).

  'Microsoft Visual Basic for Applications Extensilbility 5.3 - needed for VBProject
  'This is not built in but needs adding manually for the References module
    
  'Microsoft Scripting Runtime - needed for Scripting.dictionary/File System Object
  pWinDriver.References.AddReferenceToWorkbookOrLibrary "C:\Windows\System32\scrrun.dll"
    
  'Windows Script Host Object Model - needed for wshShell
  pWinDriver.References.AddReferenceToWorkbookOrLibrary "C:\Windows\System32\wshom.ocx"

  'UIAutomationClient
  pWinDriver.References.AddReferenceToWorkbookOrLibrary "C:\Windows\System32\UIAutomationCore.dll"

  'OLE Automation - needed for IUnknown type
  pWinDriver.References.AddReferenceToWorkbookOrLibrary "C:\Windows\System32\stdole2.tlb"

  'Microsoft WinHTTP Services, v 5.1 - Http calls
  pWinDriver.References.AddReferenceToWorkbookOrLibrary "C:\WINDOWS\system32\winhttpcom.dll"

  Dim strPPathWBFullName As String
  strPPathWBFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "pPath.xlam")
  pWinDriver.References.AddReferenceToWorkbookOrLibrary strPPathWBFullName

  Dim strPhosphorusWBFullName As String
  strPhosphorusWBFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "Phosphorus.xlam")
  pWinDriver.References.AddReferenceToWorkbookOrLibrary strPhosphorusWBFullName

  Dim strPEssenceWBFullName As String
  strPEssenceWBFullName = VBA.Strings.Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "PhosphorEssence.xlam")
  pWinDriver.References.AddReferenceToWorkbookOrLibrary strPEssenceWBFullName

End Sub

Private Sub Test_ListAllReferencesInProject()
  pWinDriver.References.ListAllReferencesInAProject ThisVBProjectName
End Sub

Private Sub Test_RemoveAllNonBuiltInReferencesInProject()
  pWinDriver.References.RemoveAllNonBuiltInReferencesFromAProject ThisVBProjectName
End Sub

'Always Save Code Changes on Closing Workbootk
Private Sub Workbook_BeforeClose(Cancel As Boolean)

  If (VBA.Interaction.Environ$("COMPUTERNAME") = "LYNNSHPENVY") Or (VBA.Interaction.Environ$("COMPUTERNAME") = "ASPIRE16") Then
    ExportPhosphorusSourceCode
  End If
  
  pWinDriver.References.RemoveAllNonBuiltInReferencesFromAProject ThisVBProjectName
  
  If (VBA.Interaction.Environ$("COMPUTERNAME") = "LYNNSHPENVY") Or (VBA.Interaction.Environ$("COMPUTERNAME") = "ASPIRE16") Then
    ThisWorkbook.Save
  End If

End Sub

Private Sub SetModulesToKeep()
  Dim strArray(0) As String
  strArray(0) = "ModuleManagement"
  Phosphorus.ModuleManagement.SetModulesToKeep strArray()
End Sub

Private Sub ExportPhosphorusSourceCode()
  Phosphorus.ModuleManagement.ExportModulesWithFolders SubFolderForExport:="\src\pWindowsDriver", projectName:=ThisVBProjectName
End Sub

Private Sub RemoveAllPhosphorusSourceCode()
  SetModulesToKeep
  Phosphorus.ModuleManagement.RemoveAllComponentsExcept ThisWorkbook.VBProject
End Sub

Private Sub ImportAllPhosphorusSourceCode()
  SetModulesToKeep
  Phosphorus.ModuleManagement.ImportModulesFromFolder "src\pWindowsDriver", ""
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
