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

Private Const ThisVBProjectName = "pExternals"

Private Sub Workbook_Open()

  'Microsoft Visual Basic for Applications Extensilbility 5.3 - needed for VBProject
  'This is not built in but needs adding manually for the References module

  'MS Forms 2.0 Object Library - needed for JPK Treeview Control
  'OLE Automation - needed for JPK Treeview Control - STDPicture

End Sub

Private Sub ExportPhosphorusSourceCode()
  On Error Resume Next
  VBA.FileSystem.MkDir ThisWorkbook.Path & "\src"
  VBA.FileSystem.MkDir ThisWorkbook.Path & "\src\pExternals"
  On Error GoTo 0
  ModuleManagement.ExportModulesWithFolders SubFolderForExport:="\src\pExternals"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

  'Always Save Code Changes on Closing Workbook
  ExportPhosphorusSourceCode
  If Not ThisWorkbook.ReadOnly Then
    ThisWorkbook.Save
  End If

End Sub
