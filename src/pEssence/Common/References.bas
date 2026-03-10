Attribute VB_Name = "References"
'@Folder Common
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

'Trust Settings: Ensure "Trust access to the VBA project object model" is enabled in Excel's Trust Center (File > Options > Trust Center > Trust Center Settings > Macro Settings).

' Add a reference to a DLL, etc or macro-enabled workbook and save its FullPath and project name
Public Sub AddReferenceToWorkbookOrLibrary(TargetFilePath As String, Optional vbTargetProjName As String)
  
  Dim vbProj As VBProject
  
  If Dir(TargetFilePath) = "" Then
    MsgBox "References: The file " & TargetFilePath & " does not exist!"
    Exit Sub
  End If

  Dim vbOpenProject As VBProject
  If vbTargetProjName <> "" Then
     For Each vbOpenProject In Application.VBE.VBProjects
       If vbOpenProject.Name = vbTargetProjName Then
         Set vbProj = vbOpenProject
       End If
     Next vbOpenProject
  End If
  
  ' Get the VBProject of the current workbook
  If vbProj Is Nothing Then
    Set vbProj = ThisWorkbook.VBProject
  End If
  
  ' Check if reference already exists
  Dim ref As Reference
  For Each ref In vbProj.References
    If ref.FullPath = TargetFilePath Then
      Debug.Print "References: Reference to " & TargetFilePath & " already exists."
      Exit Sub
    End If
  Next ref

  ' Add the reference to the target workbook
  Set ref = vbProj.References.AddFromFile(TargetFilePath)
    
  On Error GoTo 0

End Sub

Public Sub ListAllReferencesInAProject(vbTargetProjectName As String)
    
  Dim vbTargetProject As VBProject
  Dim vbOpenProject As VBProject
  For Each vbOpenProject In Application.VBE.VBProjects
    If vbOpenProject.Name = vbTargetProjectName Then
      Set vbTargetProject = vbOpenProject
      Exit For
    End If
  Next vbOpenProject
    
  Dim ExistingReference As Reference
  For Each ExistingReference In vbOpenProject.References
    Debug.Print "BuiltIn: " & ExistingReference.BuiltIn & ": " & ExistingReference.FullPath
  Next ExistingReference

End Sub

Public Sub RemoveAllNonBuiltInReferencesFromAProject(vbTargetProjectName As String)

  Dim vbTargetProject As VBProject
  Dim vbOpenProject As VBProject
  For Each vbOpenProject In Application.VBE.VBProjects
    If vbOpenProject.Name = vbTargetProjectName Then
      Set vbTargetProject = vbOpenProject
      Exit For
    End If
  Next vbOpenProject

  Dim ExistingReference As Reference
  For Each ExistingReference In vbOpenProject.References
    'Remove all non-BuiltIn references
    ' - but not Microsoft Visual Basic for Applications Extensilbility 5.3 - needed for VBProject
    ' - and not Microsoft Office 16 Object Library - needed for IRibbonUI
    If Not ExistingReference.BuiltIn And _
       ExistingReference.FullPath <> "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB" And _
       ExistingReference.FullPath <> "C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL" _
    Then
      Dim ExistingReferenceFullPath As String
      ExistingReferenceFullPath = ExistingReference.FullPath
      vbOpenProject.References.Remove ExistingReference
      Debug.Print "ExistingReference " & ExistingReferenceFullPath & " removed!"
    End If
  Next ExistingReference

End Sub






