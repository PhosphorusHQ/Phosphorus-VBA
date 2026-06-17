Attribute VB_Name = "ModuleManagement"
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

'This code requires the "Microsoft Visual Basic for Applications Extensibility" reference to be enabled in VBA (Tools > References).
'The workbook must be saved in a macro-enabled format (e.g. .xlsm).
'Trust access to the VBA project object model must be enabled (File > Options > Trust Center > Trust Center Settings > Macro Settings).
 
Private ModulesToKeep() As String

Public Sub ExportModulesWithFolders(SubFolderForExport As String, Optional projectName As String = "")
    
  Dim vbProj As VBProject
  Dim vbComp As VBComponent
  Dim folderName As String
  Dim FilePath As String
  Dim basePath As String
  Dim fso As Object
  Dim file As Object
  Dim firstLine As String
    
  ' Get VBE and set project reference
  With Application.VBE
    If projectName = "" Then
      Set vbProj = ThisWorkbook.VBProject
      basePath = ThisWorkbook.Path & SubFolderForExport
    Else
      Dim projFound As Boolean
      projFound = False
      For Each vbProj In .VBProjects
        If vbProj.Name = projectName Then
          projFound = True
          basePath = VBA.Strings.Left(vbProj.Filename, VBA.Strings.InStrRev(vbProj.Filename, "\", -1) - 1) & SubFolderForExport
          Exit For
        End If
      Next vbProj
      If Not projFound Then
        MsgBox "Project '" & projectName & "' not found.", vbCritical
        Exit Sub
      End If
    End If
  End With
    
  ' Create FileSystemObject
  Set fso = CreateObject("Scripting.FileSystemObject")
    
  ' Create base folder if it doesn't exist
  If Not fso.FolderExists(basePath) Then
    fso.CreateFolder basePath
  End If
    
  ' Loop through all components
  For Each vbComp In vbProj.VBComponents
    folderName = "Uncategorized"
    If vbComp.CodeModule.CountOfLines > 0 Then
      firstLine = VBA.Strings.Trim(vbComp.CodeModule.Lines(1, 1))
      If VBA.Strings.Left(firstLine, 9) = "'@Folder " Then
        folderName = VBA.Strings.Trim(VBA.Strings.Mid(firstLine, 9))
      End If
    End If
    
    If folderName <> "None" Then
    
      If Not fso.FolderExists(basePath & "\" & folderName) Then
        fso.CreateFolder basePath & "\" & folderName
      End If
        
      FilePath = basePath & "\" & folderName & "\" & vbComp.Name & ".bas"
      vbComp.Export FilePath
        
      If vbComp.Type <> 3 Then
        Dim fileText As String
        With CreateObject("Scripting.FileSystemObject")
          fileText = .OpenTextFile(FilePath, 1).ReadAll
          Set file = .CreateTextFile(FilePath, True)
          If VBA.Strings.Left(firstLine, 7) = "@Folder " Then
            file.WriteLine firstLine
          End If
          file.Write fileText
          file.Close
        End With
      End If
    
    End If
  
  Next vbComp

  'MsgBox "Modules exported successfully to: " & basePath, vbInformation

End Sub

Public Sub SetModulesToKeep(strModulesToKeep() As String)
  ModulesToKeep = strModulesToKeep()
End Sub

'Public Sub RemoveAllComponentsExcept(vbProj As Object, ModulesToKeep() As Variant)
Public Sub RemoveAllComponentsExcept(vbProj As Object)
  Dim vbComp As Object
  Dim i As Long
  For Each vbComp In vbProj.VBComponents
    If vbComp.Type <> 100 Then ' vbext_ct_Document
      Dim shouldKeep As Boolean
      shouldKeep = False
      If IsArray(ModulesToKeep) Then
        For i = LBound(ModulesToKeep) To UBound(ModulesToKeep)
'Debug.Print ModulesToKeep(i)
          If vbComp.Name = CStr(ModulesToKeep(i)) Then
           shouldKeep = True
            Exit For
          End If
        Next i
      End If
      If Not shouldKeep Then
        vbProj.VBComponents.Remove vbComp
      End If
    End If
  Next vbComp
End Sub

'Updated
'Public Sub ImportModulesFromFolder(SubFolderForExport As String, ProjectName As String, ParamArray ModulesToKeep() As Variant)
Public Sub ImportModulesFromFolder(SubFolderForExport As String, projectName As String)
   
  Dim vbProj As Object
  Dim fso As Object
  Dim folder As Object
  Dim subFolder As Object
  Dim file As Object
  Dim importPath As String
  Dim projFound As Boolean
    
  ' Get VBE and set project reference
  With Application.VBE
    If projectName = "" Then
      Set vbProj = ThisWorkbook.VBProject
      importPath = ThisWorkbook.Path & "\" & SubFolderForExport
    Else
      projFound = False
      For Each vbProj In .VBProjects
        If vbProj.Name = projectName Then
          projFound = True
          importPath = vbProj.Filename
          importPath = VBA.Strings.Left(importPath, VBA.Strings.InStrRev(importPath, "\")) & "\" & SubFolderForExport
          Exit For
        End If
      Next vbProj
            
      If Not projFound Then
        MsgBox "Project '" & projectName & "' not found.", vbCritical
        Exit Sub
      End If
    End If
  End With
    
  ' Create FileSystemObject
  Set fso = CreateObject("Scripting.FileSystemObject")
    
  ' Check if folder exists
  If Not fso.FolderExists(importPath) Then
    MsgBox "Import folder not found: " & importPath, vbCritical
    Exit Sub
  End If
    
  ' Remove all components except specified modules
    RemoveAllComponentsExcept vbProj
    
  ' Import all .bas files from folder and subfolders
  Set folder = fso.GetFolder(importPath)
  Dim i As Long
  Dim keep As Boolean
  
  For Each file In folder.Files
    If VBA.Strings.LCase(VBA.Strings.Right(file.Name, 4)) = ".bas" Then
     keep = False
      If IsArray(ModulesToKeep) Then
        For i = LBound(ModulesToKeep) To UBound(ModulesToKeep)
          If VBA.Strings.Left(file.Name, VBA.Strings.Len(file.Name) - 4) = ModulesToKeep(i) Then
            keep = True
            Exit For
          End If
        Next i
        If Not keep Then
          vbProj.VBComponents.Import file.Path
        End If
      End If
    End If
  Next file
    
  For Each subFolder In folder.SubFolders
    For Each file In subFolder.Files
      If VBA.Strings.LCase(VBA.Strings.Right(file.Name, 4)) = ".bas" Then
        keep = False
        If IsArray(ModulesToKeep) Then
          For i = LBound(ModulesToKeep) To UBound(ModulesToKeep)
            If VBA.Strings.Left(file.Name, VBA.Strings.Len(file.Name) - 4) = ModulesToKeep(i) Then
              keep = True
              Exit For
            End If
          Next i
          If Not keep Then
            vbProj.VBComponents.Import file.Path
          End If
        End If
      End If
    Next file
  Next subFolder
    
  MsgBox "Modules imported successfully from: " & importPath, vbInformation

End Sub

