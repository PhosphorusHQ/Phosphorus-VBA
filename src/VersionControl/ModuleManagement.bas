Attribute VB_Name = "ModuleManagement"
'@Folder VersionControl
Option Explicit

'This code requires the "Microsoft Visual Basic for Applications Extensibility" reference to be enabled in VBA (Tools > References).
'The workbook must be saved in a macro-enabled format (e.g. .xlsm).
'Trust access to the VBA project object model must be enabled (File > Options > Trust Center > Trust Center Settings > Macro Settings).

' Subroutine to export all modules to text files with folder tagging
Public Sub ExportModulesWithFolders(SubFolderForExport As String, Optional ProjectName As String = "")
    
  Dim vbProj As VBProject
  Dim vbComp As VBComponent
  Dim folderName As String
  Dim filePath As String
  Dim basePath As String
  Dim fso As Object
  Dim file As Object
  Dim firstLine As String
    
  ' Get VBE and set project reference
  With Application.VBE
    If ProjectName = "" Then
      Set vbProj = ThisWorkbook.VBProject
      basePath = ThisWorkbook.Path & SubFolderForExport
    Else
      Dim projFound As Boolean
      projFound = False
      For Each vbProj In .VBProjects
        If vbProj.Name = ProjectName Then
          projFound = True
          basePath = vbProj.Filename
          basePath = Left(basePath, InStrRev(basePath, "\")) & "VBA_Modules_Export\"
          Exit For
        End If
      Next vbProj
            
      If Not projFound Then
        MsgBox "Project '" & ProjectName & "' not found.", vbCritical
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
      firstLine = Trim(vbComp.CodeModule.Lines(1, 1))
      If Left(firstLine, 9) = "'@Folder " Then
        folderName = Trim(Mid(firstLine, 9))
      End If
    End If
    
    If folderName <> "None" Then
    
      If Not fso.FolderExists(basePath & "\" & folderName) Then
        fso.CreateFolder basePath & "\" & folderName
      End If
        
      filePath = basePath & "\" & folderName & "\" & vbComp.Name & ".bas"
      vbComp.Export filePath
        
      If vbComp.Type <> 3 Then
        Dim fileText As String
        With CreateObject("Scripting.FileSystemObject")
          fileText = .OpenTextFile(filePath, 1).ReadAll
          Set file = .CreateTextFile(filePath, True)
          If Left(firstLine, 7) = "@Folder " Then
            file.WriteLine firstLine
          End If
          file.Write fileText
          file.Close
        End With
      End If
    
    End If
  
  Next vbComp

  MsgBox "Modules exported successfully to: " & basePath, vbInformation
End Sub

' Subroutine to remove all components except specified ones
Public Sub RemoveAllComponentsExcept(vbProj As Object, ParamArray ModulesToKeep() As Variant)
    Dim vbComp As Object
    Dim i As Long
    
    ' Remove components
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type <> 100 Then ' vbext_ct_Document
            Dim shouldKeep As Boolean
            shouldKeep = False
            For i = LBound(ModulesToKeep) To UBound(ModulesToKeep)
                If vbComp.Name = ModulesToKeep(i) Then
                    shouldKeep = True
                    Exit For
                End If
            Next i
            
            If Not shouldKeep Then
                vbProj.VBComponents.Remove vbComp
            End If
        End If
    Next vbComp
End Sub

' Subroutine to import modules from folder structure
Public Sub ImportModulesFromFolder(ProjectName As String, ParamArray ModulesToKeep() As Variant)
    Dim vbProj As Object
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim importPath As String
    Dim projFound As Boolean
    
    ' Get VBE and set project reference
    With Application.VBE
        If ProjectName = "" Then
            Set vbProj = ThisWorkbook.VBProject
            importPath = ThisWorkbook.Path & "\VBA_Modules_Export\"
        Else
            projFound = False
            For Each vbProj In .VBProjects
                If vbProj.Name = ProjectName Then
                    projFound = True
                    importPath = vbProj.Filename
                    importPath = Left(importPath, InStrRev(importPath, "\")) & "VBA_Modules_Export\"
                    Exit For
                End If
            Next vbProj
            
            If Not projFound Then
                MsgBox "Project '" & ProjectName & "' not found.", vbCritical
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
    RemoveAllComponentsExcept vbProj, ModulesToKeep
    
    ' Import all .bas files from folder and subfolders
    Set folder = fso.GetFolder(importPath)
    
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".bas" Then
            vbProj.VBComponents.Import file.Path
        End If
    Next file
    
    For Each subFolder In folder.SubFolders
        For Each file In subFolder.Files
            If LCase(Right(file.Name, 4)) = ".bas" Then
                vbProj.VBComponents.Import file.Path
            End If
        Next file
    Next subFolder
    
    MsgBox "Modules imported successfully from: " & importPath, vbInformation
End Sub

'' Subroutine to export all modules to text files with folder tagging
'Sub ExportModulesWithFolders(SubFolderForExport As String, Optional ProjectName As String = "")
'
'  Dim VBProject As Object
'  Dim VBComponent As Object
'  Dim folderName As String
'  Dim filePath As String
'  Dim basePath As String
'  Dim fso As Object
'  Dim file As Object
'  Dim firstLine As String
'  Dim projFound As Boolean
'
'  ' Get VBE and set project reference
'  With Application.VBE
'    If ProjectName = "" Then
'      ' Use current project if no name provided
'      Set VBProject = ThisWorkbook.VBProject
'      basePath = ThisWorkbook.Path & SubFolderForExport
'    Else
'      ' Search for specified project among open projects
'      projFound = False
'      For Each VBProject In .VBProjects
'        If VBProject.Name = ProjectName Then
'          projFound = True
'          basePath = VBProject.Filename
'          basePath = Left(basePath, InStrRev(basePath, "\")) & "VBA_Modules_Export\"
'          Exit For
'        End If
'      Next VBProject
'
'      If Not projFound Then
'        MsgBox "Project '" & ProjectName & "' not found.", vbCritical
'        Exit Sub
'      End If
'    End If
'  End With
'
'  ' Create FileSystemObject
'  Set fso = CreateObject("Scripting.FileSystemObject")
'
'  ' Create base folder if it doesn't exist
'  If Not fso.FolderExists(basePath) Then
'    fso.CreateFolder basePath
'  End If
'
'  ' Loop through all components
'  For Each VBComponent In VBProject.VBComponents
'    ' Default folder if no tag found
'     folderName = "Uncategorized"
'
'    ' Check if module has a folder tag in first line
'    If VBComponent.CodeModule.CountOfLines > 0 Then
'      firstLine = Trim(VBComponent.CodeModule.Lines(1, 1))
'      If Left(firstLine, 9) = "'@Folder " Then
'        folderName = (Mid(firstLine, 9)) ' Extract folder name after "@Folder "
'      End If
'    End If
'
'    ' Create subfolder if it doesn't exist
'    If Not fso.FolderExists(basePath & "\" & folderName) Then
'      fso.CreateFolder basePath & "\" & folderName
'    End If
'
'    ' Set file path
'    filePath = basePath & "\" & folderName & "\" & VBComponent.Name & ".bas"
'
'    ' Export the component
'    VBComponent.Export filePath
'
'    ' For non-form components, preserve the folder tag
'    If VBComponent.Type <> 3 Then ' Skip forms as they use .frm format
'      ' Read the exported file
'      Dim fileText As String
'      With CreateObject("Scripting.FileSystemObject")
'        fileText = .OpenTextFile(filePath, 1).ReadAll
'
'        ' Create new file with existing folder tag (if any)
'        Set file = .CreateTextFile(filePath, True)
'        If Left(firstLine, 7) = "@Folder " Then
'          file.WriteLine firstLine ' Preserve existing tag
'        End If
'        file.Write fileText
'        file.Close
'      End With
'    End If
'  Next VBComponent
'
'   MsgBox "Modules exported successfully to: " & basePath, vbInformation
'
'End Sub
'
'' Subroutine to import modules from folder structure
'Sub ImportModulesFromFolder(Optional ProjectName As String = "")
'    Dim vbProj As Object
'    Dim fso As Object
'    Dim folder As Object
'    Dim subFolder As Object
'    Dim file As Object
'    Dim importPath As String
'    Dim projFound As Boolean
'
'    ' Get VBE and set project reference
'    With Application.VBE
'        If ProjectName = "" Then
'            ' Use current project if no name provided
'            Set vbProj = ThisWorkbook.VBProject
'            importPath = ThisWorkbook.Path & "\VBA_Modules_Export\"
'        Else
'            ' Search for specified project among open projects
'            projFound = False
'            For Each vbProj In .VBProjects
'                If vbProj.Name = ProjectName Then
'                    projFound = True
'                    importPath = vbProj.Filename
'                    importPath = Left(importPath, InStrRev(importPath, "\")) & "VBA_Modules_Export\"
'                    Exit For
'                End If
'            Next vbProj
'
'            If Not projFound Then
'                MsgBox "Project '" & ProjectName & "' not found.", vbCritical
'                Exit Sub
'            End If
'        End If
'    End With
'
'    ' Create FileSystemObject
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    ' Check if folder exists
'    If Not fso.FolderExists(importPath) Then
'        MsgBox "Import folder not found: " & importPath, vbCritical
'        Exit Sub
'    End If
'
'    ' Remove all existing components except the document module of the project
'    Dim vbComp As Object
'    For Each vbComp In vbProj.VBComponents
'        If vbComp.Type <> 100 Then ' vbext_ct_Document
'            vbProj.VBComponents.Remove vbComp
'        End If
'    Next vbComp
'
'    ' Import all .bas files from folder and subfolders
'    Set folder = fso.GetFolder(importPath)
'
'    ' Process main folder
'    For Each file In folder.Files
'        If LCase(Right(file.Name, 4)) = ".bas" Then
'            vbProj.VBComponents.Import file.Path
'        End If
'    Next file
'
'    ' Process subfolders
'    For Each subFolder In folder.SubFolders
'        For Each file In subFolder.Files
'            If LCase(Right(file.Name, 4)) = ".bas" Then
'                vbProj.VBComponents.Import file.Path
'            End If
'        Next file
'    Next subFolder
'
'    MsgBox "Modules imported successfully from: " & importPath, vbInformation
'End Sub

