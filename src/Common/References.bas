Attribute VB_Name = "References"
'@Folder Common
Option Explicit

' Declare a global Collection to store added reference details (FullPath and Workbook)
Private AddedReferences As Collection

'Trust Settings: Ensure "Trust access to the VBA project object model" is enabled in Excel's Trust Center (File > Options > Trust Center > Trust Center Settings > Macro Settings).
'NOTE: We don't seem to be able to add a reference while running debug mode!

' Initialize the collection (call this before adding references)
Sub InitialiseAddedReferences()
  Set AddedReferences = New Collection
End Sub

' Add a reference to a macro-enabled workbook and save its FullPath and Workbook
Sub AddReferenceToWorkbook(targetWb As Workbook)
  
  Dim vbProj As VBProject
  Dim ref As Reference
  Dim refDetails As Collection ' To store FullPath and Workbook
    
  ' Initialize the collection if not already done
  If AddedReferences Is Nothing Then InitialiseAddedReferences
    
  On Error Resume Next
    
  ' Validate the target workbook
  If targetWb Is Nothing Then
'    MsgBox "No valid workbook provided."
    Exit Sub
  End If
    
  ' Get the VBProject of the current workbook
  Set vbProj = ThisWorkbook.VBProject
    
  ' Check if reference already exists
  For Each ref In vbProj.References
    If ref.FullPath = targetWb.FullName Then
'      MsgBox "Reference to " & targetWb.FullName & " already exists."
      Exit Sub
    End If
  Next ref
    
  ' Add the reference to the target workbook
  Set ref = vbProj.References.AddFromFile(targetWb.FullName)
     
  ' Create a collection to store reference details
  Set refDetails = New Collection
  refDetails.Add targetWb.FullName, "FullPath" ' Store the filepath
  refDetails.Add targetWb, "Workbook"         ' Store the Workbook object
    
  ' Save the reference details to the global collection
  AddedReferences.Add refDetails
'  MsgBox "Reference to " & targetWb.FullName & " added and saved."
    
  On Error GoTo 0

End Sub

' Remove all references stored in the AddedReferences collection
Sub RemoveAllAddedReferences()
    
  Dim vbProj As Object ' VBProject
  Dim ref As Object ' Reference
  Dim refDetails As Collection
  Dim removedCount As Long
    
  ' Check if collection is initialized
  If AddedReferences Is Nothing Then
'    MsgBox "No references have been added to remove."
    Exit Sub
  End If
    
  Set vbProj = ThisWorkbook.VBProject
  removedCount = 0

  ' Iterate through the saved reference details
  On Error Resume Next
  For Each refDetails In AddedReferences
      
    ' Get the FullPath from the refDetails collection
    Dim refPath As String
    refPath = refDetails("FullPath")
        
    ' Find and remove the reference in the VBProject
    For Each ref In vbProj.References
      If ref.FullPath = refPath Then
        vbProj.References.Remove ref
        removedCount = removedCount + 1
        Exit For ' Exit inner loop once reference is removed
      End If
    Next ref
  
  Next refDetails
    
  On Error GoTo 0
    
  ' Clear the collection
  Set AddedReferences = New Collection
    
'MsgBox "Added Reference"
    
End Sub

' Example usage: Open a workbook and add its reference
Sub TestAddReference()

  Dim strPath As String
  strPath = VBA.Strings.Left(ThisWorkbook.FullName, VBA.InStr(1, ThisWorkbook.FullName, ThisWorkbook.Name) - 1) & "tests\UserDefinedFunctionsTest.xlam"
    
  Dim wb As Workbook
  Dim wbLoop As Workbook
  For Each wbLoop In Application.Workbooks
    If wbLoop.FullName = strPath Then
      wbLoop.Close
      Exit For
    End If
  Next
  If wb Is Nothing Then
    Set wb = Workbooks.Open(strPath)
  End If
  
  If Not wb Is Nothing Then
    AddReferenceToWorkbook wb
    MsgBox Application.Run("AllDogsAreCats", "Dog")
    RemoveAllAddedReferences
    wb.Close
  Else
    MsgBox "Failed to open the workbook."
  End If
End Sub
