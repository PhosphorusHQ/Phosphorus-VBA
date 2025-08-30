Attribute VB_Name = "AssertionsStatic"
'@Folder Assertions
Option Explicit

Public pAssert As Assertions

'TODO: Can't make this available until we remove Rubberduck
'Public Assert As Assertions
'
'Public Sub GetAssert()
'  If Assert Is Nothing Then
'    Set Assert = New Phosphorus.Assertions
'  End If
'End Sub

Public Sub GetAssert()
  If pAssert Is Nothing Then
    Set pAssert = New Phosphorus.Assertions
  End If
End Sub

Public Sub CloseAssert()
  Set pAssert = Nothing
End Sub
