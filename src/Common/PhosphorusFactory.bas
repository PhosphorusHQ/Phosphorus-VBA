Attribute VB_Name = "PhosphorusFactory"
'@Folder Common
Option Explicit

Public Configuration As Phosphorus.ConfigurationReader

'Global Logger object available for use
Public Logger As Log4P

Public Sub GetConfigurationReader()
  If Configuration Is Nothing Then
    Set Configuration = New Phosphorus.ConfigurationReader
  End If
End Sub

Public Sub GetLogger()
  If Logger Is Nothing Then
    Set Logger = New Phosphorus.Log4P
  End If
End Sub

Public Function GetNewPhosphorusPPath() As Phosphorus.PPath
  Set GetNewPhosphorusPPath = New Phosphorus.PPath
End Function

Public Function GetNewPhosphorusPPathReturnClass(intNumberOfPPathExpressions As Integer) As Phosphorus.PPathReturnClass
  Set GetNewPhosphorusPPathReturnClass = New Phosphorus.PPathReturnClass
  GetNewPhosphorusPPathReturnClass.Initialise intNumberOfPPathExpressions
End Function

