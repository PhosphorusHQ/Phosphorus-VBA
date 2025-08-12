Attribute VB_Name = "Factory"
'@Folder Common
Option Explicit

Public Configuration As Phosphorus.ConfigurationReader

Public Sub GetConfigurationReader()
  If Configuration Is Nothing Then
    Set Configuration = New Phosphorus.ConfigurationReader
  End If
End Sub

Public Function GetNewPhosphorusPPath() As Phosphorus.PPath
  Set GetNewPhosphorusPPath = New Phosphorus.PPath
End Function

Public Function GetNewPhosphorusPPathReturnClass(intNumberOfPPathExpressions As Integer) As Phosphorus.PPathReturnClass
  Set GetNewPhosphorusPPathReturnClass = New Phosphorus.PPathReturnClass
  GetNewPhosphorusPPathReturnClass.Initialise intNumberOfPPathExpressions
End Function

Public Function GetNewPhosphorusLog4P() As Phosphorus.Log4P
  Set GetNewPhosphorusLog4P = New Phosphorus.Log4P
End Function

Public Function GetNewPhosphorusAssertions() As Phosphorus.Assertions
  Set GetNewPhosphorusAssertions = New Phosphorus.Assertions
End Function

Public Function GetNewPhosphorusStringBuilder() As Phosphorus.StringBuilder
  Set GetNewPhosphorusStringBuilder = New Phosphorus.StringBuilder
End Function

