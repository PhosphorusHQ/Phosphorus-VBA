Attribute VB_Name = "pUnitErrorStatic"
'@Folder pUnit
Option Explicit

'Custom error onject - callign Applicaiton.Run xxxx does not pass back errors as normally expected
Public gpUnitError As Phosphorus.pUnitError

Public Function TrapError(Number As Long, Description As String)
 Set gpUnitError = New pUnitError
 gpUnitError.Number = Number
 gpUnitError.Description = Description
End Function

Public Function ClearAllErrors()
  Set gpUnitError = Nothing
End Function
