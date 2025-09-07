VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assertions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder Assertions
Option Explicit

Private testCount As Long
Private failedCount As Long

Private Type FileIOType
  HasCriticalFailure As Boolean
End Type

Private this As FileIOType

' Initialize the counters when the class is created
Private Sub Class_Initialize()
  testCount = 0
  failedCount = 0
  this.HasCriticalFailure = False
End Sub

' Get total number of tests run
Public Property Get TotalTests() As Long
  TotalTests = testCount
End Property

' Get number of failed tests
Public Property Get FailedTests() As Long
  FailedTests = failedCount
End Property

' Get whether a critical assertion has failed
Public Property Get HasCriticalFailure() As Boolean
  HasCriticalFailure = this.HasCriticalFailure
End Property

' Assert that two values are equal
Public Sub Equal(expected As Variant, actual As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
  testCount = testCount + 1
  If Not AreEqual(expected, actual) Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: Expected '" & expected & "' but got '" & actual & "'" & _
             IIf(Len(message) > 0, " - " & message, "") & _
             IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: Values are equal" & IIf(Len(message) > 0, " - " & message, "") & _
             IIf(isCritical, " [CRITICAL]", "")
  End If
End Sub

' Assert that two values are not equal
Public Sub NotEqual(expected As Variant, actual As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
  testCount = testCount + 1
  If AreEqual(expected, actual) Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: Expected '" & expected & "' to be different from '" & actual & "'" & _
              IIf(Len(message) > 0, " - " & message, "") & _
              IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: Values are not equal" & IIf(Len(message) > 0, " - " & message, "") & _
              IIf(isCritical, " [CRITICAL]", "")
  End If
End Sub

' Helper function to compare two variants
Private Function AreEqual(value1 As Variant, value2 As Variant) As Boolean
  On Error Resume Next
  AreEqual = (value1 = value2)
  If Err.Number <> 0 Then
    AreEqual = False
  End If
  On Error GoTo 0
End Function

' Assert that a condition is true
Public Sub IsTrue(condition As Boolean, Optional message As String = "", Optional isCritical As Boolean = False)
  testCount = testCount + 1
  If Not condition Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: Expected True but got False" & _
               IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: Condition is True" & IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  End If
End Sub

' Assert that a condition is false
Public Sub IsFalse(condition As Boolean, Optional message As String = "", Optional isCritical As Boolean = False)
  testCount = testCount + 1
  If condition Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: Expected False but got True" & _
               IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: Condition is False" & IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  End If
End Sub

' Assert that a value is null
Public Sub ThisIsNull(this As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
  testCount = testCount + 1
  If Not IsNull(this) Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: Expected Null but got '" & this & "'" & _
               IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: Value is Null" & IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  End If
End Sub

' Assert that a value is not null
Public Sub ThisIsNotNull(this As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
  testCount = testCount + 1
  If IsNull(this) Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
      Logger.ExternalInfo "ASSERTION FAILURE: Expected non-Null value but got Null" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
      Logger.ExternalDebug "ASSERTION PASS: Value is not Null" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that a value is within a specified range
Public Sub InRange(value As Variant, minValue As Variant, maxValue As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
  testCount = testCount + 1
  On Error Resume Next
  Dim isInRange As Boolean
  isInRange = (value >= minValue And value <= maxValue)
  If Err.Number <> 0 Then isInRange = False
  On Error GoTo 0
    
  If Not isInRange Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: Value '" & value & "' not in range [" & minValue & ", " & maxValue & "]" & _
               IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: Value in range" & IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  End If
End Sub

' Assert that a collection/array contains a specific item
Public Sub Contains(collection As Variant, item As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
  
  testCount = testCount + 1
  Dim found As Boolean
  found = False
    
  If IsArray(collection) Then
    Dim i As Long
    For i = LBound(collection) To UBound(collection)
      If AreEqual(collection(i), item) Then
        found = True
        Exit For
      End If
    Next i
  ElseIf typeName(collection) = "Collection" Then
    Dim element As Variant
    For Each element In collection
      If AreEqual(element, item) Then
        found = True
        Exit For
      End If
    Next element
  End If
    
  If Not found Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: Item '" & item & "' not found in collection" & _
               IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: Item found in collection" & IIf(Len(message) > 0, " - " & message, "") & _
              IIf(isCritical, " [CRITICAL]", "")
  End If

End Sub

' Assert that a string matches a pattern (simple wildcard matching)
Public Sub StringMatchesSimplePattern(actual As String, pattern As String, Optional message As String = "", Optional isCritical As Boolean = False)
  
  testCount = testCount + 1
  Dim matches As Boolean
  On Error Resume Next
  matches = (actual Like pattern)
  On Error GoTo 0
    
  If Not matches Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: String '" & actual & "' doesn't match pattern '" & pattern & "'" & _
               IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: String matches pattern" & IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  End If

End Sub

' Assert that a value is approximately equal to another within a tolerance
Public Sub ApproximatelyEqual(expected As Double, actual As Double, tolerance As Double, Optional message As String = "", Optional isCritical As Boolean = False)
 
  testCount = testCount + 1
  Dim difference As Double
  difference = Abs(expected - actual)
    
  If difference > tolerance Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: Expected '" & expected & "' ±" & tolerance & " but got '" & actual & "'" & _
               IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: Values approximately equal" & IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  End If

End Sub

' Assert that an object exists (is not Nothing)
Public Sub ObjectExists(obj As Object, Optional message As String = "", Optional isCritical As Boolean = False)
    
  testCount = testCount + 1
  Dim exists As Boolean
  exists = Not (obj Is Nothing)
    
  If Not exists Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: Object is Nothing" & _
               IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: Object exists" & IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  End If

End Sub

' Assert that specific error is raised
Public Sub Throws(methodToRun As String, expectedError As Long, Optional message As String = "", Optional isCritical As Boolean = False)
    
  testCount = testCount + 1
  Dim errorOccurred As Boolean
  errorOccurred = False
    
  'NOTE:
  'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement
  'An On Error Resume Next statement becomes inactive when another procedure is called, so you should execute an On Error Resume Next statement in each called routine if you want inline error handling within that routine.
  'So the method being called must use On Error Resume Next to pass the error back to this assertion
  Application.Run methodToRun
  If Err.Number = expectedError Then
    errorOccurred = True
  End If
  'Always reset error handling
  On Error GoTo 0
    
  If Not errorOccurred Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: Expected error " & expectedError & " but didn't occur or different error occurred" & _
               IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: Expected error occurred" & IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  End If

End Sub

' Assert that a collection/array is empty
Public Sub IsEmpty(collection As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
    
  testCount = testCount + 1
  Dim IsEmpty As Boolean
    
  If IsArray(collection) Then
    Dim l As Integer
    Dim u As Integer
    On Error Resume Next
    l = LBound(collection)
    u = UBound(collection)
    On Error GoTo 0
    IsEmpty = (u = 0 And l = 0)
  ElseIf typeName(collection) = "Collection" Then
    IsEmpty = (collection.count = 0)
  Else
    IsEmpty = False ' Non-collection types considered not empty
  End If
    
  If Not IsEmpty Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: Collection/Array is not empty" & _
               IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: Collection/Array is empty or not a collection or array" & IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  End If

End Sub

' Assert that a variable is of an expected type
Public Sub IsType(value As Variant, expectedTypeName As String, Optional message As String = "", Optional isCritical As Boolean = False)
    
  testCount = testCount + 1
  Dim actualType As String
  actualType = typeName(value)
    
  If VBA.Strings.LCase(actualType) <> VBA.Strings.LCase(expectedTypeName) Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: Expected type '" & expectedTypeName & "' but got '" & actualType & "'" & _
             IIf(Len(message) > 0, " - " & message, "") & _
             IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: Type matches" & IIf(Len(message) > 0, " - " & message, "") & _
             IIf(isCritical, " [CRITICAL]", "")
  End If

End Sub

' Assert that a value is greater than another
Public Sub GreaterThan(value As Variant, threshold As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
    
  testCount = testCount + 1
  Dim isGreater As Boolean
  On Error Resume Next
  isGreater = (value > threshold)
  If Err.Number <> 0 Then isGreater = False
  On Error GoTo 0
    
  If Not isGreater Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: Value '" & value & "' not greater than '" & threshold & "'" & _
               IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: Value greater than threshold" & IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  End If

End Sub

' Assert that a value is less than another
Public Sub LessThan(value As Variant, threshold As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
    
  testCount = testCount + 1
  Dim isLess As Boolean
  On Error Resume Next
  isLess = (value < threshold)
  If Err.Number <> 0 Then isLess = False
  On Error GoTo 0
    
  If Not isLess Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: Value '" & value & "' not less than '" & threshold & "'" & _
               IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: Value less than threshold" & IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  End If

End Sub

' Assert that a file exists
Public Sub FileExists(filePath As String, Optional message As String = "", Optional isCritical As Boolean = False)
    
  testCount = testCount + 1
  Dim exists As Boolean
  exists = (Dir(filePath) <> "")
    
  If Not exists Then
    failedCount = failedCount + 1
    If isCritical Then this.HasCriticalFailure = True
    Logger.ExternalInfo "ASSERTION FAILURE: File '" & filePath & "' does not exist" & _
               IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  Else
    Logger.ExternalDebug "ASSERTION PASS: File exists" & IIf(Len(message) > 0, " - " & message, "") & _
               IIf(isCritical, " [CRITICAL]", "")
  End If

End Sub

' Print summary of test results
Public Sub PrintSummary()
  
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="Test Summary", _
    message:="Test Summary:"
    
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="Test Summary", _
    message:="Total Tests: " & testCount
    
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="Test Summary", _
    message:="Passed: " & (testCount - failedCount)
           
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="Test Summary", _
    message:="Failed: " & failedCount
    
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="Test Summary", _
    message:="Critical Failure Occurred: " & IIf(HasCriticalFailure, "Yes", "No")
    
  Logger.LogFixedLevelMessage _
    level:=LogLevel.EXTERNAL_INFO, _
    AnalysisCode1:="Test Summary", _
    message:="Success Rate: " & VBA.Strings.Format((testCount - failedCount) / testCount, "0%")
  
End Sub

