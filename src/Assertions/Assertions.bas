VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assertions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Assertions
'Option Explicit

'Private testCount As Long
'Private failedCount As Long
'
'' Initialize the counters when the class is created
'Private Sub Class_Initialize()
'  testCount = 0
'  failedCount = 0
'End Sub
'
'' Get total number of tests run
'Public Property Get TotalTests() As Long
'  TotalTests = testCount
'End Property
'
'' Get number of failed tests
'Public Property Get FailedTests() As Long
'  FailedTests = failedCount
'End Property
'
'' Assert that two values are equal
'Public Sub AssertEqual(expected As Variant, actual As Variant, Optional message As String = "")
'  Logger.ExternalInfo "Assert '" & VBA.Conversion.CStr(actual) & "' equals '" & VBA.Conversion.CStr(expected) & "' - " & message
'  testCount = testCount + 1
'  If Not AreEqual(expected, actual) Then
'    failedCount = failedCount + 1
'    Logger.ExternalError "Assertion Failed: Expected '" & expected & "' but got '" & actual & "'" & IIf(Len(message) > 0, " - " & message, "")
'  Else
'    Logger.ExternalDebug "Assertion Passed: Values are equal" & IIf(Len(message) > 0, " - " & message, "")
'  End If
'End Sub
'
'' Assert that two values are not equal
'Public Sub AssertNotEqual(expected As Variant, actual As Variant, Optional message As String = "")
'  testCount = testCount + 1
'  If AreEqual(expected, actual) Then
'    failedCount = failedCount + 1
'    Logger.ExternalError "Assertion Failed: Expected '" & expected & "' to be different from '" & actual & "'" & IIf(Len(message) > 0, " - " & message, "")
'  Else
'    Logger.ExternalDebug "Assertion Passed: Values are not equal" & IIf(Len(message) > 0, " - " & message, "")
'  End If
'End Sub
'
'' Assert that a condition is true
'Public Sub AssertTrue(condition As Boolean, Optional message As String = "")
'  testCount = testCount + 1
'  If Not condition Then
'    failedCount = failedCount + 1
'    Logger.ExternalError "Assertion Failed: Expected True but got False" & IIf(Len(message) > 0, " - " & message, "")
'  Else
'    Logger.ExternalDebug "Assertion Passed: Condition is True" & IIf(Len(message) > 0, " - " & message, "")
'  End If
'End Sub
'
'' Assert that a condition is false
'Public Sub AssertFalse(condition As Boolean, Optional message As String = "")
'  testCount = testCount + 1
'  If condition Then
'    failedCount = failedCount + 1
'    Logger.ExternalError "Assertion Failed: Expected False but got True" & IIf(Len(message) > 0, " - " & message, "")
'  Else
'    Logger.ExternalDebug "Assertion Passed: Condition is False" & IIf(Len(message) > 0, " - " & message, "")
'  End If
'End Sub
'
'' Assert that a value is null
'Public Sub AssertNull(value As Variant, Optional message As String = "")
'  testCount = testCount + 1
'  If Not IsNull(value) Then
'    failedCount = failedCount + 1
'    Logger.ExternalError "Assertion Failed: Expected Null but got '" & value & "'" & IIf(Len(message) > 0, " - " & message, "")
'  Else
'    Logger.ExternalDebug "Assertion Passed: Value is Null" & IIf(Len(message) > 0, " - " & message, "")
'  End If
'End Sub
'
'' Assert that a value is not null
'Public Sub AssertNotNull(value As Variant, Optional message As String = "")
'  testCount = testCount + 1
'  If IsNull(value) Then
'    failedCount = failedCount + 1
'    Logger.ExternalError "Assertion Failed: Expected non-Null value but got Null" & IIf(Len(message) > 0, " - " & message, "")
'  Else
'    Logger.ExternalDebug "Assertion Passed: Value is not Null" & IIf(Len(message) > 0, " - " & message, "")
'  End If
'End Sub
'
'' Helper function to compare two variants
'Private Function AreEqual(value1 As Variant, value2 As Variant) As Boolean
'  On Error Resume Next
'  AreEqual = (value1 = value2)
'  If Err.Number <> 0 Then
'      AreEqual = False
'  End If
'  On Error GoTo 0
'End Function
'
'' Print summary of test results
'Public Sub PrintSummary()
'  Logger.InternalInfo "Test Summary:"
'  Logger.InternalInfo "Total Tests: " & testCount
'  Logger.InternalInfo "Passed: " & (testCount - failedCount)
'  Logger.InternalInfo "Failed: " & failedCount
'  Logger.InternalInfo "Success Rate: " & Format((testCount - failedCount) / testCount, "0%")
'End Sub


' Assertions.cls
Option Explicit

Private testCount As Long
Private failedCount As Long
Private HasCriticalFailure As Boolean

' Initialize the counters when the class is created
Private Sub Class_Initialize()
    testCount = 0
    failedCount = 0
    HasCriticalFailure = False
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
Public Property Get pjgHasCriticalFailure() As Boolean
    pjgHasCriticalFailure = HasCriticalFailure
End Property

' Assert that two values are equal
Public Sub AssertEqual(expected As Variant, actual As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    If Not AreEqual(expected, actual) Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Expected '" & expected & "' but got '" & actual & "'" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Values are equal" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that two values are not equal
Public Sub AssertNotEqual(expected As Variant, actual As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    If AreEqual(expected, actual) Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Expected '" & expected & "' to be different from '" & actual & "'" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Values are not equal" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that a condition is true
Public Sub AssertTrue(condition As Boolean, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    If Not condition Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Expected True but got False" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Condition is True" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that a condition is false
Public Sub AssertFalse(condition As Boolean, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    If condition Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Expected False but got True" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Condition is False" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that a value is null
Public Sub AssertNull(value As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    If Not IsNull(value) Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Expected Null but got '" & value & "'" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Value is Null" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that a value is not null
Public Sub AssertNotNull(value As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    If IsNull(value) Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Expected non-Null value but got Null" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Value is not Null" & IIf(Len(message) > 0, " - " & message, "") & _
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

' Additional methods to add to Assertions.cls

' Assert that a value is within a specified range
Public Sub AssertInRange(value As Variant, minValue As Variant, maxValue As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    On Error Resume Next
    Dim isInRange As Boolean
    isInRange = (value >= minValue And value <= maxValue)
    If Err.Number <> 0 Then isInRange = False
    On Error GoTo 0
    
    If Not isInRange Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Value '" & value & "' not in range [" & minValue & ", " & maxValue & "]" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Value in range" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that a collection/array contains a specific item
Public Sub AssertContains(collection As Variant, item As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
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
    ElseIf TypeName(collection) = "Collection" Then
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
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Item '" & item & "' not found in collection" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Item found in collection" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that a string matches a pattern (simple wildcard matching)
Public Sub AssertStringMatches(actual As String, pattern As String, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    Dim matches As Boolean
    On Error Resume Next
    matches = (actual Like pattern)
    On Error GoTo 0
    
    If Not matches Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: String '" & actual & "' doesn't match pattern '" & pattern & "'" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: String matches pattern" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that a value is approximately equal to another within a tolerance
Public Sub AssertApproximatelyEqual(expected As Double, actual As Double, tolerance As Double, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    Dim difference As Double
    difference = Abs(expected - actual)
    
    If difference > tolerance Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Expected '" & expected & "' ±" & tolerance & " but got '" & actual & "'" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Values approximately equal" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that an object exists (is not Nothing)
Public Sub AssertObjectExists(obj As Object, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    Dim exists As Boolean
    exists = Not (obj Is Nothing)
    
    If Not exists Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Object is Nothing" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Object exists" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Additional methods to add to Assertions.cls

' Assert that specific error is raised
Public Sub AssertThrows(codeToRun As String, expectedError As Long, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    Dim errorOccurred As Boolean
    errorOccurred = False
    
    On Error Resume Next
    ExecuteScript codeToRun
    If Err.Number = expectedError Then
        errorOccurred = True
    End If
    On Error GoTo 0
    
    If Not errorOccurred Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Expected error " & expectedError & " but didn't occur or different error occurred" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Expected error occurred" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that a collection/array is empty
Public Sub AssertEmpty(collection As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    Dim isEmpty As Boolean
    
    If IsArray(collection) Then
        isEmpty = (UBound(collection) < LBound(collection))
    ElseIf TypeName(collection) = "Collection" Then
        isEmpty = (collection.Count = 0)
    Else
        isEmpty = False ' Non-collection types considered not empty
    End If
    
    If Not isEmpty Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Collection is not empty" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Collection is empty" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that a variable is of an expected type
Public Sub AssertType(value As Variant, expectedTypeName As String, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    Dim actualType As String
    actualType = TypeName(value)
    
    If LCase(actualType) <> LCase(expectedTypeName) Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Expected type '" & expectedTypeName & "' but got '" & actualType & "'" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Type matches" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that a value is greater than another
Public Sub AssertGreaterThan(value As Variant, threshold As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    Dim isGreater As Boolean
    On Error Resume Next
    isGreater = (value > threshold)
    If Err.Number <> 0 Then isGreater = False
    On Error GoTo 0
    
    If Not isGreater Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Value '" & value & "' not greater than '" & threshold & "'" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Value greater than threshold" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that a value is less than another
Public Sub AssertLessThan(value As Variant, threshold As Variant, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    Dim isLess As Boolean
    On Error Resume Next
    isLess = (value < threshold)
    If Err.Number <> 0 Then isLess = False
    On Error GoTo 0
    
    If Not isLess Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: Value '" & value & "' not less than '" & threshold & "'" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: Value less than threshold" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Assert that a file exists
Public Sub AssertFileExists(filePath As String, Optional message As String = "", Optional isCritical As Boolean = False)
    testCount = testCount + 1
    Dim exists As Boolean
    exists = (Dir(filePath) <> "")
    
    If Not exists Then
        failedCount = failedCount + 1
        If isCritical Then HasCriticalFailure = True
        Debug.Print "FAIL: File '" & filePath & "' does not exist" & _
                   IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    Else
        Debug.Print "PASS: File exists" & IIf(Len(message) > 0, " - " & message, "") & _
                   IIf(isCritical, " [CRITICAL]", "")
    End If
End Sub

' Helper function for AssertThrows (Note: VBA has limited support for this)
Private Sub ExecuteScript(code As String)
    Application.VBE.ActiveVBProject.VBComponents.Add(vbext_ct_StdModule).CodeModule.AddFromString code
    Application.Run "TempModule"
    Application.VBE.ActiveVBProject.VBComponents.Remove Application.VBE.ActiveVBProject.VBComponents("TempModule")
End Sub


' Print summary of test results
Public Sub PrintSummary()
    Debug.Print vbNewLine & "Test Summary:"
    Debug.Print "Total Tests: " & testCount
    Debug.Print "Passed: " & (testCount - failedCount)
    Debug.Print "Failed: " & failedCount
    Debug.Print "Critical Failure Occurred: " & IIf(HasCriticalFailure, "Yes", "No")
    Debug.Print "Success Rate: " & Format((testCount - failedCount) / testCount, "0%")
End Sub

