Attribute VB_Name = "TestModule10"
'@Folder pUnit
'@TestModule
Option Explicit

'@TestMethod
Public Sub TestAdditionNotParameterised()
  Phosphorus.AssertionsStatic.pAssert.Equal 2 + 2, 4, "2 + 2 should equal 4"
End Sub

'NOTE: Test has parameters but no test data is passed
'@TestMethod
Public Sub TestAdditionParameterised(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub
