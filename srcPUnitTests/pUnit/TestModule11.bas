Attribute VB_Name = "TestModule11"
'@Folder pUnit
'@TestModule
'@Tag(TestingTestData)
Option Explicit

'@TestMethod
Public Sub TestAdditionNotParameterised()
  Phosphorus.AssertionsStatic.pAssert.Equal 2 + 2, 4, "2 + 2 should equal 4"
End Sub

'Invalid Test Data definition
'@TestData
'@TestMethod
Public Sub TestAdditionParameterised01(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Invalid Test Data definition
'@TestData()
'@TestMethod
Public Sub TestAdditionParameterised02(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Malformed Test Data Sets
'@TestData({}
'@TestMethod
Public Sub TestAdditionParameterised03(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Malformed Test Data Sets
'@TestData(
'@TestMethod
Public Sub TestAdditionParameterised04(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Malformed Test Data Sets
'@TestData)
'@TestMethod
Public Sub TestAdditionParameterised05(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Malformed Test Data Sets
'@TestData)(
'@TestMethod
Public Sub TestAdditionParameterised06(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'No Test Data Sets
'@TestData()
'@TestMethod
Public Sub TestAdditionParameterised07(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Empty Test Data Sets
'@TestData({})
'@TestMethod
Public Sub TestAdditionParameterised08(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Empty Test Data
'@TestData({{}})
'@TestMethod
Public Sub TestAdditionParameterised09(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'1 Empty Test Data
'@TestData({{1, 2, 3}, {}})
'@TestMethod
Public Sub TestAdditionParameterised10(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Unmatched Braces
'@TestData({)
'@TestMethod
Public Sub TestAdditionParameterised11(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Unmatched Parenthesis
'@TestData({1, 2, 3}
'@TestMethod
Public Sub TestAdditionParameterised12(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Stray Comma - NB This gets ignored
'@TestData({{1, 2, 3},})
'@TestMethod
Public Sub TestAdditionParameterised13(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Stray Comma - NB last parameter ignored
'@TestData({{1, 2,}})
'@TestMethod
Public Sub TestAdditionParameterised14(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Parameter count mismatch - too few
'@TestData({{1, 2, 3}, {3, 4}})
'@TestMethod
Public Sub TestAdditionParameterised15(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Parameter count mismatch - too many
'@TestData({{1, 2, 3}, {3, 4, 7, 8}})
'@TestMethod
Public Sub TestAdditionParameterised16(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Max no of allowed parameters allowed 30 - limit of Application Run
'@TestData({{1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30}})
'@TestMethod
Public Sub TestAdditionParameterised17(p1 As Variant, p2 As Variant, p3 As Variant, p4 As Variant, p5 As Variant, p6 As Variant, p7 As Variant, p8 As Variant, p9 As Variant, p10 As Variant, p11 As Variant, p12 As Variant, p13 As Variant, p14 As Variant, p15 As Variant, p16 As Variant, p17 As Variant, p18 As Variant, p19 As Variant, p20 As Variant, p21 As Variant, p22 As Variant, p23 As Variant, p24 As Variant, p25 As Variant, p26 As Variant, p27 As Variant, p28 As Variant, p29 As Variant, p30 As Variant)
  Phosphorus.AssertionsStatic.pAssert.IsTrue (p1 + p2 + p3 + p4 + p5 + p6 + p7 + p8 + p9 + p10 + p11 + p12 + p13 + p14 + p15 + p16 + p17 + p18 + p19 + p20 + p21 + p22 + p23 + p24 + p25 + p26 + p27 + p28 + p29 + p30) > 1, "Total should be > 1"
End Sub

'More than max no of allowed parameters allowed 31 - limit of Application Run is 30
'@TestData({{1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31}})
'@TestMethod
Public Sub TestAdditionParameterised18(p1 As Variant, p2 As Variant, p3 As Variant, p4 As Variant, p5 As Variant, p6 As Variant, p7 As Variant, p8 As Variant, p9 As Variant, p10 As Variant, p11 As Variant, p12 As Variant, p13 As Variant, p14 As Variant, p15 As Variant, p16 As Variant, p17 As Variant, p18 As Variant, p19 As Variant, p20 As Variant, p21 As Variant, p22 As Variant, p23 As Variant, p24 As Variant, p25 As Variant, p26 As Variant, p27 As Variant, p28 As Variant, p29 As Variant, p30 As Variant, p31 As Variant)
  Phosphorus.AssertionsStatic.pAssert.IsTrue (p1 + p2 + p3 + p4 + p5 + p6 + p7 + p8 + p9 + p10 + p11 + p12 + p13 + p14 + p15 + p16 + p17 + p18 + p19 + p20 + p21 + p22 + p23 + p24 + p25 + p26 + p27 + p28 + p29 + p30 + p31) > 1, "Total should be > 1"
End Sub

'Invalid Test Data - data split over multiple lines fails - the data must all be on one line!
'@TestData({{1, 2, 3},
'           {3, 4, 7},
'           {0, 0, 0}})
'@TestMethod
Public Sub TestAdditionParameterised19(a As Variant, b As Variant, expected As Variant)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Parameter type mismatch
'Valid Test Data
'@TestData({{1, "Two", 3}})
'@TestMethod
'Public Sub TestAdditionParameterised20(a As Variant, b As Variant, expected As Variant)
Public Sub TestAdditionParameterised20(a As Integer, b As Integer, expected As Integer)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub

'Valid Test Data
'@TestData({{1, 2, "3", 1.5, FALSE, "TRUE"}})
'@TestMethod
Public Sub TestAdditionParameterised21(a As Variant, b As String, c As Integer, d As Long, e As Double, f As Boolean)
  Phosphorus.AssertionsStatic.pAssert.IsTrue True, True, "True Is True"
End Sub

'Q. How do we pass arrays as parameters
'@TestData({{{1, 2, 3}}})
'Public Sub TestAdditionParameterised98(g() As String)
'      Case "COLLECTION"
'      Case "DICTIONARY"
'      Case "ARRAY"
' Handle arrays (e.g., {1, 2, 3}), collections (e.g., {1, "two", 3}), and dictionaries (e.g., {key1=value1, key2=value2}) within @TestData.

'Valid Test Data
'@TestData({{1, 2, 3}, {3, 4, 7}, {0, 0, 0}})
'@TestMethod
Public Sub TestAdditionParameterised99(a As Integer, b As Integer, expected As Integer)
  Phosphorus.AssertionsStatic.pAssert.Equal a + b, expected, a & " + " & b & " should equal " & expected
End Sub


