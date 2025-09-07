Attribute VB_Name = "pPathTestRuns"
'@Folder pUnit
Option Explicit

Sub RunAPPathTestBasic()
  pUnit.TestRunner.RunAllTests "PPathTestsBasic", "Test02Eval114", "pUnitTests"
End Sub

Sub RunAAllPPathTestsBasic()
  pUnit.TestRunner.RunAllTests "PPathTestsBasic", "", "pUnitTests"
End Sub
'149 - 0 failed

Sub RunAllPPathPreValidationPPathTests()
  pUnit.TestRunner.RunAllTests "PPathTestsBasic", "PPathPreValidation", "pUnitTests"
End Sub
'15 - 0 failed

Sub RunAllPPathEvaluationPPathTests()
  pUnit.TestRunner.RunAllTests "PPathTestsBasic", "PPathEvaluation", "pUnitTests"
End Sub
'134 - 0 failed

Sub RunAPPathTestExcel()
  pUnit.TestRunner.RunAllTests "PPathTestsExcel", "Test06Excel104", "pUnitTests"
End Sub

Sub RunAllPPathTestExcel()
  pUnit.TestRunner.RunAllTests "PPathTestsExcel", "", "pUnitTests"
End Sub
'78 - 0 failed

Sub RunAllPPathTests()
  pUnit.TestRunner.RunAllTests "OR(PPathTestsBasic,PPathTestsExcel)", "", "pUnitTests"
End Sub
'149 + 78 = 227 Tests
