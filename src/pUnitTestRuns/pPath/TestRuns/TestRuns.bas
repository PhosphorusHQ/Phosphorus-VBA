Attribute VB_Name = "TestRuns"
'@Folder TestRuns
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Sub RunAPPathTestBasic()
  pUnit.TestRunner.RunAllTests "pPathTestsBasic", "Test02Eval114", "pUnitTests_pPath"
End Sub

Sub RunAllPPathTestsBasic()
  pUnit.TestRunner.RunAllTests "pPathTestsBasic", "", "pUnitTests_pPath"
End Sub
'149 - 0 failed

Sub RunAllPPathPreValidationPPathTests()
  pUnit.TestRunner.RunAllTests "pPathTestsBasic", "PPathPreValidation", "pUnitTests_pPath"
End Sub
'15 - 0 failed

Sub RunAllPPathEvaluationPPathTests()
  pUnit.TestRunner.RunAllTests "pPathTestsBasic", "PPathEvaluation", "pUnitTests_pPath"
End Sub
'134 - 0 failed

Sub RunAPPathTestExcel()
  pUnit.TestRunner.RunAllTests "pPathTestsExcel", "Test06Excel103", "pUnitTests_pPath"
End Sub

Sub RunAllPPathTestExcel()
  pUnit.TestRunner.RunAllTests "pPathTestsExcel", "", "pUnitTests_pPath"
End Sub
'78 - 0 failed

Sub RunAllPPathTests()
  pUnit.TestRunner.RunAllTests "OR(pPathTestsBasic,pPathTestsExcel)", "", "pUnitTests_pPath"
End Sub
'149 + 78 = 227 Tests
