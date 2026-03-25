Attribute VB_Name = "ErrorLogging"
'@Folder pEssence
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public Enum Errors
  FaultyEvaluationLogicUnspecifiedError = VBA.Constants.vbObjectError + 1001
  FaultyEvaluationLogicMismatchBracketsError
  FaultyEvaluationLogicConditionIsNotUsed
  FaultyEvaluationLogicUnspecifiedErrorOnEvaluation
  FaultyEvaluationLogicUnhandledPropertyComparison
  FindElementsExpectedOneElementFoundMany = VBA.Constants.vbObjectError + 2001
  FindElementsExpectedOneElementFoundNone
  FindElementsRootElementIsNothing
  FindElementsFindNoElementsBelowRoot
  FindElementsFindNoElements
  FindElementUnhandledByInLocator
  ElementIsNotAlive = VBA.Constants.vbObjectError + 3001
  UnhandledGetName = VBA.Constants.vbObjectError + 4001
  PatternFailedForElement = VBA.Constants.vbObjectError + 5001
End Enum

Public Sub LogError(ErrNumber As Errors, Description As String)
  Err.Raise VBA.Conversion.CLng(ErrNumber), , Description
End Sub
