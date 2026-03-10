Attribute VB_Name = "ErrorLogging"
'@Folder pEssence
Option Explicit

Public Enum Errors
  FaultyEvaluationLogicUnspecifiedError = VBA.Constants.vbObjectError + 1001
  FaultyEvaluationLogicMismatchBracketsError = VBA.Constants.vbObjectError + 1002
  FaultyEvaluationLogicUnspecifiedErrorOnEvaluation = VBA.Constants.vbObjectError + 1003
  FindElementsExpectedOneElementFoundMany = VBA.Constants.vbObjectError + 2001
  FindElementsExpectedOneElementFoundNone = VBA.Constants.vbObjectError + 2002
  FindElementsRootElementIsNothing = VBA.Constants.vbObjectError + 2003
  FindElementsFindNoElements = VBA.Constants.vbObjectError + 2004
  FindElementUnhandledByInLocator = VBA.Constants.vbObjectError + 2005
  PatternFailedForElement = VBA.Constants.vbObjectError + 3001
End Enum

Public Sub LogError(ErrNumber As Errors, Description As String)
  Err.Raise VBA.Conversion.CLng(ErrNumber), , Description
End Sub
