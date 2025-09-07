Attribute VB_Name = "PPathTestsBasic"
'@Folder PPath
'@TestModule

'https://www.bernardvukas.com/testing/tutorial-excel-vba-unit-testing/

Option Explicit
Option Private Module
Option Base 1

Private strTestPPath As String
Private PPath As Phosphorus.PPath
Private EvaluatedPPath As Phosphorus.PPathReturnClass
Private InterimPPath As Phosphorus.PPath
Private llongExpectedNumberOfMatchingElements As Long
Private InterimEvaluatedPPath1 As Phosphorus.PPathReturnClass
Private InterimEvaluatedPPath2 As Phosphorus.PPathReturnClass
Private InterimEvaluatedPPath3 As Phosphorus.PPathReturnClass

'Run all PPath test on the current Excel instance
Private eleExcelRootElement As UIAutomationClient.IUIAutomationElement

Private Sub BeforeModule()
  On Error GoTo ErrorHandler
  'Make sure Excel Ribbon is on the Home tab
  Phosphorus.OfficeRibbon.ActivateHomeTab
  Set eleExcelRootElement = FindExcelRootElement
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Private Sub AfterModule()
  'this method runs once per module.
  On Error GoTo ErrorHandler

'  Set Assert = Nothing
'  Set Fakes = Nothing
  Phosphorus.Utils.CloseAllOtherWorkbooks
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Private Sub BeforeTest()
  On Error GoTo ErrorHandler
  'This method runs before every test in the module..
  Set PPath = Nothing
  Set PPath = Phosphorus.Factory.GetNewPhosphorusPPath
  Set InterimPPath = Nothing
  Set InterimPPath = Phosphorus.Factory.GetNewPhosphorusPPath
  PPath.Initialise
'PJG
'PPath.SetDebugMode = True
  PPath.SetApplicationRootElement eleExcelRootElement
  InterimPPath.Initialise
  InterimPPath.SetApplicationRootElement eleExcelRootElement
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Private Sub AfterTest()
  On Error GoTo ErrorHandler
  'this method runs after every test in the module.
  If PPath.GetDebugMode Then
    PPathTestsCommon.OutputActualXPaths EvaluatedPPath.GetMatchingNavigationalPPaths
  End If
  Set InterimEvaluatedPPath1 = Nothing
  Set InterimEvaluatedPPath2 = Nothing
  Set InterimEvaluatedPPath3 = Nothing
  Set EvaluatedPPath = Nothing
  Set PPath = Nothing
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Private Function FindExcelRootElement() As UIAutomationClient.IUIAutomationElement
  Dim oAutomation As CUIAutomation
  Set oAutomation = New CUIAutomation
  Dim rootElement As UIAutomationClient.IUIAutomationElement
  Set rootElement = oAutomation.GetRootElement
  Set rootElement = rootElement.FindFirst( _
    UIAutomationClient.TreeScope.TreeScope_Children, _
    oAutomation.CreateAndCondition( _
      oAutomation.CreatePropertyCondition(UIAutomationClient.UIA_PropertyIds.UIA_ControlTypePropertyId, UIAutomationClient.UIA_ControlTypeIds.UIA_WindowControlTypeId), _
      oAutomation.CreatePropertyCondition(UIAutomationClient.UIA_PropertyIds.UIA_NamePropertyId, "Excel") _
    ) _
  )
  Set FindExcelRootElement = rootElement
End Function

'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre001()
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "((((()))"
Act:
 Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.MISMATCHING_PARENTHESES_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre002()
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "[[[]]]]]"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.MISMATCHING_SQUARE_BARCKETS_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements, isCritical:=True"
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre003()
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "@NotAUIProperty"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.INVALID_PROPERTY_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre004()
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "./..*"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.MISSING_RELATIVE_PPATH_CONTEXT_NODE, EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre005()
  On Error GoTo ErrorHandler
  
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  strTestPPath = "/Edit"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.UNUSED_RELATIVE_PPATH_CONTEXT_NODE, EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Union of expressions, 3 context nodes, 3 uses
'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre006()
  On Error GoTo ErrorHandler
  
Arrange:
'  PPath.SetDebugMode = True

  strTestPPath = "//Tab"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "//ToolBar"
  Set InterimEvaluatedPPath2 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 3
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath2.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "//SplitButton"
  Set InterimEvaluatedPPath3 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 13
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath3.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = ".//* | .//* | .//*"
  llongExpectedNumberOfMatchingElements = 20
Act:
'  PPath.SetDebugMode = True
  PPath.AddContextNode InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/", True
  PPath.AddContextNode InterimEvaluatedPPath2.GetMatchingElement(1), InterimEvaluatedPPath2.GetMatchingNavigationalPPath(1) & "/"
  'NB Context Node 3 not set!
  'PPath.AddContextNode EvaluatedPPath3.GetMatchingElement(1), EvaluatedPPath2.GetMatchingNavigationalPPath(1) & "/", True
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.NUMBER_OF_RELATIVE_PPATHS_TO_CONTEXT_NODES_MISMATCH, EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre007()
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "([)]"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "Unexpected ')' Bracket at position 3", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre008()
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "(()])["
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "Unexpected ']' Bracket at position 4", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre009()
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "]["
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "Unexpected ']' Bracket at position 1", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre010()
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "[)]("
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "Unexpected ')' Bracket at position 2", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre011()
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "//*[And(count(./Button)>=1,count(./SplitButton)>1,count(./ComboBox)>1)]/Button[position)(=6]"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "Unexpected ')' Bracket at position 88", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre012()
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "/button"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.INVALID_NODETEST_PPATH_ERROR_MESSAGE & " 'button'!", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre013()
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "/abc"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.INVALID_NODETEST_PPATH_ERROR_MESSAGE & " 'abc'!", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre014()
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "//Button/attribute::isenabled"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.INVALID_NODETEST_PPATH_ERROR_MESSAGE & " 'isenabled'!", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathPreValidation)
'@TestMethod
Private Sub Test01Pre015()
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "//Button/attribute::NotANavigableAttribute"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.INVALID_NODETEST_PPATH_ERROR_MESSAGE & " 'NotANavigableAttribute'!", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval001()
  On Error GoTo ErrorHandler
  
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = ""
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.NULL_PPATH_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test001, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval002()
  On Error GoTo ErrorHandler
  
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "( | )"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "PPath #1: " & Phosphorus.PPathConstants.NULL_PPATH_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test002, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval003()
  On Error GoTo ErrorHandler
  
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.NO_NODETEST_PPATH_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test003, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval004() 'Ignore spaces at start and end of PPath
  On Error GoTo ErrorHandler
  
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "    /    "
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.NO_NODETEST_PPATH_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test004, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval005()
  On Error GoTo ErrorHandler
  
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/*"
  llongExpectedNumberOfMatchingElements = 7
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test005, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval006() 'Unrecognised predicate
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "/*a"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.ILLEGAL_START_OF_PREDICATE_ERROR_MESSAGE & " 'a'", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test006, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval007()
  On Error GoTo ErrorHandler
  
Arrange:
  strTestPPath = "( /* | / )"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.NO_NODETEST_PPATH_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test007, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval008()
  On Error GoTo ErrorHandler
  
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(/* | /*)"
  llongExpectedNumberOfMatchingElements = 7
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test008, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval009()
  On Error GoTo ErrorHandler
  
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*"
  llongExpectedNumberOfMatchingElements = 155
Act:
   Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test009, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval010()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//* | //*"
  llongExpectedNumberOfMatchingElements = 155
Act:
   Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test010, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval011()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane"
  llongExpectedNumberOfMatchingElements = 5
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test011, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval012()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "( /Pane | /Pane )"
  llongExpectedNumberOfMatchingElements = 5
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test012, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval013()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem"
  llongExpectedNumberOfMatchingElements = 28
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test013, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval014()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/child::*"
  llongExpectedNumberOfMatchingElements = 7
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test014, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval015()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/descendant::*"
  llongExpectedNumberOfMatchingElements = 155
Act:
   Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test015, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Self Node shorthand
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval016()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/.*"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test016, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Self Node
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval017()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/self::*"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test017, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Parent shorthand - root element
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval018()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/..*"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test018, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal False, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Parent shorthand
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval019()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"

  strTestPPath = "./..*"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test019, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Parent
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval020()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./parent::*"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test020, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Ancestor
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval021()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//StatusBar"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./ancestor::*"
  llongExpectedNumberOfMatchingElements = 6
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test021, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'AncestorOrSelf
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval022()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True

  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./ancestor-or-self::*"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test022, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval023()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/descendant-or-self::*"
  llongExpectedNumberOfMatchingElements = 156
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test023, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'PrecedingSibling
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval024()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./preceding-sibling::*"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test024, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'FollowingSibling
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval025()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./following-sibling::*"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test025, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Preceding - Nothing before first element test
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval026()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 5
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./preceding::*"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test026, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal False, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Preceding - Lots before last child root element test
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval027()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/TitleBar"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./preceding::*"
  llongExpectedNumberOfMatchingElements = 149
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test027, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Preceding - Lots before middle low level element test
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval028()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 28
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./preceding::*"
  llongExpectedNumberOfMatchingElements = 23
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test028, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Following - Nothing after last root element test
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval029()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/TitleBar"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./following::*"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test029, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal False, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Following - Lots after the first child root element test
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval030()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 5
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./following::*"
  llongExpectedNumberOfMatchingElements = 154
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test030, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Following - Lots after middle low level element test
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval031()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 28
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./following::*"
  llongExpectedNumberOfMatchingElements = 131
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test031, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval032()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/child::Pane"
  llongExpectedNumberOfMatchingElements = 5
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test032, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval033()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/descendant::Button"
  llongExpectedNumberOfMatchingElements = 58
Act:
   Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test033, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Self Node shorthand
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval034()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/.Window"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test034, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Self Node
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval035()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/self::Window"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
 Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
 Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test035, EvaluatedPPath
 Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Parent shorthand - root element
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval036()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/..Window"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
 Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
 Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test036, EvaluatedPPath
 Phosphorus.AssertionsStatic.pAssert.Equal False, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Parent shorthand
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval037()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"

  strTestPPath = "./..Window"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test037, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Parent
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval038()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./parent::*"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test038, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Ancestor
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval039()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//StatusBar"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./ancestor::ToolBar"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test039, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'AncestorOrSelf
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval040()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./ancestor-or-self::Window"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test040, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval041()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 5
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./descendant-or-self::Pane"
  llongExpectedNumberOfMatchingElements = 6
Act:
  'NOTE: The third pane is the first one with children
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(3), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(3) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test041, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'PrecedingSibling
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval042()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//SplitButton"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 13
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./preceding-sibling::Button"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test042, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'FollowingSibling
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval043()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 58
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./following-sibling::Button"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test043, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval044()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//SplitButton"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 13
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  strTestPPath = "./preceding::Button"
  llongExpectedNumberOfMatchingElements = 6
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test044, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval045()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 58
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./following::ComboBox"
  llongExpectedNumberOfMatchingElements = 4
'  PPath.SetDebugMode = True
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test045, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Attribute
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval046()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 58
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./attribute::*"
  llongExpectedNumberOfMatchingElements = 31
'  PPath.SetDebugMode = True
Act:
  'Need to run this test is TestMode as some parts of elements, e.g. PID are dynamic!
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/", True)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test046, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Attribute
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval047()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 58
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./attribute::IsEnabled"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test047, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Node()
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval048()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 4
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./node()"
  llongExpectedNumberOfMatchingElements = 65
Act:
  'Need to run this test is TestMode as some parts of elements, e.g. PID are dynamic!
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/", True)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test048, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Element()
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval049()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 4
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./element()"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test049, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Text()
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval050()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 4
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./text()"
  llongExpectedNumberOfMatchingElements = 2
Act:
  'Need to run this test is TestMode as some parts of elements, e.g. PID are dynamic!
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/", True)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test050, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Text()
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval051()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 155
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "./text()"
  llongExpectedNumberOfMatchingElements = 0
Act:
  'No Excel UIElements have any text
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test051, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal False, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Element() by kind - integer
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval052()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//element(*, integer)"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test052, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal False, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Element() by kind - string
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval053()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//element(*, string)"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test053, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Nested steps
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval054()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/TitleBar/*"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test054, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Nested steps
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval055()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//SplitButton/Button"
  llongExpectedNumberOfMatchingElements = 13
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test055, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Nested steps
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval056()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ToolBar/attribute::*"
  llongExpectedNumberOfMatchingElements = 90
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, , , True)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test056, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Nested steps
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval057()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/TitleBar/node()/attribute::AutomationId"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, , , True)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test057, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Union of expressions, no context node
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval058()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane | /Edit"
  llongExpectedNumberOfMatchingElements = 6
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, , , True)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test058, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Union of expressions, context node, 1 use
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval059()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Tab"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "/Pane | .//*"
  llongExpectedNumberOfMatchingElements = 21
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test059, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Union of expressions, 1 context node, 2 uses
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval060()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ToolBar"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 3
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = ".//ToolBar | .//MenuItem"
  llongExpectedNumberOfMatchingElements = 28
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(2), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(2) & "/")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test060, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Missing Context mode initial PPath
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval061()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//StatusBar"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  'strInitialPPaths(1) = EvaluatedPPath.GetMatchingNavigationalPPath(1) & "/" 'NOT USED!!
  strTestPPath = "./ancestor::*"
Act:
  'PPath not set!
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), "")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.MISSING_CONTEXT_NODE_INITIAL_PPATH, EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Union of expressions, 3 context nodes, 3 uses
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval062()
  On Error GoTo ErrorHandler

Arrange:
'   PPath.SetDebugMode = True
  Dim strInitialPPaths(3) As String

  strTestPPath = "//Tab"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "//ToolBar"
  Set InterimEvaluatedPPath2 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 3
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath2.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = "//SplitButton"
  Set InterimEvaluatedPPath3 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 13
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath3.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True

  strTestPPath = ".//* | .//* | .//*"
  llongExpectedNumberOfMatchingElements = 27
Act:
'  PPath.SetDebugMode = True
  PPath.AddContextNode InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/", True
  PPath.AddContextNode InterimEvaluatedPPath2.GetMatchingElement(1), InterimEvaluatedPPath2.GetMatchingNavigationalPPath(1) & "/"
  PPath.AddContextNode InterimEvaluatedPPath3.GetMatchingElement(1), InterimEvaluatedPPath3.GetMatchingNavigationalPPath(1) & "/"
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test062, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval063()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane[2]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test063, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval064()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Button)[1]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test064, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval065()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button[2]"
  llongExpectedNumberOfMatchingElements = 9
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test065, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval066()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Button[2])[4]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test066, EvaluatedPPath
 Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval067()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane[position()=2]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test067, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval068()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane[first()]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test068, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval069()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Button[2])[first()]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test069, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval070()
  On Error GoTo ErrorHandler

Arrange:
'   PPath.SetDebugMode = True
  strTestPPath = "(//Button[2])[first()+1]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test070, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval071()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(/Pane)[last()]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test071, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval072()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button[last()]"
  llongExpectedNumberOfMatchingElements = 32
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test072, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval073()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button[last()-1]"
  llongExpectedNumberOfMatchingElements = 9
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test073, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval074()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button[And(position()>1,position()<last())]"
  llongExpectedNumberOfMatchingElements = 17
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test074, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval075()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Button[Or(position()=1,position()=last())])[And(position()>1,Or(position()=1,position()=last()-1))]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test075, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval076()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "count(//StatusBar/Button)"
  llongExpectedNumberOfMatchingElements = 4
Act:
   Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test076, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, VBA.Conversion.CLng(EvaluatedPPath.ReturnedValue), "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval077()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//StatusBar[count(./Button)=4]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test077, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Missing "." in nested ppath
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval078()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//StatusBar[count(/Button)=6]"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.UNUSED_RELATIVE_PPATH_CONTEXT_NODE, EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Invalid funciton name in nested ppath
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval079()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//StatusBar[notafunction(/Button)=6]"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal Phosphorus.PPathConstants.INVALID_PREDICATE & " [notafunction(/Button)=6] => notafunction(/Button)=6", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval080()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[And(count(./Button)>=1,count(./SplitButton)>1,count(./ComboBox)>1)]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test080, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval081()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Tab[not(./Button)]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test081, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Test 3 step levels
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval082()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane/ComboBox/Edit"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test082, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval083()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane[count(./ComboBox)=1]/ComboBox/Edit[1]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test083, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval084()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem[./preceding-sibling::Button]"
  llongExpectedNumberOfMatchingElements = 16
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test084, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval085()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem[./following-sibling::Button]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test085, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval086()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem[Or(./preceding-sibling::Button, ./following-sibling::Button)]"
  llongExpectedNumberOfMatchingElements = 18
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test086, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval087()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ToolBar[@Name=" & VBA.Strings.Chr(34) & "Quick Access Toolbar" & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test087, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval088()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button[@AcceleratorKey=" & VBA.Strings.Chr(34) & "Ctrl+S" & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test088, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval089()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button[@AccessKey=" & VBA.Strings.Chr(34) & "Alt+Down Arrow" & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test089, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval090()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Text[@AriaProperties=" & VBA.Strings.Chr(34) & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test090, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval091()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@AriaRole=" & VBA.Strings.Chr(34) & "" & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test091, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval092()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[@AutomationId=" & VBA.Strings.Chr(34) & "AutoSaveSwitch" & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test092, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval093()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//TabItem[@BoundingRectangle=" & VBA.Strings.Chr(34) & "#" & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 14
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, , , True)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test093, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval094()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@CenterPoint=""""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test094, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval095()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[@ClassName=" & VBA.Strings.Chr(34) & "DropShadow" & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test095, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval096()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@Culture=0]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test096, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval097()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@FillColor=""""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test097, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval098()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@FillType=0]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test098, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval099()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@FrameworkId=""Win32""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test099, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval100()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  'Use //* to test for not matching long descriptions
  strTestPPath = "//*[@FullDescription=""Pick a new font for your text.""]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test100, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval101()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button[@FullDescription=""Like the look of a particular selection? You can apply that look to other content in the document.\n\nTo get started: \n1. Select content with the formatting you like\n2. Click Format Painter\n3. Select something else to automatically apply the formatting\n\nFYI: To apply the formatting in multiple places, double-click Format Painter.""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test101, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval102()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane[@HasKeyboardFocus=False]"
  llongExpectedNumberOfMatchingElements = 5
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test102, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval103()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ToolBar[@HeadingLevel=80050]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test103, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval104()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane/ToolBar[@HelpText=""Ribbon toolbar""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test104, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval105()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Tab[@IsContentElement=True]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test105, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval106()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/TitleBar//MenuItem[@IsControlElement=True]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test106, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval107()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/TitleBar//MenuItem[@IsDataValidForForm=False]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test107, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval108()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/TitleBar//MenuItem[@IsDialog=False]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test108, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval109()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@IsEnabled=False]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test109, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval110()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Text[@IsKeyboardFocusable=True]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test110, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval111()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[@IsOffscreen=True]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test111, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval112()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Text[@IsPassword=False]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test112, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval113()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit[@IsPeripheral=False]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test113, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval114()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Text[@IsRequiredForForm=False]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test114, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval115()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[@ItemStatus<>""""]"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test115, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal False, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval116()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Edit[@ItemType=""Edit Formula""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test116, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval117()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit[@LandmarkType=0]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test117, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval118()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit[@Level=0]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test118, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval119()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@LiveSetting=2]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test119, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval120()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@LocalizedControlType=""ComboBox""]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test120, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval121()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@LocalizedLandmarkType=""""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test121, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval122()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[@Name=""Macro Recording Not Recording""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test122, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval123()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@NativeWindowHandle=""#""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, , , True)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test123, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval124()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@OptimizeForVisualContent=False]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test124, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval125()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@Orientation=1]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test125, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval126()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[@PositionInSet=13]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test126, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval127()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@ProcessId=""#""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, , , True)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test127, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval128()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@ProviderDescription=""#""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, , , True)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test128, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval129()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@Rotation=""""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test129, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval130()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[@SizeOfSet=2]"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test130, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval131()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@VisualEffects=0]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test131, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval132()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem[position()<=2][1]"
  llongExpectedNumberOfMatchingElements = 22
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test132, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval133()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem[And(./preceding-sibling::Button,./following-sibling::SplitButton)]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test133, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Variation of #133
'@Tag(PPathEvaluation)
'@TestMethod
Private Sub Test02Eval134()
  On Error GoTo ErrorHandler

Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem[./preceding-sibling::Button][./following-sibling::SplitButton]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  PPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests.PPathTestsExpectedResults.Test134, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub
