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

'@ModuleInitialize
Private Sub ModuleInitialize()
  Set Assert = CreateObject("Rubberduck.AssertClass")
  Set Fakes = CreateObject("Rubberduck.FakesProvider")
  Set eleExcelRootElement = FindExcelRootElement
  Log4PStatic.GetLogger
  Logger.SetTempLevel LogLevel.INTERNAL_INFO
  Logger.InternalInfo "Logger Started"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
  'this method runs once per module.
  Set Assert = Nothing
  Set Fakes = Nothing
  Utils.CloseAllOtherWorkbooks
  Logger.InternalInfo "Logger Stopped"
  Set Logger = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
  'This method runs before every test in the module..
  Set PPath = Nothing
  Set PPath = Phosphorus.Factory.GetNewPhosphorusPPath
  Set InterimPPath = Nothing
  Set InterimPPath = Phosphorus.Factory.GetNewPhosphorusPPath
  PPath.Initialise
  PPath.SetApplicationRootElement eleExcelRootElement
  InterimPPath.Initialise
  InterimPPath.SetApplicationRootElement eleExcelRootElement
End Sub

'@TestCleanup
Private Sub TestCleanup()
  'this method runs after every test in the module.
  If PPath.GetDebugMode Then
    OutputActualXPaths EvaluatedPPath.GetMatchingNavigationalPPaths
  End If
  Set InterimEvaluatedPPath1 = Nothing
  Set InterimEvaluatedPPath2 = Nothing
  Set InterimEvaluatedPPath3 = Nothing
  Set EvaluatedPPath = Nothing
  Set PPath = Nothing
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

'@TestMethod("PPath PreValidation")
Private Sub Test01Pre001()
Arrange:
  strTestPPath = "((((()))"
Act:
 Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.MISMATCHING_PARENTHESES_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath PreValidation")
Private Sub Test01Pre002()
Arrange:
  strTestPPath = "[[[]]]]]"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.MISMATCHING_SQUARE_BARCKETS_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath PreValidation")
Private Sub Test01Pre003()
Arrange:
  strTestPPath = "@NotAUIProperty"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.INVALID_PROPERTY_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath PreValidation")
Private Sub Test01Pre004()
  strTestPPath = "./..*"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.MISSING_RELATIVE_PPATH_CONTEXT_NODE, EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath PreValidation")
Private Sub Test01Pre005()
'  PPath.SetDebugMode = True
  
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "/Edit"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.UNUSED_RELATIVE_PPATH_CONTEXT_NODE, EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Union of expressions, 3 context nodes, 3 uses
'@TestMethod("PPath PreValidation")
Private Sub Test01Pre006()
Arrange:
'  PPath.SetDebugMode = True
  
  strTestPPath = "//Tab"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "//ToolBar"
  Set InterimEvaluatedPPath2 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 3
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath2.GetFinalNumberOfMatchingElements, "GetMatchingElements"

  strTestPPath = "//SplitButton"
  Set InterimEvaluatedPPath3 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 13
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath3.GetFinalNumberOfMatchingElements, "GetMatchingElements"

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
  Assert.AreEqual Phosphorus.PPathConstants.NUMBER_OF_RELATIVE_PPATHS_TO_CONTEXT_NODES_MISMATCH, EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath PreValidation")
Private Sub Test01Pre007()
Arrange:
  strTestPPath = "([)]"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "Unexpected ')' Bracket at position 3", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath PreValidation")
Private Sub Test01Pre008()
Arrange:
  strTestPPath = "(()])["
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "Unexpected ']' Bracket at position 4", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath PreValidation")
Private Sub Test01Pre009()
Arrange:
  strTestPPath = "]["
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "Unexpected ']' Bracket at position 1", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath PreValidation")
Private Sub Test01Pre010()
Arrange:
  strTestPPath = "[)]("
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "Unexpected ')' Bracket at position 2", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath PreValidation")
Private Sub Test01Pre011()
Arrange:
  strTestPPath = "//*[And(count(./Button)>=1,count(./SplitButton)>1,count(./ComboBox)>1)]/Button[position)(=6]"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "Unexpected ')' Bracket at position 88", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath PreValidation")
Private Sub Test01Pre012()
Arrange:
  strTestPPath = "/button"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.INVALID_NODETEST_PPATH_ERROR_MESSAGE & " 'button'!", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath PreValidation")
Private Sub Test01Pre013()
Arrange:
  strTestPPath = "/abc"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.INVALID_NODETEST_PPATH_ERROR_MESSAGE & " 'abc'!", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath PreValidation")
Private Sub Test01Pre014()
Arrange:
  strTestPPath = "//Button/attribute::isenabled"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.INVALID_NODETEST_PPATH_ERROR_MESSAGE & " 'isenabled'!", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath PreValidation")
Private Sub Test01Pre015()
Arrange:
  strTestPPath = "//Button/attribute::NotANavigableAttribute"
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.INVALID_NODETEST_PPATH_ERROR_MESSAGE & " 'NotANavigableAttribute'!", EvaluatedPPath.GetErrorMessage, "PreValidationErrorMessages"
  Assert.AreEqual VBA.Conversion.CLng(0), EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.TestPre, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval001()
Arrange:
  strTestPPath = ""
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.NULL_PPATH_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test001, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval002()
Arrange:
  strTestPPath = "( | )"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "PPath #1: " & Phosphorus.PPathConstants.NULL_PPATH_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test002, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval003()
Arrange:
  strTestPPath = "/"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.NO_NODETEST_PPATH_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test003, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval004() 'Ignore spaces at start and end of PPath
Arrange:
  strTestPPath = "    /    "
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.NO_NODETEST_PPATH_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test004, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval005()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/*"
  llongExpectedNumberOfMatchingElements = 7
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test005, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval006() 'Unrecognised predicate
Arrange:
  strTestPPath = "/*a"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.ILLEGAL_START_OF_PREDICATE_ERROR_MESSAGE & " 'a'", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test006, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'TODO Test 7
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval007()
Arrange:
  strTestPPath = "( /* | / )"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.NO_NODETEST_PPATH_ERROR_MESSAGE, EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test007, EvaluatedPPath
  Assert.AreEqual "", EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval008()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(/* | /*)"
  llongExpectedNumberOfMatchingElements = 7
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test008, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval009()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*"
  llongExpectedNumberOfMatchingElements = 163
Act:
   Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test009, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub
 
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval010()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//* | //*"
  llongExpectedNumberOfMatchingElements = 163
Act:
   Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test010, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval011()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane"
  llongExpectedNumberOfMatchingElements = 5
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test011, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval012()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "( /Pane | /Pane )"
  llongExpectedNumberOfMatchingElements = 5
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test012, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval013()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem"
  llongExpectedNumberOfMatchingElements = 28
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test013, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval014()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/child::*"
  llongExpectedNumberOfMatchingElements = 7
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test014, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval015()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/descendant::*"
  llongExpectedNumberOfMatchingElements = 163
Act:
   Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test015, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Self Node shorthand
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval016()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/.*"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test016, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Self Node
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval017()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/self::*"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test017, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Parent shorthand - root element
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval018()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/..*"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test018, EvaluatedPPath
  Assert.AreEqual False, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Parent shorthand
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval019()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./..*"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test019, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Parent
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval020()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"

  strTestPPath = "./parent::*"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test020, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Ancestor
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval021()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//StatusBar"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
    
  strTestPPath = "./ancestor::*"
  llongExpectedNumberOfMatchingElements = 6
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test021, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'AncestorOrSelf
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval022()
Arrange:
'  PPath.SetDebugMode = True
  
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
    
  strTestPPath = "./ancestor-or-self::*"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test022, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval023()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/descendant-or-self::*"
  llongExpectedNumberOfMatchingElements = 164
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test023, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'PrecedingSibling
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval024()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./preceding-sibling::*"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test024, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'FollowingSibling
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval025()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./following-sibling::*"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test025, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Preceding - Nothing before first element test
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval026()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 5
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
    
  strTestPPath = "./preceding::*"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test026, EvaluatedPPath
  Assert.AreEqual False, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Preceding - Lots before last child root element test
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval027()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/TitleBar"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./preceding::*"
  llongExpectedNumberOfMatchingElements = 157
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test027, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Preceding - Lots before middle low level element test
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval028()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 28
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./preceding::*"
  llongExpectedNumberOfMatchingElements = 29
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test028, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Following - Nothing after last root element test
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval029()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/TitleBar"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./following::*"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test029, EvaluatedPPath
  Assert.AreEqual False, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Following - Lots after the first child root element test
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval030()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 5
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./following::*"
  llongExpectedNumberOfMatchingElements = 162
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test030, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Following - Lots after middle low level element test
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval031()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 28
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./following::*"
  llongExpectedNumberOfMatchingElements = 133
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test031, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval032()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/child::Pane"
  llongExpectedNumberOfMatchingElements = 5
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test032, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval033()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/descendant::Button"
  llongExpectedNumberOfMatchingElements = 64
Act:
   Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test033, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Self Node shorthand
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval034()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/.Window"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test034, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Self Node
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval035()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/self::Window"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test035, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Parent shorthand - root element
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval036()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/..Window"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test036, EvaluatedPPath
  Assert.AreEqual False, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Parent shorthand
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval037()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./..Window"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test037, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Parent
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval038()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./parent::*"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test038, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Ancestor
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval039()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//StatusBar"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./ancestor::ToolBar"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test039, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'AncestorOrSelf
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval040()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./ancestor-or-self::Window"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test040, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval041()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 5
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
    
  strTestPPath = "./descendant-or-self::Pane"
  llongExpectedNumberOfMatchingElements = 6
Act:
  'NOTE: The third pane is the first one with children
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(3), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(3) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test041, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'PrecedingSibling
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval042()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//SplitButton"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 13
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./preceding-sibling::Button"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test042, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'FollowingSibling
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval043()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 64
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./following-sibling::Slider"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test043, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval044()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Slider"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  strTestPPath = "./preceding::Button"
  llongExpectedNumberOfMatchingElements = 5
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test044, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval045()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 64
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./following::ComboBox"
  llongExpectedNumberOfMatchingElements = 4
'  PPath.SetDebugMode = True
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test045, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Attribute
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval046()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 64
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./attribute::*"
  llongExpectedNumberOfMatchingElements = 31
'  PPath.SetDebugMode = True
Act:
  'Need to run this test is TestMode as some parts of elements, e.g. PID are dynamic!
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/", True)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test046, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Attribute
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval047()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 64
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./attribute::IsEnabled"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test047, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Node()
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval048()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Slider"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./node()"
  llongExpectedNumberOfMatchingElements = 64
Act:
  'Need to run this test is TestMode as some parts of elements, e.g. PID are dynamic!
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/", True)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test048, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Element()
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval049()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Slider"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./element()"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test049, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Text()
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval050()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Slider"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./text()"
  llongExpectedNumberOfMatchingElements = 2
Act:
  'Need to run this test is TestMode as some parts of elements, e.g. PID are dynamic!
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/", True)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test050, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Text()
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval051()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 163
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "./text()"
  llongExpectedNumberOfMatchingElements = 0
Act:
  'No Excel UIElements have any text
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test051, EvaluatedPPath
  Assert.AreEqual False, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Element() by kind - integer
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval052()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//element(*, integer)"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test052, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Element() by kind - string
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval053()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//element(*, string)"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test053, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Nested steps
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval054()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/TitleBar/*"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test054, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Nested steps
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval055()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//SplitButton/Button"
  llongExpectedNumberOfMatchingElements = 13
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test055, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Nested steps
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval056()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ToolBar/attribute::*"
  llongExpectedNumberOfMatchingElements = 90
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, , , True)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test056, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Nested steps
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval057()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/TitleBar/node()/attribute::AutomationId"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, , , True)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test057, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Union of expressions, no context node
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval058()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane | /Edit"
  llongExpectedNumberOfMatchingElements = 6
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, , , True)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test058, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Union of expressions, context node, 1 use
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval059()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Tab"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "/Pane | .//*"
  llongExpectedNumberOfMatchingElements = 20
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test059, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Union of expressions, 1 context node, 2 uses
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval060()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ToolBar"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 3
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = ".//ToolBar | .//MenuItem"
  llongExpectedNumberOfMatchingElements = 28
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(2), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(2) & "/")
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test060, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Missing Context mode initial PPath
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval061()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//StatusBar"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  'strInitialPPaths(1) = EvaluatedPPath.GetMatchingNavigationalPPath(1) & "/" 'NOT USED!!
  strTestPPath = "./ancestor::*"
Act:
  'PPath not set!
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, InterimEvaluatedPPath1.GetMatchingElement(1), "")
Assert:
  Assert.AreEqual Phosphorus.PPathConstants.MISSING_CONTEXT_NODE_INITIAL_PPATH, EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
End Sub

'Union of expressions, 3 context nodes, 3 uses
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval062()
Arrange:
'  PPath.SetDebugMode = True
  Dim strInitialPPaths(3) As String

  strTestPPath = "//Tab"
  Set InterimEvaluatedPPath1 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 1
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath1.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "//ToolBar"
  Set InterimEvaluatedPPath2 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 3
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath2.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = "//SplitButton"
  Set InterimEvaluatedPPath3 = InterimPPath.Evaluate(strTestPPath)
  llongExpectedNumberOfMatchingElements = 13
  Assert.AreEqual llongExpectedNumberOfMatchingElements, InterimEvaluatedPPath3.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  
  strTestPPath = ".//* | .//* | .//*"
  llongExpectedNumberOfMatchingElements = 32
Act:
'  PPath.SetDebugMode = True
  PPath.AddContextNode InterimEvaluatedPPath1.GetMatchingElement(1), InterimEvaluatedPPath1.GetMatchingNavigationalPPath(1) & "/", True
  PPath.AddContextNode InterimEvaluatedPPath2.GetMatchingElement(1), InterimEvaluatedPPath2.GetMatchingNavigationalPPath(1) & "/"
  PPath.AddContextNode InterimEvaluatedPPath3.GetMatchingElement(1), InterimEvaluatedPPath3.GetMatchingNavigationalPPath(1) & "/"
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test062, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval063()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane[2]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test063, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval064()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Button)[1]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test064, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval065()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button[2]"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test065, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval066()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Button[2])[4]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test066, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval067()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane[position()=2]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test067, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval068()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane[first()]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test068, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval069()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Button[2])[first()]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test069, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval070()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Button[2])[first()+1]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test070, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval071()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(/Pane)[last()]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test071, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval072()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button[last()]"
  llongExpectedNumberOfMatchingElements = 35
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test072, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval073()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button[last()-1]"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test073, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval074()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button[And(position()>1,position()<last())]"
  llongExpectedNumberOfMatchingElements = 19
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test074, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval075()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Button[Or(position()=1,position()=last())])[And(position()>1,Or(position()=1,position()=last()-1))]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test075, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval076()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "count(//StatusBar/Button)"
  llongExpectedNumberOfMatchingElements = 6
Act:
   Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test076, EvaluatedPPath
  Assert.AreEqual llongExpectedNumberOfMatchingElements, VBA.Conversion.CLng(EvaluatedPPath.ReturnedValue), "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval077()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//StatusBar[count(./Button)=6]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test077, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
'Missing "." in nested ppath
Private Sub Test02Eval078()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//StatusBar[count(/Button)=6]"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
 Assert.AreEqual Phosphorus.PPathConstants.UNUSED_RELATIVE_PPATH_CONTEXT_NODE, EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
End Sub

'@TestMethod("PPath Evaluation")
'Invalid funciton name in nested ppath
Private Sub Test02Eval079()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//StatusBar[notafunction(/Button)=6]"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
 Assert.AreEqual Phosphorus.PPathConstants.INVALID_PREDICATE & " [notafunction(/Button)=6] => notafunction(/Button)=6", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval080()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[And(count(./Button)>=1,count(./SplitButton)>1,count(./ComboBox)>1)]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test080, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval081()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Tab[not(./Button)]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test081, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Test 3 step levels
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval082()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane/ComboBox/Edit"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test082, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval083()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane[count(./ComboBox)=1]/ComboBox/Edit[1]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test083, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval084()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem[./preceding-sibling::Button]"
  llongExpectedNumberOfMatchingElements = 16
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test084, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval085()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem[./following-sibling::Button]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test085, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval086()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem[Or(./preceding-sibling::Button, ./following-sibling::Button)]"
  llongExpectedNumberOfMatchingElements = 18
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test086, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval087()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ToolBar[@Name=" & VBA.Strings.Chr(34) & "Quick Access Toolbar" & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test087, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval088()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button[@AcceleratorKey=" & VBA.Strings.Chr(34) & "Ctrl+S" & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test088, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval089()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button[@AccessKey=" & VBA.Strings.Chr(34) & "Alt+Down Arrow" & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test089, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval090()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Text[@AriaProperties=" & VBA.Strings.Chr(34) & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test090, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval091()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@AriaRole=" & VBA.Strings.Chr(34) & "" & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test091, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval092()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[@AutomationId=" & VBA.Strings.Chr(34) & "AutoSaveSwitch" & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test092, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval093()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//TabItem[@BoundingRectangle=" & VBA.Strings.Chr(34) & "#" & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 13
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, , , True)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test093, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval094()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@CenterPoint=""""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test094, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval095()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[@ClassName=" & VBA.Strings.Chr(34) & "DropShadow" & VBA.Strings.Chr(34) & "]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test095, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval096()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@Culture=0]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test096, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval097()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@FillColor=""""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test097, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval098()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@FillType=0]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test098, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval099()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@FrameworkId=""Win32""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test099, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval100()
Arrange:
'  PPath.SetDebugMode = True
  'Use //* to test for not matching long descriptions
  strTestPPath = "//*[@FullDescription=""Pick a new font for your text.""]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test100, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval101()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Button[@FullDescription=""Like the look of a particular selection? You can apply that look to other content in the document.\n\nTo get started: \n1. Select content with the formatting you like\n2. Click Format Painter\n3. Select something else to automatically apply the formatting\n\nFYI: To apply the formatting in multiple places, double-click Format Painter.""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test101, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub


'@TestMethod("PPath Evaluation")
Private Sub Test02Eval102()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane[@HasKeyboardFocus=False]"
  llongExpectedNumberOfMatchingElements = 5
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test102, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval103()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ToolBar[@HeadingLevel=80050]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test103, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval104()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Pane/ToolBar[@HelpText=""Ribbon toolbar""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test104, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval105()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Tab[@IsContentElement=True]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test105, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval106()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/TitleBar//MenuItem[@IsControlElement=True]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test106, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval107()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/TitleBar//MenuItem[@IsDataValidForForm=False]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test107, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval108()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/TitleBar//MenuItem[@IsDialog=False]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test108, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval109()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@IsEnabled=False]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test109, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval110()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Text[@IsKeyboardFocusable=True]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test110, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval111()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[@IsOffscreen=True]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test111, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval112()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Text[@IsPassword=False]"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test112, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval113()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit[@IsPeripheral=False]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test113, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval114()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Text[@IsRequiredForForm=False]"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test114, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval115()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[@ItemStatus<>""""]"
  llongExpectedNumberOfMatchingElements = 0
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test115, EvaluatedPPath
  Assert.AreEqual False, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval116()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Edit[@ItemType=""Edit Formula""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test116, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval117()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit[@LandmarkType=0]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test117, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval118()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "/Edit[@Level=0]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test118, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval119()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@LiveSetting=2]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test119, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval120()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@LocalizedControlType=""ComboBox""]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test120, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval121()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@LocalizedLandmarkType=""""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test121, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval122()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[@Name=""Zoom 10%""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test122, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval123()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@NativeWindowHandle=""#""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, , , True)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test123, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval124()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@OptimizeForVisualContent=False]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test124, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval125()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@Orientation=1]"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test125, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval126()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[@PositionInSet=13]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test126, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval127()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@ProcessId=""#""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, , , True)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test127, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval128()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@ProviderDescription=""#""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath, , , True)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test128, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval129()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@Rotation=""""]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test129, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval130()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//*[@SizeOfSet=2]"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test130, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval131()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//ComboBox[@VisualEffects=0]"
  llongExpectedNumberOfMatchingElements = 4
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test131, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval132()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem[position()<=2][1]"
  llongExpectedNumberOfMatchingElements = 22
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test132, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation")
Private Sub Test02Eval133()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem[And(./preceding-sibling::Button,./following-sibling::SplitButton)]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test133, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Variation of #133
'@TestMethod("PPath Evaluation")
Private Sub Test02Eval134()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//MenuItem[./preceding-sibling::Button][./following-sibling::SplitButton]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PhosphorusTests.PPathTestsExpectedResults.Test134, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub
