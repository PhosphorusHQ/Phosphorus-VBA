Attribute VB_Name = "pPathTestsExcel"
'@Folder pPath
'@TestModule
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit
Option Private Module
Option Base 1

'Run all PPath test on a blank new Excel Workbook
Private eleExcelRootElement As UIAutomationClient.IUIAutomationElement
Private wbTest As Excel.Workbook

Private strTestPPath As String
Private testpPath As pPath.Core
Private EvaluatedPPath As pPath.ReturnClass
Private llongExpectedNumberOfMatchingElements As Long

'Private Sub BeforeModule()
'End Sub

Private Sub AfterModule()
  On Error GoTo ErrorHandler
 
'  Set pUnitTests.pPathTestsCommon.Assert = Nothing
'  Set pUnitTests.pPathTestsCommon.Fakes = Nothing
  CloseTestWorkbook
  pPath.Workbook.CloseWB
  Phosphorus.Utils.CloseAllOtherWorkbooks

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Private Sub BeforeTest()
  On Error GoTo ErrorHandler
 
  VBA.Interaction.DoEvents
 
  'Test117 loses the reference to the test wb so we need to close it and reopen it
  If wbTest Is Nothing Then
    CreateTestWorkbook
  End If
'  'Test117 loses these references
'  If pPathTestsCommon.Assert Is Nothing Then
'    Set pUnitTests.pPathTestsCommon.Assert = CreateObject("Rubberduck.AssertClass")
'  End If
'  If pUnitTests.pPathTestsCommon.Fakes Is Nothing Then
'    Set pUnitTests.pPathTestsCommon.Fakes = CreateObject("Rubberduck.FakesProvider")
'  End If
  
  FindExcelRootElement
  Set testpPath = Nothing
  Set testpPath = pPath.ConstantsAndStatic.GetNewPhosphorusPPath
  testpPath.Initialise
  testpPath.SetApplicationRootElement eleExcelRootElement

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

Private Sub AfterTest()
  On Error GoTo ErrorHandler
  
  If Not testpPath Is Nothing Then
'    If testpPath.GetDebugMode Then
'      pUnitTests_pPath.pPathTestsCommon.OutputActualXPaths EvaluatedPPath.GetMatchingNavigationalPPaths
'    End If
  End If
  Set EvaluatedPPath = Nothing
  Set testpPath = Nothing

  VBA.Interaction.DoEvents
  
  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
  VBA.Interaction.DoEvents
End Sub

Private Function CreateTestWorkbook()

  Set wbTest = Excel.Workbooks.Add
  
  pUnitTests_pPath.pPathTestsExpectedResults.TestWorkbookName = wbTest.Name
  
  wbTest.Sheets("Sheet1").Range("A1") = "Test Data"
  
  Dim myVariantArray As Variant
  myVariantArray = Array("Cat", "Dog", "", "Rabbit", " " & "One" & vbCrLf & "Two" & vbCr & "Three" & vbLf & "Four" & vbTab & "Five" & "        " & "Six" & " ")
  wbTest.Sheets("Sheet1").Range("A2:E2") = myVariantArray
  
  myVariantArray = Array(1, 2, 3, 4, -1)
  wbTest.Sheets("Sheet1").Range("A3:E3") = myVariantArray
  
  myVariantArray = Array(101.45, 22.7, 365.25, 404.36, 521.01)
  wbTest.Sheets("Sheet1").Range("A4:E4") = myVariantArray
  
  myVariantArray = Array("https://www.google.co.uk/", "https://www.bbc.co.uk/")
  wbTest.Sheets("Sheet1").Range("A5:B5") = myVariantArray
  
  myVariantArray = Array("31/5/2021", "1st April 1992", "19/5/2025 14:43", "April 6, 2003")
  wbTest.Sheets("Sheet1").Range("A6:D6") = myVariantArray
  
  myVariantArray = Array(True, False, "=1=1", "=1=2", "'True", "'False")
  wbTest.Sheets("Sheet1").Range("A7:F7") = myVariantArray
  
  VBA.Interaction.DoEvents
  
End Function

Private Function CloseTestWorkbook()
  If Not wbTest Is Nothing Then
    wbTest.Close savechanges:=False
    Set wbTest = Nothing
  End If
End Function

Private Function FindExcelRootElement()
  Dim oAutomation As CUIAutomation
  Set oAutomation = New CUIAutomation
  Dim rootElement As UIAutomationClient.IUIAutomationElement
  Set rootElement = oAutomation.GetRootElement
  Set rootElement = rootElement.FindFirst( _
    UIAutomationClient.TreeScope.TreeScope_Children, _
    oAutomation.CreateAndCondition( _
      oAutomation.CreatePropertyCondition(UIAutomationClient.UIA_PropertyIds.UIA_ControlTypePropertyId, UIAutomationClient.UIA_ControlTypeIds.UIA_WindowControlTypeId), _
      oAutomation.CreatePropertyCondition(UIAutomationClient.UIA_PropertyIds.UIA_NamePropertyId, pUnitTests_pPath.pPathTestsExpectedResults.TestWorkbookName & " - Excel") _
    ) _
  )
  Set eleExcelRootElement = rootElement
End Function

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel001()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel001, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel002()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//DataItem[@Name=""A1""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel002, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel003()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)"
  llongExpectedNumberOfMatchingElements = 6
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel003, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel004()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, integer)"
  llongExpectedNumberOfMatchingElements = 5
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel004, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel005()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, decimal)"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel005, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel006()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, hyperlink)"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel006, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel007()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, date)"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel007, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel008()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, boolean)"
  llongExpectedNumberOfMatchingElements = 6
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel008, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel009()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "count(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, VBA.Conversion.CLng(EvaluatedPPath.ReturnedValue), "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel010()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "sum(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal 1423.77, VBA.Conversion.CDbl(EvaluatedPPath.ReturnedValue), "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel011()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "average(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal 142.377, VBA.Conversion.CDbl(EvaluatedPPath.ReturnedValue), "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel012()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "stdeva(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "204.529890075863", VBA.Conversion.CStr(EvaluatedPPath.ReturnedValue), "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel013()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "round(stdeva((//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))),2)"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal 204.53, VBA.Conversion.CDbl(EvaluatedPPath.ReturnedValue), "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel014()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "round(stdeva((//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))),6)"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal 204.52989, VBA.Conversion.CDbl(EvaluatedPPath.ReturnedValue), "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel015()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "min(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal -1, VBA.Conversion.CInt(EvaluatedPPath.ReturnedValue), "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel016()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "max(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal 521.01, VBA.Conversion.CDbl(EvaluatedPPath.ReturnedValue), "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel017()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "roundup(stdeva((//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))),0)"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal 205, VBA.Conversion.CInt(EvaluatedPPath.ReturnedValue), "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel018()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "rounddown(stdeva((//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))),4)"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal 204.5298, VBA.Conversion.CDbl(EvaluatedPPath.ReturnedValue), "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel019()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "product(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal "-4252982755186.95", VBA.Conversion.CStr(EvaluatedPPath.ReturnedValue), "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Name with xp:starts-with
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test05Excel020()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[xp:starts-with(@Name,""Sheet Sheet1"")]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel001, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Use text() function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel100()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//DataItem[text()=""Cat""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel100, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'PPath in Excel can levereage any Excel forumulas in predicates
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel101()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//DataItem[lower(text())=""cat""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel100, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel102()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//DataItem[upper(text())=""CAT""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel100, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel103()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Pane[@Name=""Sheet Sheet1""]//DataItem[@Name=""A2""])[upper(substitute(text(),""at"",""ure""))=""CURE""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel100, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'REM This is slow and fails randomly - use VBA Web Service Call instead?
Private Sub Test06Excel104()
  On Error GoTo ErrorHandler
 
  Logger.Info "Running PPath Evaluation Excel - Test06Excel104"
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Pane[@Name=""Sheet Sheet1""]//DataItem[@Name=""A2""])[translate(text(),""en"",""fr"")=""Chat""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel100, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'User defined function local to Phosphorus
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel105()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Pane[@Name=""Sheet Sheet1""]//DataItem[@Name=""A2""])[udf:reversetext(text())=""taC""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel100, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel106()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "sum(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal)[ceiling_math(value())=405])"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel106, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal 404.36, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Excel formula with parameters
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel107()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "sum(//Pane[@Name=""Sheet Sheet1""]//element(*, integer)[ceiling_math(value(),6,TRUE)=-6])"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel107, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal -1, VBA.Conversion.CInt(EvaluatedPPath.ReturnedValue), "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel108()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "sum(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal)[floor_math(value(),5,false)=-5])"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel108, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal -1, VBA.Conversion.CInt(EvaluatedPPath.ReturnedValue), "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel109()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[concat(text(),text(),text())=""RabbitRabbitRabbit""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel109, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'User defined function to replicate an xpath function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel110()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, hyperlink)[xp:contains(text(),""google"")]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel110, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'User defined function to replicate an xpath function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel111()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True  '[@Name=""B6""]
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, date)[xp:format-number(text(),""ddmmyyyy"")=""06042003""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel111, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'User defined function to replicate an xpath function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel112()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:starts-with(text(),""1st April"")]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel112, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'User defined function to replicate an xpath function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel113()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:normalize-space(text())=""OneTwoThreeFourFive Six""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel113, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'User defined function to replicate an xpath function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel114()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:string-length(text())=6]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'User defined function to replicate an xpath function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel115()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:substring(text(),4)=""bit""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'User defined function to replicate an xpath function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel116()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:substring(text(),2,5)=""abbi""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'References to workbooks of End-User user defined functions can be added through the framework
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel117()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  
  Dim strPath As String
  strPath = VBA.Strings.Left(ThisWorkbook.FullName, VBA.InStr(1, ThisWorkbook.FullName, ThisWorkbook.Name) - 1) & "tests\UserDefinedFunctionsTest.xlam"
  Dim wbLoop As Workbook
  For Each wbLoop In Application.Workbooks
    If wbLoop.FullName = strPath Then
      wbLoop.Close
      Exit For
    End If
  Next
  
  Dim wbUDFs As Excel.Workbook
  If wbUDFs Is Nothing Then
    Set wbUDFs = Workbooks.Open(strPath)
  End If

  If Not wbUDFs Is Nothing Then
    pUnitTests_pPath.References.AddReferenceToWorkbookOrLibrary wbUDFs.FullName
  End If
    
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[udf:alldogsarecats(text())=""Cat""]"
  llongExpectedNumberOfMatchingElements = 2
    
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)

Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel117, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True
    
'Tidy up
'  RemoveAllAddedReferences
  wbUDFs.Close savechanges:=False
  Set wbUDFs = Nothing
  
  'This test loses a reference to the testWB so we need to close it while we have a reference to it and then force it to reopen
'PJG  CloseTestWorkbook

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'User defined function to replicate an xpath function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel118()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:string-after(text(),""abb"")=""it""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'User defined function to replicate an xpath function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel119()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:string-before(text(),""bit"")=""Rab""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'User defined function to replicate an xpath function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel120()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:translate(text(),""bit"",""BIT"")=""RabBIT""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'User defined function to replicate an xpath function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel121()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[udf:text-between(text(),""Ra"",""it"")=""bb""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Test for mapping a renamed a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test06Excel122()
  On Error GoTo ErrorHandler
 
Arrange:
'  PPath.SetDebugMode = True
  pPath.ExcelUserDefinedFunctions.ClearDownFunctionNameMappings
  pPath.ExcelUserDefinedFunctions.InitialiseFunctionNameMappings
  pPath.ExcelUserDefinedFunctions.AddFunctionNameMapping "udftest:starts-with", "udf_test_starts_with"
  pPath.ExcelUserDefinedFunctions.AddFunctionNameMapping "udftest:ends-with", "udf_test_ends_with"
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[and(udftest:starts-with(text(),""Rab""),udftest:ends-with(text(),""bit""))]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = testpPath.Evaluate(strTestPPath)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements", isCritical:=True
  pPathTestsCommon.TestExpectedAndActualAllElementsPPath pUnitTests_pPath.pPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Phosphorus.AssertionsStatic.pAssert.Equal True, EvaluatedPPath.ReturnedValue, "ReturnedValue", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel201()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.udf_reversetext("abc")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "cba", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel202()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim boolReturn As Boolean
Act:
  boolReturn = pPath.ExcelUserDefinedFunctions.xp_contains("Rabbit", "abb")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal True, boolReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel203()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim boolReturn As Boolean
Act:
  boolReturn = pPath.ExcelUserDefinedFunctions.xp_contains("Rabbit", "cat")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal False, boolReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel204()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim boolReturn As Boolean
Act:
  boolReturn = pPath.ExcelUserDefinedFunctions.xp_starts_with("Rabbit", "Rabbi")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal True, boolReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel205()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim boolReturn As Boolean
Act:
  boolReturn = pPath.ExcelUserDefinedFunctions.xp_starts_with("Rabbit", "Rob")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal False, boolReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel206()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.xp_format_number(45735, "dd/mm/yyyy")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "19/03/2025", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel207()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.xp_normalize_space(" " & "One" & vbCrLf & "Two" & vbCr & "Three" & vbLf & "Four" & vbTab & "Five" & "        " & "Six" & " ")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "OneTwoThreeFourFive Six", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel208()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim intReturn As Integer
Act:
  intReturn = pPath.ExcelUserDefinedFunctions.xp_string_length(" " & "One" & vbCrLf & "Two" & vbCr & "Three" & vbLf & "Four" & vbTab & "Five" & "        " & "Six" & " ")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal 37, intReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel209()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.xp_substring("Rabbit", 2, 4)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "abb", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub
  
'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel210()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.xp_substring("Rabbit", 4)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "bit", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel211()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.xp_string_after("Extraneous", "")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "Extraneous", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel212()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.xp_string_after("Extraneous", "ran")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "eous", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel213()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.xp_string_after("Extraneous", "run")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "Extraneous", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel214()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.xp_string_before("Extraneous", "ran")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "Ext", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel215()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.xp_string_before("Extraneous", "run")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "Extraneous", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel216()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.xp_translate("bar", "abc", "ABC")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "BAr", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel217()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.xp_translate("--aaa--", "abc-", "ABC")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "AAA", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel218()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.xp_translate("abcdabc", "abc", "AB")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "ABdAB", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel219()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.udf_text_between("USER: myusername ADDRESS: unknown", "USER:", "ADDRESS:")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal " myusername ", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel220()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.udf_text_between("abc-123-xyz-000", "-", "-")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "123", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel221()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.udf_text_between("abc-123-xyz-000", "-", "-", True)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "xyz", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel222()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.udf_text_between("How now brown cow", "now", "anyrandomstring")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal " brown cow", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel223()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.udf_text_between("C:\Users\Ryan\Documents\readme.txt", "\", ".txt", True)
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "readme", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a test user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel224()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim boolReturn As Boolean
Act:
  boolReturn = pPath.ExcelUserDefinedFunctions.udf_test_starts_with("Rabbit", "Rab")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal True, boolReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a test user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test04Excel225()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim boolReturn As Boolean
Act:
  boolReturn = pPath.ExcelUserDefinedFunctions.udf_test_ends_with("Rabbit", "bit")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal True, boolReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function renaming
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test03Excel300()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("ceiling_math()")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "ceiling.math()", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function renaming
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test03Excel301()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("floor_math()")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "floor.math()", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function renaming
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test03Excel302()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("xp:format-number()")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "xp_format_number()", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function renaming
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test03Excel303()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("xp:starts-with()")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "xp_starts_with()", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function renaming
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test03Excel304()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("xp:normalize-space()")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "xp_normalize_space()", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function renaming
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test03Excel305()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("xp:substring()")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "xp_substring()", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function renaming
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test03Excel306()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("xp:string-length()")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "xp_string_length()", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function renaming
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test03Excel307()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
  pPath.ExcelUserDefinedFunctions.ClearDownFunctionNameMappings
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("xp:string-after()")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "xp_string_after()", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function renaming
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test03Excel308()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("xp:string-before()")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "xp_string_before()", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function renaming
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test03Excel309()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("xp:translate()")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "xp_translate()", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a user defined function renaming
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test03Excel310()
  On Error GoTo ErrorHandler
 
Arrange:
  Dim strReturn As String
Act:
  strReturn = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("udf:text-between()")
Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "udf_text_between()", strReturn, isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub

'Unit test for a temporarily renamed user defined function
'@Tag(PPathEvaluationExcel)
'@TestMethod
Private Sub Test03Excel311()
  On Error GoTo ErrorHandler
 
Arrange:
  pPath.ExcelUserDefinedFunctions.ClearDownFunctionNameMappings
  pPath.ExcelUserDefinedFunctions.InitialiseFunctionNameMappings
  
  Dim strReturn1 As String
  Dim strReturn2 As String
  strReturn1 = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("udftest:starts-with()")
  strReturn2 = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("udftest:ends-with()")
  
  Dim strReturn3 As String
  Dim strReturn4 As String
  pPath.ExcelUserDefinedFunctions.AddFunctionNameMapping "udftest:starts-with", "udf_test_starts_with"
  pPath.ExcelUserDefinedFunctions.AddFunctionNameMapping "udftest:ends-with", "udf_test_ends_with"

Act:
  strReturn3 = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("udftest:starts-with()")
  strReturn4 = pPath.ExcelUserDefinedFunctions.RenameExcelFunctions("udftest:ends-with()")

Assert:
  Phosphorus.AssertionsStatic.pAssert.Equal "udftest:starts-with()", strReturn1, "strReturn1", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal "udftest:ends-with()", strReturn2, "strReturn2", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal "udf_test_starts_with()", strReturn3, "strReturn3", isCritical:=True
  Phosphorus.AssertionsStatic.pAssert.Equal "udf_test_ends_with()", strReturn4, "strReturn4", isCritical:=True

  Exit Sub
ErrorHandler:
  Phosphorus.pUnitErrorStatic.TrapError Err.Number, Err.Description
End Sub
