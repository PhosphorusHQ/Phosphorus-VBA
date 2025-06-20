Attribute VB_Name = "PPathTestsExcel"
'@Folder PPathTests
'@TestModule

Option Explicit
Option Private Module
Option Base 1

'Run all PPath test on a blank new Excel Workbook
Private eleExcelRootElement As UIAutomationClient.IUIAutomationElement
Private wbTest As Excel.Workbook

Private strTestPPath As String
Private PPath As Phosphorus.PPath
Private EvaluatedPPath As Phosphorus.PPathReturnClass
Private llongExpectedNumberOfMatchingElements As Long

'@ModuleInitialize
Private Sub ModuleInitialize()
  PhosphorusFactory.GetLogger
  Logger.SetTempLevel LogLevel.INTERNAL_INFO
  Logger.InternalInfo "Logger Started"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
  Set Assert = Nothing
  Set Fakes = Nothing
  CloseTestWorkbook
  PPathWorkbook.CloseWB
  Utils.CloseAllOtherWorkbooks
  Logger.InternalInfo "Logger Stopped"
  Set Logger = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
  'Test117 loses the referece to the test wb so we need to close it and reopen it
  If wbTest Is Nothing Then
    CreateTestWorkbook
  End If
  'Test117 loses these references
  If Assert Is Nothing Then
    Set Assert = CreateObject("Rubberduck.AssertClass")
  End If
  If Fakes Is Nothing Then
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
  End If
  
  FindExcelRootElement
  Set PPath = Nothing
  Set PPath = Phosphorus.PhosphorusFactory.GetNewPhosphorusPPath
  PPath.Initialise
  PPath.SetApplicationRootElement eleExcelRootElement
End Sub

'@TestCleanup
Private Sub TestCleanup()

  If Not PPath Is Nothing Then
    If PPath.GetDebugMode Then
      OutputActualXPaths EvaluatedPPath.GetMatchingNavigationalPPaths
    End If
  End If
  Set EvaluatedPPath = Nothing
  Set PPath = Nothing

End Sub

Private Function CreateTestWorkbook()
  Set wbTest = Excel.Workbooks.Add
  
  PPathTestsExpectedResults.TestWorkbookName = wbTest.Name
  
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
      oAutomation.CreatePropertyCondition(UIAutomationClient.UIA_PropertyIds.UIA_NamePropertyId, PPathTestsExpectedResults.TestWorkbookName & " - Excel") _
    ) _
  )
  Set eleExcelRootElement = rootElement
End Function

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel001()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel001, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel002()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//DataItem[@Name=""A1""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel002, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel003()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)"
  llongExpectedNumberOfMatchingElements = 6
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel003, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel004()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, integer)"
  llongExpectedNumberOfMatchingElements = 5
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel004, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel005()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, decimal)"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel005, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel006()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, hyperlink)"
  llongExpectedNumberOfMatchingElements = 2
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel006, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel007()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, date)"
  llongExpectedNumberOfMatchingElements = 3
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel007, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel008()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, boolean)"
  llongExpectedNumberOfMatchingElements = 6
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel008, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel009()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "count(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Assert.AreEqual llongExpectedNumberOfMatchingElements, VBA.Conversion.CLng(EvaluatedPPath.ReturnedValue), "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel010()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "sum(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Assert.AreEqual 1423.77, VBA.Conversion.CDbl(EvaluatedPPath.ReturnedValue), "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel011()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "average(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Assert.AreEqual 142.377, VBA.Conversion.CDbl(EvaluatedPPath.ReturnedValue), "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel012()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "stdeva(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Assert.AreEqual "204.529890075863", VBA.Conversion.CStr(EvaluatedPPath.ReturnedValue), "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel013()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "round(stdeva((//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))),2)"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Assert.AreEqual 204.53, VBA.Conversion.CDbl(EvaluatedPPath.ReturnedValue), "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel014()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "round(stdeva((//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))),6)"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Assert.AreEqual 204.52989, VBA.Conversion.CDbl(EvaluatedPPath.ReturnedValue), "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel015()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "min(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Assert.AreEqual -1, VBA.Conversion.CInt(EvaluatedPPath.ReturnedValue), "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel016()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "max(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Assert.AreEqual 521.01, VBA.Conversion.CDbl(EvaluatedPPath.ReturnedValue), "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel017()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "roundup(stdeva((//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))),0)"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Assert.AreEqual 205, VBA.Conversion.CInt(EvaluatedPPath.ReturnedValue), "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel018()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "rounddown(stdeva((//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))),4)"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Assert.AreEqual 204.5298, VBA.Conversion.CDbl(EvaluatedPPath.ReturnedValue), "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test05Excel019()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "product(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal))"
  llongExpectedNumberOfMatchingElements = 10
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel009, EvaluatedPPath
  Assert.AreEqual "-4252982755186.95", VBA.Conversion.CStr(EvaluatedPPath.ReturnedValue), "ReturnedValue"
End Sub

'Use text() function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel100()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//DataItem[text()=""Cat""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel100, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'PPath in Excel can levereage any Excel forumulas in predicates
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel101()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//DataItem[lower(text())=""cat""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel100, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel102()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//DataItem[upper(text())=""CAT""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel100, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel103()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Pane[@Name=""Sheet Sheet1""]//DataItem[@Name=""A2""])[upper(substitute(text(),""at"",""ure""))=""CURE""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel100, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'REM This is slow and fails randomly - us VBA Web Service Call instead? @TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel104()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Pane[@Name=""Sheet Sheet1""]//DataItem[@Name=""A2""])[translate(text(),""en"",""fr"")=""Chat""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel100, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'User defined function local to Phosphorus
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel105()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "(//Pane[@Name=""Sheet Sheet1""]//DataItem[@Name=""A2""])[udf:reversetext(text())=""taC""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel100, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel106()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "sum(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal)[ceiling_math(value())=405])"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel106, EvaluatedPPath
  Assert.AreEqual 404.36, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Excel formula with parameters
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel107()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "sum(//Pane[@Name=""Sheet Sheet1""]//element(*, integer)[ceiling_math(value(),6,TRUE)=-6])"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel107, EvaluatedPPath
  Assert.AreEqual -1, VBA.Conversion.CInt(EvaluatedPPath.ReturnedValue), "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel108()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "sum(//Pane[@Name=""Sheet Sheet1""]//element(*, decimal)[floor_math(value(),5,false)=-5])"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel108, EvaluatedPPath
  Assert.AreEqual -1, VBA.Conversion.CInt(EvaluatedPPath.ReturnedValue), "ReturnedValue"
End Sub

'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel109()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[concat(text(),text(),text())=""RabbitRabbitRabbit""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel109, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'User defined function to replicate an xpath function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel110()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, hyperlink)[xp:contains(text(),""google"")]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel110, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'User defined function to replicate an xpath function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel111()
Arrange:
'  PPath.SetDebugMode = True  '[@Name=""B6""]
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, date)[xp:format-number(text(),""ddmmyyyy"")=""06042003""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel111, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'User defined function to replicate an xpath function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel112()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:starts-with(text(),""1st April"")]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel112, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'User defined function to replicate an xpath function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel113()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:normalize-space(text())=""OneTwoThreeFourFive Six""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel113, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'User defined function to replicate an xpath function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel114()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:string-length(text())=6]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'User defined function to replicate an xpath function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel115()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:substring(text(),4)=""bit""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'User defined function to replicate an xpath function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel116()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:substring(text(),2,5)=""abbi""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'References to workbooks of End-User user defined functions can be added through the framework
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel117()
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
    AddReferenceToWorkbook wbUDFs
  End If
    
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[udf:alldogsarecats(text())=""Cat""]"
  llongExpectedNumberOfMatchingElements = 2
    
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)

Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel117, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
    
'Tidy up
  RemoveAllAddedReferences
  wbUDFs.Close savechanges:=False
  Set wbUDFs = Nothing
  
  'This test loses a reference to the testWB so we need to close it while we hav a reference to it and then for it to reopen
  CloseTestWorkbook

End Sub

'User defined function to replicate an xpath function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel118()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:string-after(text(),""abb"")=""it""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'User defined function to replicate an xpath function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel119()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:string-before(text(),""bit"")=""Rab""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'User defined function to replicate an xpath function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel120()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[xp:translate(text(),""bit"",""BIT"")=""RabBIT""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'User defined function to replicate an xpath function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel121()
Arrange:
'  PPath.SetDebugMode = True
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[udf:text-between(text(),""Ra"",""it"")=""bb""]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Test for mapping a renamed a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test06Excel122()
Arrange:
'  PPath.SetDebugMode = True
  PPathExceUserDefinedFunctions.ClearDownFunctionNameMappings
  PPathExceUserDefinedFunctions.InitialiseFunctionNameMappings
  PPathExceUserDefinedFunctions.AddFunctionNameMapping "udftest:starts-with", "udf_test_starts_with"
  PPathExceUserDefinedFunctions.AddFunctionNameMapping "udftest:ends-with", "udf_test_ends_with"
  strTestPPath = "//Pane[@Name=""Sheet Sheet1""]//element(*, string)[and(udftest:starts-with(text(),""Rab""),udftest:ends-with(text(),""bit""))]"
  llongExpectedNumberOfMatchingElements = 1
Act:
  Set EvaluatedPPath = PPath.Evaluate(strTestPPath)
Assert:
  Assert.AreEqual "", EvaluatedPPath.GetErrorMessage, "GetErrorMessage"
  Assert.AreEqual llongExpectedNumberOfMatchingElements, EvaluatedPPath.GetFinalNumberOfMatchingElements, "GetMatchingElements"
  TestExpectedAndActualAllElementsPPath PPathTestsExpectedResults.TestExcel114, EvaluatedPPath
  Assert.AreEqual True, EvaluatedPPath.ReturnedValue, "ReturnedValue"
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel201()
Arrange:
  Dim strReturn As String
Act:
  strReturn = udf_reversetext("abc")
Assert:
  Assert.AreEqual "cba", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel202()
Arrange:
  Dim boolReturn As Boolean
Act:
  boolReturn = xp_contains("Rabbit", "abb")
Assert:
  Assert.AreEqual True, boolReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel203()
Arrange:
  Dim boolReturn As Boolean
Act:
  boolReturn = xp_contains("Rabbit", "cat")
Assert:
  Assert.AreEqual False, boolReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel204()
Arrange:
  Dim boolReturn As Boolean
Act:
  boolReturn = xp_starts_with("Rabbit", "Rabbi")
Assert:
  Assert.AreEqual True, boolReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel205()
Arrange:
  Dim boolReturn As Boolean
Act:
  boolReturn = xp_starts_with("Rabbit", "Rob")
Assert:
  Assert.AreEqual False, boolReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel206()
Arrange:
  Dim strReturn As String
Act:
  strReturn = xp_format_number(45735, "dd/mm/yyyy")
Assert:
  Assert.AreEqual "19/03/2025", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel207()
Arrange:
  Dim strReturn As String
Act:
  strReturn = xp_normalize_space(" " & "One" & vbCrLf & "Two" & vbCr & "Three" & vbLf & "Four" & vbTab & "Five" & "        " & "Six" & " ")
Assert:
  Assert.AreEqual "OneTwoThreeFourFive Six", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel208()
Arrange:
  Dim intReturn As Integer
Act:
  intReturn = xp_string_length(" " & "One" & vbCrLf & "Two" & vbCr & "Three" & vbLf & "Four" & vbTab & "Five" & "        " & "Six" & " ")
Assert:
  Assert.AreEqual 37, intReturn
End Sub


'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel209()
Arrange:
  Dim strReturn As String
Act:
  strReturn = xp_substring("Rabbit", 2, 4)
Assert:
  Assert.AreEqual "abb", strReturn
End Sub
  
'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel210()
Arrange:
  Dim strReturn As String
Act:
  strReturn = xp_substring("Rabbit", 4)
Assert:
  Assert.AreEqual "bit", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel211()
Arrange:
  Dim strReturn As String
Act:
  strReturn = xp_string_after("Extraneous", "")
Assert:
  Assert.AreEqual "Extraneous", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel212()
Arrange:
  Dim strReturn As String
Act:
  strReturn = xp_string_after("Extraneous", "ran")
Assert:
  Assert.AreEqual "eous", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel213()
Arrange:
  Dim strReturn As String
Act:
  strReturn = xp_string_after("Extraneous", "run")
Assert:
  Assert.AreEqual "Extraneous", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel214()
Arrange:
  Dim strReturn As String
Act:
  strReturn = xp_string_before("Extraneous", "ran")
Assert:
  Assert.AreEqual "Ext", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel215()
Arrange:
  Dim strReturn As String
Act:
  strReturn = xp_string_before("Extraneous", "run")
Assert:
  Assert.AreEqual "Extraneous", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel216()
Arrange:
  Dim strReturn As String
Act:
  strReturn = xp_translate("bar", "abc", "ABC")
Assert:
  Assert.AreEqual "BAr", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel217()
Arrange:
  Dim strReturn As String
Act:
  strReturn = xp_translate("--aaa--", "abc-", "ABC")
Assert:
  Assert.AreEqual "AAA", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel218()
Arrange:
  Dim strReturn As String
Act:
  strReturn = xp_translate("abcdabc", "abc", "AB")
Assert:
  Assert.AreEqual "ABdAB", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel219()
Arrange:
  Dim strReturn As String
Act:
  strReturn = udf_text_between("USER: myusername ADDRESS: unknown", "USER:", "ADDRESS:")
Assert:
  Assert.AreEqual " myusername ", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel220()
Arrange:
  Dim strReturn As String
Act:
  strReturn = udf_text_between("abc-123-xyz-000", "-", "-")
Assert:
  Assert.AreEqual "123", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel221()
Arrange:
  Dim strReturn As String
Act:
  strReturn = udf_text_between("abc-123-xyz-000", "-", "-", True)
Assert:
  Assert.AreEqual "xyz", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel222()
Arrange:
  Dim strReturn As String
Act:
  strReturn = udf_text_between("How now brown cow", "now", "anyrandomstring")
Assert:
  Assert.AreEqual " brown cow", strReturn
End Sub

'Unit test for a user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel223()
Arrange:
  Dim strReturn As String
Act:
  strReturn = udf_text_between("C:\Users\Ryan\Documents\readme.txt", "\", ".txt", True)
Assert:
  Assert.AreEqual "readme", strReturn
End Sub

'Unit test for a test user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel224()
Arrange:
  Dim boolReturn As Boolean
Act:
  boolReturn = udf_test_starts_with("Rabbit", "Rab")
Assert:
  Assert.AreEqual True, boolReturn
End Sub

'Unit test for a test user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test04Excel225()
Arrange:
  Dim boolReturn As Boolean
Act:
  boolReturn = udf_test_ends_with("Rabbit", "bit")
Assert:
  Assert.AreEqual True, boolReturn
End Sub

'Unit test for a user defined function renaming
'@TestMethod("PPath Evaluation Excel")
Private Sub Test03Excel300()
Arrange:
  Dim strReturn As String
Act:
  strReturn = RenameExcelFunctions("ceiling_math()")
Assert:
  Assert.AreEqual "ceiling.math()", strReturn
End Sub

'Unit test for a user defined function renaming
'@TestMethod("PPath Evaluation Excel")
Private Sub Test03Excel301()
Arrange:
  Dim strReturn As String
Act:
  strReturn = RenameExcelFunctions("floor_math()")
Assert:
  Assert.AreEqual "floor.math()", strReturn
End Sub

'Unit test for a user defined function renaming
'@TestMethod("PPath Evaluation Excel")
Private Sub Test03Excel302()
Arrange:
  Dim strReturn As String
Act:
  strReturn = RenameExcelFunctions("xp:format-number()")
Assert:
  Assert.AreEqual "xp_format_number()", strReturn
End Sub

'Unit test for a user defined function renaming
'@TestMethod("PPath Evaluation Excel")
Private Sub Test03Excel303()
Arrange:
  Dim strReturn As String
Act:
  strReturn = RenameExcelFunctions("xp:starts-with()")
Assert:
  Assert.AreEqual "xp_starts_with()", strReturn
End Sub

'Unit test for a user defined function renaming
'@TestMethod("PPath Evaluation Excel")
Private Sub Test03Excel304()
Arrange:
  Dim strReturn As String
Act:
  strReturn = RenameExcelFunctions("xp:normalize-space()")
Assert:
  Assert.AreEqual "xp_normalize_space()", strReturn
End Sub

'Unit test for a user defined function renaming
'@TestMethod("PPath Evaluation Excel")
Private Sub Test03Excel305()
Arrange:
  Dim strReturn As String
Act:
  strReturn = RenameExcelFunctions("xp:substring()")
Assert:
  Assert.AreEqual "xp_substring()", strReturn
End Sub

'Unit test for a user defined function renaming
'@TestMethod("PPath Evaluation Excel")
Private Sub Test03Excel306()
Arrange:
  Dim strReturn As String
Act:
  strReturn = RenameExcelFunctions("xp:string-length()")
Assert:
  Assert.AreEqual "xp_string_length()", strReturn
End Sub

'Unit test for a user defined function renaming
'@TestMethod("PPath Evaluation Excel")
Private Sub Test03Excel307()
Arrange:
  Dim strReturn As String
  PPathExceUserDefinedFunctions.ClearDownFunctionNameMappings
Act:
  strReturn = RenameExcelFunctions("xp:string-after()")
Assert:
  Assert.AreEqual "xp_string_after()", strReturn
End Sub

'Unit test for a user defined function renaming
'@TestMethod("PPath Evaluation Excel")
Private Sub Test03Excel308()
Arrange:
  Dim strReturn As String
Act:
  strReturn = PPathExceUserDefinedFunctions.RenameExcelFunctions("xp:string-before()")
Assert:
  Assert.AreEqual "xp_string_before()", strReturn
End Sub

'Unit test for a user defined function renaming
'@TestMethod("PPath Evaluation Excel")
Private Sub Test03Excel309()
Arrange:
  Dim strReturn As String
Act:
  strReturn = PPathExceUserDefinedFunctions.RenameExcelFunctions("xp:translate()")
Assert:
  Assert.AreEqual "xp_translate()", strReturn
End Sub

'Unit test for a user defined function renaming
'@TestMethod("PPath Evaluation Excel")
Private Sub Test03Excel310()
Arrange:
  Dim strReturn As String
Act:
  strReturn = PPathExceUserDefinedFunctions.RenameExcelFunctions("udf:text-between()")
Assert:
  Assert.AreEqual "udf_text_between()", strReturn
End Sub

'Unit test for a temporarily renamed user defined function
'@TestMethod("PPath Evaluation Excel")
Private Sub Test03Excel311()
Arrange:
  
  PPathExceUserDefinedFunctions.ClearDownFunctionNameMappings
  PPathExceUserDefinedFunctions.InitialiseFunctionNameMappings
  
  Dim strReturn1 As String
  Dim strReturn2 As String
  strReturn1 = PPathExceUserDefinedFunctions.RenameExcelFunctions("udftest:starts-with()")
  strReturn2 = PPathExceUserDefinedFunctions.RenameExcelFunctions("udftest:ends-with()")
  
  Dim strReturn3 As String
  Dim strReturn4 As String
  PPathExceUserDefinedFunctions.AddFunctionNameMapping "udftest:starts-with", "udf_test_starts_with"
  PPathExceUserDefinedFunctions.AddFunctionNameMapping "udftest:ends-with", "udf_test_ends_with"
Act:
  strReturn3 = PPathExceUserDefinedFunctions.RenameExcelFunctions("udftest:starts-with()")
  strReturn4 = PPathExceUserDefinedFunctions.RenameExcelFunctions("udftest:ends-with()")
Assert:
  Assert.AreEqual "udftest:starts-with()", strReturn1, "strReturn1"
  Assert.AreEqual "udftest:ends-with()", strReturn2, "strReturn2"
  Assert.AreEqual "udf_test_starts_with()", strReturn3, "strReturn3"
  Assert.AreEqual "udf_test_ends_with()", strReturn4, "strReturn4"
End Sub

