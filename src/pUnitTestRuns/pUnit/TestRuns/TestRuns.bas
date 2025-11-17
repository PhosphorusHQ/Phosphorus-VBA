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

Private LastCreatedPUnitLogFileContents As String
Private lboolRunningAllTests As Boolean
Private lstrTestName As String
Private lintCountOfAssertionPasses As Integer
Private lintCountOfAssertionFails As Integer
Private lintCountOfTestFails As Integer
Private lboolRunnningAllTests As Boolean
Private thisLogger As Object 'Phosphorus.Log4P

Public Sub RunAllTests()
  Phosphorus.Log4PStatic.GetLogger
  Set thisLogger = Phosphorus.Factory.GetNewLogger
  lboolRunningAllTests = True
  lintCountOfTestFails = 0
  lboolRunnningAllTests = True
  RunUnitTestingTests01
  RunUnitTestingTests02
  RunUnitTestingTests03
  RunUnitTestingTests04
  RunUnitTestingTests05
  RunUnitTestingTests06
  RunUnitTestingTests07
  RunUnitTestingTests08
  RunUnitTestingTests09
  RunUnitTestingTests10
  RunUnitTestingTests11
  RunUnitTestingTests12
  RunUnitTestingTests13

  lboolRunningAllTests = False
  If lintCountOfTestFails = 0 Then
    WriteToLog "No Tests FAILED!"
    Debug.Print "No Tests FAILED!"
  Else
    WriteToLog lintCountOfTestFails & " Tests FAILED!"
    MsgBox lintCountOfTestFails & " Tests FAILED!"
  End If
'  Set log = Nothing
  Phosphorus.Log4PStatic.CloseLogger
  Set thisLogger = Nothing
  lboolRunnningAllTests = False
End Sub

Private Sub RunUnitTestingTests01()
Arrange:
  thisLogger.ExternalInfo "RunUnitTestingTests1"
  StartTest "RunUnitTestingTests01"
Act:
  pUnit.TestRunner.RunAllTests "TestModule01", "", "pUnitTests_pUnit"
  GetLastCreatedPUnitLogFileContents
Assert:
  AssertLogFileContains "Target project name is: pUnitTests_pUnit"
  AssertLogFileContains "Starting unit test execution"
  AssertLogFileContains "Module filter: TestModule01"
  AssertLogFileContains "Annotation filter: All tests"
  AssertLogFileDoesntContain "Test skipped due to annotations at line"
  AssertLogFileContains "Processing module: pUnitTests_pUnit.TestModule01"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.BeforeModule"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.BeforeTest"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule01.TestAddition"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule01.TestDivision"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule01.TestFailure"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.AfterTest"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.AfterModule"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule01: Setup Duration:"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule01 completed"
  AssertLogFileContains "Detailed Test Results:"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule01.TestAddition, Status: Passed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule01.TestDivision, Status: Passed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule01.TestFailure, Status: Passed"
  AssertLogFileContains "Total Test Method Duration (excluding setup/teardown):"
  AssertLogFileContains "Total Setup Duration:"
  AssertLogFileContains "Total Teardown Duration:"
  AssertLogFileContains "Total Run Duration:"
  AssertLogFileContains "Test execution completed."
  AssertLogFileContains "Total Tests: 3, Passed: 3, Failed: 0, Skipped: 0"
  LogSummary
End Sub

Private Sub RunUnitTestingTests02()
Arrange:
  StartTest "RunUnitTestingTests02"
Act:
  pUnit.TestRunner.RunAllTests "TestModule01", "Regression", "pUnitTests_pUnit"
  GetLastCreatedPUnitLogFileContents
Assert:
  AssertLogFileContains "Target project name is: pUnitTests_pUnit"
  AssertLogFileContains "Starting unit test execution"
  AssertLogFileContains "Module filter: TestModule01"
  AssertLogFileContains "Annotation filter: Regression"
  AssertLogFileContains "Test skipped due to annotations at line"
  AssertLogFileContains "Processing module: pUnitTests_pUnit.TestModule01"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.BeforeModule"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.BeforeTest"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.TestAddition"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.TestDivision"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule01.TestFailure"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.AfterTest"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.AfterModule"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule01 completed"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule01: Setup Duration:"
  AssertLogFileContains "Detailed Test Results:"
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule01.TestAddition, Status: Passed"
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule01.TestDivision, Status: Passed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule01.TestFailure, Status: Passed"
  AssertLogFileContains "Total Test Method Duration (excluding setup/teardown):"
  AssertLogFileContains "Total Setup Duration:"
  AssertLogFileContains "Total Teardown Duration:"
  AssertLogFileContains "Total Run Duration:"
  AssertLogFileContains "Test execution completed."
  AssertLogFileContains "Total Tests: 1, Passed: 1, Failed: 0, Skipped: 0"
  LogSummary
End Sub

Private Sub RunUnitTestingTests03()
Arrange:
  StartTest "RunUnitTestingTests03"
Act:
  pUnit.TestRunner.RunAllTests "TestModule01", "OR(Regression,SmokeTest)", "pUnitTests_pUnit"
  GetLastCreatedPUnitLogFileContents
Assert:
  AssertLogFileContains "Target project name is: pUnitTests_pUnit"
  AssertLogFileContains "Starting unit test execution"
  AssertLogFileContains "Module filter: TestModule01"
  AssertLogFileContains "Annotation filter: OR(Regression,SmokeTest)"
  AssertLogFileContains "Test skipped due to annotations at line"
  AssertLogFileContains "Processing module: pUnitTests_pUnit.TestModule01"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.BeforeModule"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.BeforeTest"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule01.TestAddition"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.TestDivision"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule01.TestFailure"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.AfterTest"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.AfterModule"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule01 completed"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule01: Setup Duration:"
  AssertLogFileContains "Detailed Test Results:"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule01.TestAddition, Status: Passed"
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule01.TestDivision, Status: Passed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule01.TestFailure, Status: Passed"
  AssertLogFileContains "Total Test Method Duration (excluding setup/teardown):"
  AssertLogFileContains "Total Setup Duration:"
  AssertLogFileContains "Total Teardown Duration:"
  AssertLogFileContains "Total Run Duration:"
  AssertLogFileContains "Test execution completed."
  AssertLogFileContains "Total Tests: 2, Passed: 2, Failed: 0, Skipped: 0"
  LogSummary
End Sub

Private Sub RunUnitTestingTests04()
Arrange:
  StartTest "RunUnitTestingTests04"
Act:
  pUnit.TestRunner.RunAllTests "TestModule01", "AND(OR(Regression,SmokeTest),NOT(Slow))", "pUnitTests_pUnit"
  GetLastCreatedPUnitLogFileContents
Assert:
  AssertLogFileContains "Target project name is: pUnitTests_pUnit"
  AssertLogFileContains "Starting unit test execution"
  AssertLogFileContains "Module filter: TestModule01"
  AssertLogFileContains "Annotation filter: AND(OR(Regression,SmokeTest),NOT(Slow))"
  AssertLogFileContains "Test skipped due to annotations at line"
  AssertLogFileContains "Processing module: pUnitTests_pUnit.TestModule01"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.BeforeModule"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.BeforeTest"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule01.TestAddition"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.TestDivision"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.TestFailure"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.AfterTest"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.AfterModule"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule01: Setup Duration:"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule01 completed"
  AssertLogFileContains "Detailed Test Results:"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule01.TestAddition, Status: Passed"
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule01.TestDivision, Status: Passed"
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule01.TestFailure, Status: Passed"
  AssertLogFileContains "Total Test Method Duration (excluding setup/teardown):"
  AssertLogFileContains "Total Setup Duration:"
  AssertLogFileContains "Total Teardown Duration:"
  AssertLogFileContains "Total Run Duration:"
  AssertLogFileContains "Test execution completed."
  AssertLogFileContains "Total Tests: 1, Passed: 1, Failed: 0, Skipped: 0"
  LogSummary
End Sub

Private Sub RunUnitTestingTests05()
Arrange:
  StartTest "RunUnitTestingTests05"
Act:
  pUnit.TestRunner.RunAllTests "TestModule02", "", "pUnitTests_pUnit"
  GetLastCreatedPUnitLogFileContents
Assert:
  AssertLogFileContains "Target project name is: pUnitTests"
  AssertLogFileContains "Starting unit test execution"
  AssertLogFileContains "Module filter: TestModule02"
  AssertLogFileContains "Annotation filter: All tests"
  AssertLogFileDoesntContain "Test skipped due to annotations at line"
  AssertLogFileContains "Processing module: pUnitTests_pUnit.TestModule02"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule02.BeforeModule"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule02.BeforeTest"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule02.TestAddition"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule02.TestDivision"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule02.TestFailure"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule02.AfterTest"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule02.AfterModule"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule02: Setup Duration:"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule02 completed"
  AssertLogFileContains "Detailed Test Results:"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule02.TestAddition, Status: Passed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule02.TestDivision, Status: Passed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule02.TestFailure, Status: Passed"
  AssertLogFileContains "Total Test Method Duration (excluding setup/teardown):"
  AssertLogFileContains "Total Setup Duration:"
  AssertLogFileContains "Total Teardown Duration:"
  AssertLogFileContains "Total Run Duration:"
  AssertLogFileContains "Test execution completed."
  AssertLogFileContains "Total Tests: 3, Passed: 3, Failed: 0, Skipped: 0"
  LogSummary
End Sub

Private Sub RunUnitTestingTests06()
Arrange:
  StartTest "RunUnitTestingTests06"
Act:
  pUnit.TestRunner.RunAllTests "OR(TestModule01, TestModule02, TestModule03)", "Slow", "pUnitTests_pUnit"
  GetLastCreatedPUnitLogFileContents
Assert:
  AssertLogFileContains "Target project name is: pUnitTests_pUnit"
  AssertLogFileContains "Starting unit test execution"
  AssertLogFileContains "Module filter: OR(TestModule01, TestModule02, TestModule03)"
  AssertLogFileContains "Annotation filter: Slow"
  
  AssertLogFileContains "Test skipped due to annotations at line 14: pUnitTests_pUnit.TestModule01.TestAddition (Annotations: '@SmokeTest or Procedure Name: TestAddition)"
  AssertLogFileContains "Processing module: pUnitTests_pUnit.TestModule01"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule01.TestAddition"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule01.TestDivision"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule01.TestFailure"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule01: Setup Duration:"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule01 completed"

  AssertLogFileContains "Test skipped due to annotations at line 30: pUnitTests_pUnit.TestModule02.TestAddition (Annotations: '@Regression or Procedure Name: TestAddition)"
  AssertLogFileContains "Test skipped due to annotations at line 42: pUnitTests_pUnit.TestModule02.TestDivision (Annotations: '@SmokeTest or Procedure Name: TestDivision)"
  AssertLogFileContains "Processing module: pUnitTests_pUnit.TestModule02"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule02.BeforeModule"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule02.BeforeTest(TestFailure)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule02.TestFailure"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule02.AfterTest(TestFailure)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule02.AfterModule"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule02: Setup Duration:"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule02 completed"
  
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule03 has no tests to run after filtering."
  AssertLogFileDoesntContain "Processing module: pUnitTests_pUnit.TestModule03"
    
  AssertLogFileContains "Detailed Test Results:"
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule01.TestAddition, Status: Passed"
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule01.TestAddition, Status: Failed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule01.TestDivision, Status: Passed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule01.TestFailure, Status: Passed"
  
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule02.TestAddition, Status: Passed"
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule02.TestAddition, Status: Failed"
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule02.TestDivision, Status: Passed"
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule02.TestDivision, Status: Failed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule02.TestFailure, Status: Passed"
  
  AssertLogFileContains "Total Test Method Duration (excluding setup/teardown):"
  AssertLogFileContains "Total Setup Duration:"
  AssertLogFileContains "Total Teardown Duration:"
  AssertLogFileContains "Total Run Duration:"
  AssertLogFileContains "Test execution completed."
  AssertLogFileContains "Total Tests: 3, Passed: 3, Failed: 0, Skipped: 0"
  
  LogSummary

End Sub

Private Sub RunUnitTestingTests07()
Arrange:
  StartTest "RunUnitTestingTests07"
Act:
  pUnit.TestRunner.RunAllTests "TestModule04", "", "pUnitTests_pUnit"
  GetLastCreatedPUnitLogFileContents
Assert:
  AssertLogFileContains "Target project name is: pUnitTests_pUnit"
  AssertLogFileContains "Starting unit test execution"
  AssertLogFileContains "Module filter: TestModule04"
  AssertLogFileContains "Annotation filter: All tests"
  
  AssertLogFileDoesntContain "Test skipped due to annotations at line"
  AssertLogFileContains "Processing module: pUnitTests_pUnit.TestModule04"
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule04.BeforeModule"
  AssertLogFileContains "BeforeModule failed in module pUnitTests_pUnit.TestModule04: Error #-2147220504: Intentional failure in BeforeModule"
  AssertLogFileContains "Skipping tests in module pUnitTests_pUnit.TestModule04 due to BeforeModule failure"
  
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule04.BeforeTest(TestAddition)"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule04.TestAddition"
  AssertLogFileContains "Test skipped due to BeforeModule failure: pUnitTests_pUnit.TestModule04.TestAddition (Annotations: TestMethod, SmokeTest)"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule04.AfterTest(TestAddition)"
  
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule04.BeforeTest(TestDivision)"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule04.TestDivision"
  AssertLogFileContains "Test skipped due to BeforeModule failure: pUnitTests_pUnit.TestModule04.TestDivision (Annotations: TestMethod, Slow)"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule04.AfterTest(TestDivision)"
  
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule04.BeforeTest(TestFailure)"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule04.TestFailure"
  AssertLogFileContains "Test skipped due to BeforeModule failure: pUnitTests_pUnit.TestModule04.TestFailure (Annotations: TestMethod, Regression)"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule04.AfterTest(TestFailure)"
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule04.AfterModule"
  AssertLogFileDoesntContain "AfterModule failed in module TestModule04: Error #-2147220503: Intentional failure in AfterModule"
  
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule04: Setup Duration:"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule04 completed"
  AssertLogFileContains "Detailed Test Results:"
  
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule04.TestAddition, Status: Passed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule04.TestAddition, Status: Skipped, Message: Skipped due to BeforeModule failure"
  
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule04.TestDivision, Status: Passed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule04.TestDivision, Status: Skipped, Message: Skipped due to BeforeModule failure"
  
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule04.TestFailure, Status: Passed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule04.TestFailure, Status: Skipped, Message: Skipped due to BeforeModule failure"
  
  AssertLogFileContains "Total Test Method Duration (excluding setup/teardown):"
  AssertLogFileContains "Total Setup Duration:"
  AssertLogFileContains "Total Teardown Duration:"
  AssertLogFileContains "Total Run Duration:"
  AssertLogFileContains "Test execution completed."
  AssertLogFileContains "Total Tests: 3, Passed: 0, Failed: 0, Skipped: 3"
  
  LogSummary

End Sub

Private Sub RunUnitTestingTests08()
Arrange:
  StartTest "RunUnitTestingTests08"
Act:
  pUnit.TestRunner.RunAllTests "TestModule05", "", "pUnitTests_pUnit"
  GetLastCreatedPUnitLogFileContents
Assert:
  AssertLogFileContains "Target project name is: pUnitTests_pUnit"
  AssertLogFileContains "Starting unit test execution"
  AssertLogFileContains "Module filter: TestModule05"
  AssertLogFileContains "Annotation filter: All tests"
  
  AssertLogFileDoesntContain "Test skipped due to annotations at line"
  AssertLogFileContains "Processing module: pUnitTests_pUnit.TestModule05"
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule05.BeforeModule"
  AssertLogFileDoesntContain "BeforeModule failed in module pUnitTests_pUnit.TestModule05: Error #-2147220504: Intentional failure in BeforeModule"
  AssertLogFileDoesntContain "Skipping tests in module pUnitTests_pUnit.TestModule05 due to BeforeModule failure."
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule05.BeforeTest(TestAddition)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule05.TestAddition"
  AssertLogFileDoesntContain "Test skipped due to BeforeModule failure: pUnitTests_pUnit.TestModule05.TestAddition (Annotations: TestMethod, SmokeTest)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule05.AfterTest(TestAddition)"
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule05.BeforeTest(TestDivision)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule05.TestDivision"
  AssertLogFileDoesntContain "Test skipped due to BeforeModule failure: pUnitTests_pUnit.TestModule05.TestDivision (Annotations: TestMethod, Slow)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule05.AfterTest(TestDivision)"
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule05.BeforeTest(TestFailure)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule05.TestFailure"
  AssertLogFileDoesntContain "Test skipped due to BeforeModule failure: pUnitTests_pUnit.TestModule05.TestFailure (Annotations: TestMethod, Regression)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule05.AfterTest(TestFailure)"
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule05.AfterModule"
  AssertLogFileContains "AfterModule failed in module pUnitTests_pUnit.TestModule05: Error #-2147220503: Intentional failure in AfterModule"
  
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule05: Setup Duration:"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule05 completed"
  AssertLogFileContains "Detailed Test Results:"
  
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule05.TestAddition, Status: Passed"
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule05.TestAddition, Status: Skipped, Message: Skipped due to BeforeModule failure"
  
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule05.TestDivision, Status: Passed"
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule05.TestDivision, Status: Skipped, Message: Skipped due to BeforeModule failure"
  
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule05.TestFailure, Status: Passed"
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule05.TestFailure, Status: Skipped, Message: Skipped due to BeforeModule failure"
  
  AssertLogFileContains "Total Test Method Duration (excluding setup/teardown):"
  AssertLogFileContains "Total Setup Duration:"
  AssertLogFileContains "Total Teardown Duration:"
  AssertLogFileContains "Total Run Duration:"
  AssertLogFileContains "Test execution completed."
  AssertLogFileContains "Total Tests: 3, Passed: 3, Failed: 0, Skipped: 0"
  
  LogSummary

End Sub

Private Sub RunUnitTestingTests09()
Arrange:
  StartTest "RunUnitTestingTests09"
Act:
  pUnit.TestRunner.RunAllTests "TestModule06", "", "pUnitTests_pUnit"
  GetLastCreatedPUnitLogFileContents
Assert:
  AssertLogFileContains "Target project name is: pUnitTests_pUnit"
  AssertLogFileContains "Starting unit test execution"
  AssertLogFileContains "Module filter: TestModule06"
  AssertLogFileContains "Annotation filter: All tests"
  
  AssertLogFileDoesntContain "Test skipped due to annotations at line"
  AssertLogFileContains "Processing module: pUnitTests_pUnit.TestModule06"
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule06.BeforeModule"
  AssertLogFileDoesntContain "BeforeModule failed in module pUnitTests_pUnit.TestModule06: Error #-2147220504: Intentional failure in BeforeModule"
  AssertLogFileDoesntContain "Skipping tests in module pUnitTests_pUnit.TestModule06 due to BeforeModule failure."
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule06.BeforeTest(TestAddition)"
  AssertLogFileContains "BeforeTest(TestAddition) failed in module pUnitTests_pUnit.TestModule06: Error #-2147220501: Intentional failure in BeforeTest"
  AssertLogFileContains "Test skipped due to BeforeTest failure: pUnitTests_pUnit.TestModule06.TestAddition"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule06.TestAddition"
  AssertLogFileDoesntContain "Test skipped due to BeforeModule failure: pUnitTests_pUnit.TestModule06.TestAddition (Annotations: TestMethod, SmokeTest)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule06.AfterTest(TestAddition)"
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule06.BeforeTest(TestDivision)"
  AssertLogFileContains "BeforeTest(TestDivision) failed in module pUnitTests_pUnit.TestModule06: Error #-2147220501: Intentional failure in BeforeTest"
  AssertLogFileContains "Test skipped due to BeforeTest failure: pUnitTests_pUnit.TestModule06.TestDivision"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule06.TestDivision"
  AssertLogFileDoesntContain "Test skipped due to BeforeModule failure: pUnitTests_pUnit.TestModule06.TestDivision (Annotations: TestMethod, Slow)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule06.AfterTest(TestDivision)"
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule06.BeforeTest(TestFailure)"
  AssertLogFileContains "BeforeTest(TestFailure) failed in module pUnitTests_pUnit.TestModule06: Error #-2147220501: Intentional failure in BeforeTest"
  AssertLogFileContains "Test skipped due to BeforeTest failure: pUnitTests_pUnit.TestModule06.TestFailure"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule06.TestFailure"
  AssertLogFileDoesntContain "Test skipped due to BeforeModule failure: pUnitTests_pUnit.TestModule06.TestFailure (Annotations: TestMethod, Regression)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule06.AfterTest(TestFailure)"
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule06.AfterModule"
  AssertLogFileDoesntContain "AfterModule failed in module pUnitTests_pUnit.TestModule06: Error #-2147220503: Intentional failure in AfterModule"
  
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule06: Setup Duration:"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule06 completed"
  AssertLogFileContains "Detailed Test Results:"
  
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule06.TestAddition, Status: Passed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule06.TestAddition, Status: Skipped, Message: Skipped due to BeforeTest failure"
  
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule06.TestDivision, Status: Passed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule06.TestDivision, Status: Skipped, Message: Skipped due to BeforeTest failure"
  
  AssertLogFileDoesntContain "Test: pUnitTests_pUnit.TestModule06.TestFailure, Status: Passed"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule06.TestFailure, Status: Skipped, Message: Skipped due to BeforeTest failure"
  
  AssertLogFileContains "Total Test Method Duration (excluding setup/teardown):"
  AssertLogFileContains "Total Setup Duration:"
  AssertLogFileContains "Total Teardown Duration:"
  AssertLogFileContains "Total Run Duration:"
  AssertLogFileContains "Test execution completed."
  AssertLogFileContains "Total Tests: 3, Passed: 0, Failed: 0, Skipped: 3"
  
  LogSummary

End Sub

Private Sub RunUnitTestingTests10()
Arrange:
  StartTest "RunUnitTestingTests10"
Act:
  pUnit.TestRunner.RunAllTests "TestModule07", "", "pUnitTests_pUnit"
  GetLastCreatedPUnitLogFileContents
Assert:
  AssertLogFileContains "Target project name is: pUnitTests_pUnit"
  AssertLogFileContains "Starting unit test execution"
  AssertLogFileContains "Module filter: TestModule07"
  AssertLogFileContains "Annotation filter: All tests"
  
  AssertLogFileDoesntContain "Test skipped due to annotations at line"
  AssertLogFileContains "Processing module: pUnitTests_pUnit.TestModule07"
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule07.BeforeModule"
  AssertLogFileDoesntContain "BeforeModule failed in module TestModule07: Error #-2147220504: Intentional failure in BeforeModule"
  AssertLogFileDoesntContain "Skipping tests in module TestModule07 due to BeforeModule failure."
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule07.BeforeTest(TestAddition)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule07.TestAddition"
  AssertLogFileDoesntContain "Test skipped due to BeforeModule failure: TestModule07.TestAddition (Annotations: TestMethod, SmokeTest)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule07.AfterTest(TestAddition)"
  AssertLogFileContains "AfterTest(TestAddition) failed in module TestModule07: Error #-2147220500: Intentional failure in AfterTest"
                         
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule07.BeforeTest(TestDivision)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule07.TestDivision"
  AssertLogFileDoesntContain "Test skipped due to BeforeModule failure: TestModule07.TestDivision (Annotations: TestMethod, Slow)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule07.AfterTest(TestDivision)"
  AssertLogFileContains "AfterTest(TestDivision) failed in module TestModule07: Error #-2147220500: Intentional failure in AfterTest"
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule07.BeforeTest(TestFailure)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule07.TestFailure"
  AssertLogFileDoesntContain "Test skipped due to BeforeModule failure: TestModule07.TestFailure (Annotations: TestMethod, Regression)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule07.AfterTest(TestFailure)"
  AssertLogFileContains "AfterTest(TestFailure) failed in module TestModule07: Error #-2147220500: Intentional failure in AfterTest"
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule07.AfterModule"
  AssertLogFileDoesntContain "AfterModule failed in module TestModule07: Error #-2147220503: Intentional failure in AfterModule"
  
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule07: Setup Duration:"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule07 completed"
  AssertLogFileContains "Detailed Test Results:"
  
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule07.TestAddition, Status: Passed"
  AssertLogFileDoesntContain "Test: TestModule07.TestAddition, Status: Skipped, Message: Skipped due to BeforeModule failure"
  
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule07.TestDivision, Status: Passed"
  AssertLogFileDoesntContain "Test: TestModule07.TestDivision, Status: Skipped, Message: Skipped due to BeforeModule failure"
  
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule07.TestFailure, Status: Passed"
  AssertLogFileDoesntContain "Test: TestModule07.TestFailure, Status: Skipped, Message: Skipped due to BeforeModule failure"
  
  AssertLogFileContains "Total Test Method Duration (excluding setup/teardown):"
  AssertLogFileContains "Total Setup Duration:"
  AssertLogFileContains "Total Teardown Duration:"
  AssertLogFileContains "Total Run Duration:"
  AssertLogFileContains "Test execution completed."
  AssertLogFileContains "Total Tests: 3, Passed: 3, Failed: 0, Skipped: 0"
  
  LogSummary

End Sub

'Invalid test module filter
Private Sub RunUnitTestingTests11()
Arrange:
  StartTest "RunUnitTestingTests11"
Act:
  pUnit.TestRunner.RunAllTests "XXX(TestModule08,)", "", "pUnitTests_pUnit"
  GetLastCreatedPUnitLogFileContents
Assert:
  AssertLogFileContains "Invalid moduleName filter syntax: XXX(TestModule08,) (Error: Type mismatch)"
  AssertLogFileDoesntContain "Calling project name is: pUnitTestRuns_pUnit"
  AssertLogFileDoesntContain "Starting unit test execution"
  AssertLogFileDoesntContain "Module filter: TestModule08"
  AssertLogFileDoesntContain "Annotation filter: All tests"
  LogSummary

End Sub

'TODO: Invalid test annotation filter???

'Use the target Test Name as the annotation filter
Private Sub RunUnitTestingTests12()
Arrange:
  StartTest "RunUnitTestingTests12"
Act:
  pUnit.TestRunner.RunAllTests "OR(TestModule08,TestModule09)", "TestMultiplication", "pUnitTests_pUnit"
  GetLastCreatedPUnitLogFileContents
Assert:
  AssertLogFileContains "Target project name is: pUnitTests_pUnit"
  AssertLogFileContains "Starting unit test execution"
  AssertLogFileContains "Module filter: OR(TestModule08,TestModule09)"
  AssertLogFileContains "Annotation filter: TestMultiplication"

  AssertLogFileContains "Test skipped due to annotations at line 30: pUnitTests_pUnit.TestModule08.TestAddition (Annotations: or Procedure Name: TestAddition)"
  AssertLogFileContains "Test skipped due to annotations at line 52: pUnitTests_pUnit.TestModule08.TestDivision (Annotations: or Procedure Name: TestDivision)"
  AssertLogFileContains "Test skipped due to annotations at line 63: pUnitTests_pUnit.TestModule08.TestFailure (Annotations: or Procedure Name: TestFailure)"

  AssertLogFileContains "Processing module: pUnitTests_pUnit.TestModule08"

  AssertLogFileContains "Running pUnitTests_pUnit.TestModule08.BeforeModule"
  AssertLogFileContains "BeforeModule failed in module pUnitTests_pUnit.TestModule08: Error #-2147220504: Intentional failure in BeforeModule"
  AssertLogFileContains "Skipping tests in module pUnitTests_pUnit.TestModule08 due to BeforeModule failure"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule08.BeforeTest(TestMultiplication)"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule08.TestMultiplication"
  AssertLogFileContains "Test skipped due to BeforeModule failure: pUnitTests_pUnit.TestModule08.TestMultiplication (Annotations: TestMethod)"
  AssertLogFileDoesntContain "Running pUnitTests_pUnit.TestModule08.AfterTest(TestMultiplication)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule08.AfterModule"

  AssertLogFileContains "Module pUnitTests_pUnit.TestModule08: Setup Duration:"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule08 completed"

  AssertLogFileContains "Test skipped due to annotations at line 43: pUnitTests_pUnit.TestModule09.TestSubtraction (Annotations: or Procedure Name: TestSubtraction)"

  AssertLogFileContains "Processing module: pUnitTests_pUnit.TestModule09"

  AssertLogFileContains "Running pUnitTests_pUnit.TestModule09.BeforeModule"
  AssertLogFileDoesntContain "BeforeModule failed in module pUnitTests_pUnit.TestModule09: Error #-2147220504: Intentional failure in BeforeModule"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule09.BeforeTest(TestMultiplication)"
  AssertLogFileDoesntContain "Skipping tests in module pUnitTests_pUnit.TestModule09 due to BeforeModule failure."
  AssertLogFileContains "BeforeTest(TestMultiplication) failed in module pUnitTests_pUnit.TestModule09: Error #-2147220502: Intentional failure in BeforeTest for TestMultiplication"
  AssertLogFileContains "Test skipped due to BeforeTest failure: pUnitTests_pUnit.TestModule09.TestMultiplication (Annotations: TestMethod)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule09.AfterTest(TestMultiplication)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule09.AfterModule"

  AssertLogFileContains "Module pUnitTests_pUnit.TestModule09: Setup Duration:"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule09 completed"

  AssertLogFileContains "Detailed Test Results:"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule08.TestMultiplication, Status: Skipped, Message: Skipped due to BeforeModule failure, Duration: 0.000 ms"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule09.TestMultiplication, Status: Skipped, Message: Skipped due to BeforeTest failure, Duration: 0.000 ms"

  AssertLogFileContains "Total Test Method Duration (excluding setup/teardown):"
  AssertLogFileContains "Total Setup Duration:"
  AssertLogFileContains "Total Teardown Duration:"
  AssertLogFileContains "Total Run Duration:"
  AssertLogFileContains "Test execution completed."
  AssertLogFileContains "Total Tests: 2, Passed: 0, Failed: 0, Skipped: 2"
  
  LogSummary

End Sub

'Check for test data - this uses a module level annotation to run all tests
Private Sub RunUnitTestingTests13()
Arrange:
  StartTest "RunUnitTestingTests13"
Act:
  pUnit.TestRunner.RunAllTests "", "TestingTestData", "pUnitTests_pUnit"
  GetLastCreatedPUnitLogFileContents
Assert:

  AssertLogFileContains "Target project name is: pUnitTests_pUnit"
  AssertLogFileContains "Starting unit test execution"
  AssertLogFileContains "Module filter: All modules"
  AssertLogFileContains "Annotation filter: TestingTestData"
  
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule01 has no tests to run after filtering."
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule02 has no tests to run after filtering."
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule03 has no tests to run after filtering."
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule04 has no tests to run after filtering."
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule05 has no tests to run after filtering."
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule06 has no tests to run after filtering."
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule07 has no tests to run after filtering."
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule08 has no tests to run after filtering."
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule09 has no tests to run after filtering."
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule10 has no tests to run after filtering."
  
  AssertLogFileContains "Processing module: pUnitTests_pUnit.TestModule11"
  
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule11.TestAdditionNotParameterised"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionNotParameterised"

  AssertLogFileContains "Invalid @TestData format for test TestModule11.TestAdditionParameterised01: @TestData (Expected '@TestData(...)')"
  AssertLogFileContains "Test skipped due to missing test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised01"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised01, Status: Skipped, Message: No test data provided for parameterized test expecting 3 parameters, Duration:"

  AssertLogFileContains "Empty @TestData for test TestModule11.TestAdditionParameterised02: @TestData()"
  AssertLogFileContains "Test skipped due to missing test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised02"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised02, Status: Skipped, Message: No test data provided for parameterized test expecting 3 parameters, Duration:"

  AssertLogFileContains "Malformed @TestData for test TestModule11.TestAdditionParameterised03: missing or unbalanced parentheses in @TestData({}"
  AssertLogFileContains "Test skipped due to missing test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised03"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised03, Status: Skipped, Message: No test data provided for parameterized test expecting 3 parameters, Duration:"

  AssertLogFileContains "Malformed @TestData for test TestModule11.TestAdditionParameterised04: missing or unbalanced parentheses in @TestData("
  AssertLogFileContains "Test skipped due to missing test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised04"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised04, Status: Skipped, Message: No test data provided for parameterized test expecting 3 parameters, Duration:"

  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised07, Status: Skipped, Message: No test data provided for parameterized test expecting 3 parameters, Duration:"

  AssertLogFileContains "Malformed @TestData for test TestModule11.TestAdditionParameterised05: missing or unbalanced parentheses in @TestData)"
  AssertLogFileContains "Test skipped due to missing test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised05"

  AssertLogFileContains "Malformed @TestData for test TestModule11.TestAdditionParameterised06: missing or unbalanced parentheses in @TestData)("
  AssertLogFileContains "Test skipped due to missing test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised06"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised06, Status: Skipped, Message: No test data provided for parameterized test expecting 3 parameters, Duration:"

  AssertLogFileContains "Empty @TestData for test TestModule11.TestAdditionParameterised07: @TestData()"
  AssertLogFileContains "Test skipped due to missing test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised07"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised05, Status: Skipped, Message: No test data provided for parameterized test expecting 3 parameters, Duration:"

  AssertLogFileContains "No valid test data parsed for test TestModule11.TestAdditionParameterised08: @TestData({})"
  AssertLogFileContains "Test skipped due to missing test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised08"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised08, Status: Skipped, Message: No test data provided for parameterized test expecting 3 parameters, Duration:"

  AssertLogFileContains "Invalid data set skipped for test TestModule11.TestAdditionParameterised09: {}"
  AssertLogFileContains "No valid test data parsed for test TestModule11.TestAdditionParameterised09: @TestData({{}})"
  AssertLogFileContains "Test skipped due to missing test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised09"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised09, Status: Skipped, Message: No test data provided for parameterized test expecting 3 parameters, Duration:"

  AssertLogFileContains "Empty parameter string in @TestData for test TestModule11.TestAdditionParameterised10: {}"
  AssertLogFileContains "Invalid data set skipped for test TestModule11.TestAdditionParameterised10: { }"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule11.TestAdditionParameterised10(1, 2, 3)"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised10(1, 2, 3) (Parameters: 1, 2, 3), Status: Passed, Duration:"

  AssertLogFileContains "Unclosed brace in @TestData for test TestModule11.TestAdditionParameterised11: @TestData({)"
  AssertLogFileContains "No valid test data parsed for test TestModule11.TestAdditionParameterised11: @TestData({)"
  AssertLogFileContains "Test skipped due to missing test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised11"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised11, Status: Skipped, Message: No test data provided for parameterized test expecting 3 parameters, Duration:"

  AssertLogFileContains "Malformed @TestData for test TestModule11.TestAdditionParameterised12: missing or unbalanced parentheses in @TestData({1, 2, 3}"
  AssertLogFileContains "Test skipped due to missing test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised12"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised12, Status: Skipped, Message: No test data provided for parameterized test expecting 3 parameters, Duration:"

  AssertLogFileContains "Running pUnitTests_pUnit.TestModule11.TestAdditionParameterised13(1, 2, 3)"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised13(1, 2, 3) (Parameters: 1, 2, 3), Status: Passed, Duration:"

  AssertLogFileContains "Test skipped due to parameter count mismatch: pUnitTests_pUnit.TestModule11.TestAdditionParameterised14(1, 2)"
  AssertLogFileContains "Test skipped due to no valid test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised14"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised14(1, 2), Status: Skipped, Message: Parameter count mismatch: expected 3, got 2, Duration:"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised14, Status: Skipped, Message: No valid test data sets found, Duration:"

  AssertLogFileContains "Running pUnitTests_pUnit.TestModule11.TestAdditionParameterised15(1, 2, 3)"
  AssertLogFileContains "Test skipped due to parameter count mismatch: pUnitTests_pUnit.TestModule11.TestAdditionParameterised15(3, 4)"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised15(1, 2, 3) (Parameters: 1, 2, 3), Status: Passed, Duration:"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised15(3, 4), Status: Skipped, Message: Parameter count mismatch: expected 3, got 2, Duration:"

  AssertLogFileContains "Running pUnitTests_pUnit.TestModule11.TestAdditionParameterised16(1, 2, 3)"
  AssertLogFileContains "Test skipped due to parameter count mismatch: pUnitTests_pUnit.TestModule11.TestAdditionParameterised16(3, 4, 7, 8)"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised16(1, 2, 3) (Parameters: 1, 2, 3), Status: Passed, Duration:"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised16(3, 4, 7, 8), Status: Skipped, Message: Parameter count mismatch: expected 3, got 4, Duration:"

  AssertLogFileContains "Running pUnitTests_pUnit.TestModule11.TestAdditionParameterised17(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30)"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised17(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30) (Parameters: 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30), Status: Passed, Duration:"

  AssertLogFileContains "Too many parameters for test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised18(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31) (Max supported: 30)"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised18(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31) (Parameters: 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31), Status: Skipped, Message: Too many parameters (Max supported: 30), Duration:"

  AssertLogFileContains "Malformed @TestData for test TestModule11.TestAdditionParameterised19: missing or unbalanced parentheses in @TestData({{1, 2, 3},"
  AssertLogFileContains "Test skipped due to missing test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised19"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised19, Status: Skipped, Message: No test data provided for parameterized test expecting 3 parameters, Duration:"

  AssertLogFileContains "Type mismatch for parameter 2 in test TestModule11.TestAdditionParameterised20: expected Integer/Long, got String"
  AssertLogFileContains "Test skipped due to type mismatch: pUnitTests_pUnit.TestModule11.TestAdditionParameterised20(1, Two, 3)"
  AssertLogFileContains "Test skipped due to no valid test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised20"

  AssertLogFileContains "Type mismatch for parameter 2 in test TestModule11.TestAdditionParameterised21: expected String, got Integer"
  AssertLogFileContains "Type mismatch for parameter 3 in test TestModule11.TestAdditionParameterised21: expected Integer/Long, got String"
  AssertLogFileContains "Type mismatch for parameter 4 in test TestModule11.TestAdditionParameterised21: expected Integer/Long, got Double"
  AssertLogFileContains "Type mismatch for parameter 5 in test TestModule11.TestAdditionParameterised21: expected Double, got Boolean"
  AssertLogFileContains "Type mismatch for parameter 6 in test TestModule11.TestAdditionParameterised21: expected Boolean, got String"
  AssertLogFileContains "Test skipped due to type mismatch: pUnitTests_pUnit.TestModule11.TestAdditionParameterised21(1, 2, 3, 1.5, False, TRUE)"
  AssertLogFileContains "Test skipped due to no valid test data: pUnitTests_pUnit.TestModule11.TestAdditionParameterised21"

  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised21(1, 2, 3, 1.5, False, TRUE), Status: Skipped, Message: Type mismatch in parameters, Duration:"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised21, Status: Skipped, Message: No valid test data sets found, Duration:"

  AssertLogFileContains "Running pUnitTests_pUnit.TestModule11.TestAdditionParameterised99(1, 2, 3)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule11.TestAdditionParameterised99(3, 4, 7)"
  AssertLogFileContains "Running pUnitTests_pUnit.TestModule11.TestAdditionParameterised99(0, 0, 0)"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised99(1, 2, 3) (Parameters: 1, 2, 3), Status: Passed, Duration:"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised99(3, 4, 7) (Parameters: 3, 4, 7), Status: Passed, Duration:"
  AssertLogFileContains "Test: pUnitTests_pUnit.TestModule11.TestAdditionParameterised99(0, 0, 0) (Parameters: 0, 0, 0), Status: Passed, Duration:"

  AssertLogFileContains "Module pUnitTests_pUnit.TestModule11: Setup Duration: 0.000 ms, Teardown Duration: 0.000 ms"
  AssertLogFileContains "Module pUnitTests_pUnit.TestModule11 completed"
  AssertLogFileContains "Detailed Test Results:"
  AssertLogFileContains "Total Test Method Duration (excluding setup/teardown):"
  AssertLogFileContains "Total Setup Duration: 0.000 ms"
  AssertLogFileContains "Total Teardown Duration: 0.000 ms"
  AssertLogFileContains "Total Run Duration:"
  AssertLogFileContains "Test execution completed."
  AssertLogFileContains "Total Tests: 31, Passed: 10, Failed: 0, Skipped: 21"
  
  LogSummary

End Sub

'IGNORING COMPLEX ANNOTATION TAGS FOR NOW - DO THEY ADD ANY VALUE?
'Complex Annotation Tag Formats:
'Support advanced annotation syntax, including:
' - Simple tags: @SmokeTest, @Tag(SmokeTest, Regression).
' - Quoted strings: @Tag("Smoke Test", "Regression Suite").
' - Key-value pairs: @Tag(category=SmokeTest, priority=High).
' - Nested tags: @Tag(category={SmokeTest, Regression}).
' - Mixed formats: @Tag(SmokeTest, category=Regression, "High Priority").
'Normalize tags for case-insensitive comparison (e.g., SmokeTest matches smoketest)
'Extract tags for filtering tests and modules using includeAnnotations and excludeAnnotations
'Handle malformed annotations gracefully (log warnings, skip invalid tags).

Sub StartTest(TestName As String)
  Phosphorus.AssertionsStatic.GetAssert
  Set pAssert = Nothing
  lstrTestName = TestName
  Debug.Print "Running " & lstrTestName
  WriteToLog "Running " & lstrTestName
End Sub

Private Sub WriteToLog(strMessage As String)
  If lboolRunnningAllTests Then
    thisLogger.ExternalInfo strMessage
  Else
    Debug.Print strMessage
  End If
End Sub

Sub GetLastCreatedPUnitLogFileContents()
  Dim FSO As FileSystemObject
  Dim strLogFilePath As String
  Dim FileToRead As TextStream
  Set FSO = New FileSystemObject
  strLogFilePath = pUnit.LastCreatedPUnitLogFile
  If strLogFilePath <> "" Then
    Set FileToRead = FSO.OpenTextFile(strLogFilePath, ForReading)
    LastCreatedPUnitLogFileContents = FileToRead.ReadAll
    FileToRead.Close
  End If
  lintCountOfAssertionPasses = 0
  lintCountOfAssertionFails = 0
  If Not lboolRunningAllTests Then
    lintCountOfTestFails = 0
  End If
End Sub

Sub AssertLogFileContains(TextContent As String)
  WriteToLog "AssertLogFileContains: " & TextContent
  If VBA.InStr(1, LastCreatedPUnitLogFileContents, TextContent) = 0 Then
    WriteToLog "FAIL: The log file doesn't contain: " & TextContent
    lintCountOfAssertionFails = lintCountOfAssertionFails + 1
  Else
    WriteToLog "PASS: The log file contains: " & TextContent
    lintCountOfAssertionPasses = lintCountOfAssertionPasses + 1
  End If
End Sub

Sub AssertLogFileDoesntContain(TextContent As String)
  WriteToLog "AssertLogFileDoesntContain: " & TextContent
  If VBA.InStr(1, LastCreatedPUnitLogFileContents, TextContent) = 0 Then
    WriteToLog "PASS: The log file doesn't contain: " & TextContent
    lintCountOfAssertionPasses = lintCountOfAssertionPasses + 1
  Else
    WriteToLog "FAIL: The log file does contain: " & TextContent
    lintCountOfAssertionFails = lintCountOfAssertionFails + 1
  End If
End Sub

Sub LogSummary()
  WriteToLog lintCountOfAssertionPasses & " asssertions PASSED."
  WriteToLog lintCountOfAssertionFails & " asssertions FAILED."
  If lintCountOfAssertionFails = 0 Then
    WriteToLog "Test: " & lstrTestName & " PASSED!"
  Else
    lintCountOfTestFails = lintCountOfTestFails + 1
    WriteToLog "Test: " & lstrTestName & " FAILED!"
    Debug.Print "Test: " & lstrTestName & " FAILED!"
    If Not lboolRunningAllTests Then
      MsgBox "Test: " & lstrTestName & " FAILED!"
    End If
  End If
  Phosphorus.AssertionsStatic.CloseAssert
End Sub
