Attribute VB_Name = "TestRunner"
'@Folder pUnit
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

'Enable VBA Project Access:In Excel, go to File > Options > Trust Center > Trust Center Settings > Macro Settings.
'Check "Trust access to the VBA project object model."

'Add Reference to VBIDE:In the VBA Editor (Alt+F11), go to Tools > References.
'Check "Microsoft Visual Basic for Applications Extensibility 5.3" (or similar).

' Windows API declarations for high-resolution timing
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

'' Reference to Assertions and Logging classes
'Private Assertions As Phosphorus.Assertions

' Collection to store test results
Private TestResults As collection

' Total duration of test methods (excluding setup/teardown)
Private TotalTestMethodDuration As Double

' Total setup and teardown durations across all modules
Private TotalSetupDuration As Double
Private TotalTeardownDuration As Double

' Cache for method signatures to optimize parsing
Private MethodSignatures As Object ' Scripting.Dictionary

'Last created log file for unit testing of pUnit itself
Public LastCreatedPUnitLogFile As String

' Main entry point to run tests
Public Sub RunAllTests(moduleNameFilter As String, annotationFilter As String, Optional targetProjectName As String = "")

  Phosphorus.Log4PStatic.GetLogger
  
  Set MethodSignatures = CreateObject("Scripting.Dictionary") ' Initialize signature cache
  TotalTestMethodDuration = 0
  TotalSetupDuration = 0
  TotalTeardownDuration = 0
  Dim vbProj As VBProject
  Dim passCount As Long
  Dim skipCount As Long
  Dim projectName As String
  Dim executableTests As collection
  Dim moduleSetupSuccess As Boolean
  Dim TestResult As pUnit.TestResult
  Dim TestName As Variant
  Dim runStartCount As Currency
  Dim runEndCount As Currency
  Dim runDuration As Double
  Dim moduleStartCount As Currency
  Dim moduleEndCount As Currency
  Dim moduleDuration As Double
  Dim moduleSetupDuration As Double
  Dim moduleTeardownDuration As Double
    
  If Not ValidateFilterSyntax(moduleNameFilter, "moduleName") Then
    GoTo ExitSub
  End If
  
  If Not ValidateFilterSyntax(annotationFilter, "annotation") Then
    GoTo ExitSub
  End If
    
  QueryPerformanceCounter runStartCount

  'Get project name
  If Len(targetProjectName) > 0 Then
    Dim vbTestProj As VBProject
    For Each vbTestProj In Application.VBE.VBProjects
      'Debug.Print vbTestProj.Name
      If vbTestProj.Name = targetProjectName Then
        Set vbProj = vbTestProj
        Exit For
      End If
    Next
    projectName = targetProjectName
  Else
    'Set the active project as target project
    On Error Resume Next
    Set vbProj = Application.VBE.ActiveVBProject
    If Not vbProj Is Nothing Then
      projectName = vbProj.Name
    End If
  End If
   
  'Check we got a project name
  On Error GoTo 0
  If Len(projectName) = 0 Then
    Phosphorus.Log4PStatic.Logger.ExternalError "Could not determine the calling project name and no referencedProjectName provided."
    GoTo ExitSub
  End If
  Phosphorus.Log4PStatic.Logger.ExternalInfo "Target project name is: " & projectName

  'Set project reference
  On Error Resume Next
  Set vbProj = Application.VBE.VBProjects(projectName)
  On Error GoTo 0
  If vbProj Is Nothing Then
    Phosphorus.Log4PStatic.Logger.ExternalError "Project '" & projectName & "' not found."
    GoTo ExitSub
  End If

  Phosphorus.Log4PStatic.Logger.ExternalInfo "Starting unit test execution"
  Phosphorus.Log4PStatic.Logger.ExternalInfo "Module filter: " & IIf(Len(moduleNameFilter) > 0, moduleNameFilter, "All modules")
  Phosphorus.Log4PStatic.Logger.ExternalInfo "Annotation filter: " & IIf(Len(annotationFilter) > 0, annotationFilter, "All tests")

  'Make a collection of all test module names
  Dim allTestModuleNames As collection
  Set allTestModuleNames = New collection
  Dim vbCompnnt As VBComponent
  For Each vbCompnnt In vbProj.VBComponents
    If vbCompnnt.Type = vbext_ct_StdModule And IsTestModule(vbCompnnt) Then
      allTestModuleNames.Add vbCompnnt.Name
    End If
  Next vbCompnnt
  
  'Initialise a collection of test results
  Set TestResults = New collection
  
  'Loop through all the test modules
  Dim testModuleName As Variant
  For Each testModuleName In allTestModuleNames
  
    'Get a reference to the target test module
    Dim vbTestModule As VBComponent
    Set vbTestModule = vbProj.VBComponents(testModuleName)
    
    'Is this a module that we should include?
    If ShouldIncludeTestModule(vbTestModule, moduleNameFilter, allTestModuleNames) Then
    
      'Start a new collection of executable tests
      Set executableTests = New collection
      
      'Does the test module have any tests that match the annotations filter
      If HasTestsToRun(projectName, vbTestModule, annotationFilter, executableTests) Then
        
        'Process the module setup
        Phosphorus.Log4PStatic.Logger.ExternalInfo "Processing module: " & projectName & "." & vbTestModule.Name
        moduleSetupDuration = 0
        moduleTeardownDuration = 0
        QueryPerformanceCounter moduleStartCount
        moduleSetupSuccess = ExecuteModuleSetup(projectName, vbProj, vbTestModule, moduleSetupDuration)
        
        'Module setup ok?
        If moduleSetupSuccess Then
          
          'Module setup ran ok, now run each test
          For Each TestName In executableTests
          
            'Store the current test name for use in Before/After Test methods, for example
            Phosphorus.pUnit.CurrentTestName = VBA.Conversion.CStr(TestName)
            
            'Run the test setup
            Dim testSetupSuccess As Boolean
            testSetupSuccess = ExecuteTestSetup(VBA.Conversion.CStr(TestName), vbProj, vbTestModule, moduleSetupDuration)
            
            'Test setup ok?
            If testSetupSuccess Then

              ' Get test data for parameterized tests
              Dim testData As collection
              Set testData = GetTestData(vbTestModule, VBA.Conversion.CStr(TestName))
                       
              'Get the test method's signature
              Dim signature As Variant
              signature = GetProcedureSignature(vbTestModule.CodeModule, VBA.Conversion.CStr(TestName))
              
              ' Validate the parameter count
              Dim paramCount As Long
              Dim paramTypes As Variant
              paramCount = signature(0)
              paramTypes = signature(1)
              Dim longCountOfTestData As Long
              If (testData Is Nothing) Then
                longCountOfTestData = 0
              Else
                longCountOfTestData = testData.count
              End If

              If longCountOfTestData > 0 Then
                ' Run parameterized test for each data set
                Dim dataSet As Variant
                Dim validDataSets As Long
                validDataSets = 0
                For Each dataSet In testData
                  If IsArray(dataSet) Then
                    If UBound(dataSet) + 1 = paramCount Then
                      Dim typeValid As Boolean
                      typeValid = ValidateParameterTypes(dataSet, paramTypes, vbTestModule.Name & "." & TestName)
                      If typeValid Then
                        Set TestResult = ExecuteTest(vbProj, vbTestModule, VBA.Conversion.CStr(TestName), dataSet)
                        TestResults.Add TestResult
                        validDataSets = validDataSets + 1
                      Else
                        Set TestResult = New pUnit.TestResult
                        'testResult.testName = vbComp.Name & "." & testName & "(" & Join(CollectionToArray(dataSet), ", ") & ")"
                        'PJG dataSet already is an array so doesn't need converting!
                        TestResult.TestName = projectName & "." & vbTestModule.Name & "." & TestName & "(" & VBA.Strings.Join(dataSet, ", ") & ")"
                        TestResult.Status = skipped
                        TestResult.ErrorMessage = "Type mismatch in parameters"
                        TestResult.duration = 0
                        TestResults.Add TestResult
                        Phosphorus.Log4PStatic.Logger.ExternalInfo "Test skipped due to type mismatch: " & TestResult.TestName
                      End If
                    Else
                      Set TestResult = New pUnit.TestResult
                      'testResult.testName = vbComp.Name & "." & testName & "(" & Join(CollectionToArray(dataSet), ", ") & ")"
                      'PJG dataSet already is an array so doesn't need converting!
                      TestResult.TestName = projectName & "." & vbTestModule.Name & "." & TestName & "(" & VBA.Strings.Join(dataSet, ", ") & ")"
                      TestResult.Status = skipped
                      TestResult.ErrorMessage = "Parameter count mismatch: expected " & paramCount & ", got " & (UBound(dataSet) + 1)
                      TestResult.duration = 0
                      TestResults.Add TestResult
                      Phosphorus.Log4PStatic.Logger.ExternalInfo "Test skipped due to parameter count mismatch: " & TestResult.TestName
                    End If
                  Else
                    Set TestResult = New pUnit.TestResult
                    TestResult.TestName = projectName & "." & vbTestModule.Name & "." & TestName
                    TestResult.Status = skipped
                    TestResult.ErrorMessage = "Invalid test data set: " & typeName(dataSet)
                    TestResult.duration = 0
                    TestResults.Add TestResult
                    Phosphorus.Log4PStatic.Logger.ExternalInfo "Test skipped due to invalid test data: " & TestResult.TestName
                  End If
                Next dataSet
                If validDataSets = 0 Then
                  Set TestResult = New pUnit.TestResult
                  TestResult.TestName = projectName & "." & vbTestModule.Name & "." & TestName
                  TestResult.Status = skipped
                  TestResult.ErrorMessage = "No valid test data sets found"
                  TestResult.duration = 0
                  TestResults.Add TestResult
                  Phosphorus.Log4PStatic.Logger.ExternalInfo "Test skipped due to no valid test data: " & TestResult.TestName
                End If
              Else
                'No Test Data found
                ' Does the current test method have any parameters?
                If paramCount = 0 Then
                  'No test data found & current test method has no parameters so test method is ok to run
                  Set TestResult = ExecuteTest(vbProj, vbTestModule, VBA.Conversion.CStr(TestName))
                  'Store the test result
                  TestResults.Add TestResult
                Else
                  'No test data  found but current test method has parameters so test method is not ok to run
                  Set TestResult = New pUnit.TestResult
                  TestResult.TestName = projectName & "." & vbTestModule.Name & "." & TestName
                  TestResult.Status = skipped
                  TestResult.ErrorMessage = "No test data provided for parameterized test expecting " & paramCount & " parameters"
                  TestResult.duration = 0
                  TestResults.Add TestResult
                  Phosphorus.Log4PStatic.Logger.ExternalInfo "Test skipped due to missing test data: " & TestResult.TestName
                End If
              End If
            Else
              Set TestResult = New pUnit.TestResult
              TestResult.TestName = projectName & "." & vbTestModule.Name & "." & TestName
              TestResult.Status = skipped
              TestResult.ErrorMessage = "Skipped due to BeforeTest failure"
              TestResult.duration = 0
              TestResults.Add TestResult
              Phosphorus.Log4PStatic.Logger.ExternalInfo "Test skipped due to BeforeTest failure: " & TestResult.TestName & " (Annotations: " & GetTestAnnotations(vbTestModule, VBA.Conversion.CStr(TestName)) & ")"
            End If
            
            'Run the test teardown
            ExecuteTestTeardown VBA.Conversion.CStr(TestName), vbProj, vbTestModule, moduleTeardownDuration
          
          Next TestName
          
        Else
          Phosphorus.Log4PStatic.Logger.ExternalError "Skipping tests in module " & projectName & "." & vbTestModule.Name & " due to BeforeModule failure"
          For Each TestName In executableTests
            Set TestResult = New pUnit.TestResult
            TestResult.TestName = projectName & "." & vbTestModule.Name & "." & TestName
            TestResult.Status = skipped
            TestResult.ErrorMessage = "Skipped due to BeforeModule failure"
            TestResult.duration = 0
            TestResults.Add TestResult
            Phosphorus.Log4PStatic.Logger.ExternalInfo "Test skipped due to BeforeModule failure: " & TestResult.TestName & " (Annotations: " & GetTestAnnotations(vbTestModule, VBA.Conversion.CStr(TestName)) & ")"
          Next TestName
        End If
        
        'Module clean up
        ExecuteModuleTeardown vbProj, vbTestModule, moduleTeardownDuration
        QueryPerformanceCounter moduleEndCount
        
        'Calculate module level duration
        moduleDuration = GetMilliseconds(moduleStartCount, moduleEndCount)
        
        'Log module completion stats
        Phosphorus.Log4PStatic.Logger.ExternalInfo "Module " & projectName & "." & vbTestModule.Name & ": Setup Duration: " & Format(moduleSetupDuration, "0.000") & " ms, Teardown Duration: " & Format(moduleTeardownDuration, "0.000") & " ms"
        Phosphorus.Log4PStatic.Logger.ExternalInfo "Module " & projectName & "." & vbTestModule.Name & " completed (Total Duration: " & Format(moduleDuration, "0.000") & " ms)"
      
      Else
        Phosphorus.Log4PStatic.Logger.ExternalInfo "Module " & projectName & "." & vbTestModule.Name & " has no tests to run after filtering."
      End If

    End If
  Next testModuleName
  
  'Log detailed test results
  Phosphorus.Log4PStatic.Logger.ExternalInfo "Detailed Test Results:"
  For Each TestResult In TestResults
    Dim statusText As String
    Select Case TestResult.Status
      Case Passed: statusText = "Passed"
      Case FAILED: statusText = "Failed"
      Case skipped: statusText = "Skipped"
    End Select
    If TestResult.Status = Passed Then
      Phosphorus.Log4PStatic.Logger.ExternalInfo "Test: " & TestResult.TestName & IIf(Len(TestResult.parameters) > 0, " (Parameters: " & TestResult.parameters & ")", "") & ", Status: " & statusText & ", Duration: " & Format(TestResult.duration, "0.000") & " ms"
    Else
      Phosphorus.Log4PStatic.Logger.ExternalInfo "Test: " & TestResult.TestName & IIf(Len(TestResult.parameters) > 0, " (Parameters: " & TestResult.parameters & ")", "") & ", Status: " & statusText & ", Message: " & TestResult.ErrorMessage & ", Duration: " & Format(TestResult.duration, "0.000") & " ms"
    End If
  Next TestResult
    
  'Log detailed duration stats
  Phosphorus.Log4PStatic.Logger.ExternalInfo "Total Test Method Duration (excluding setup/teardown): " & Format(TotalTestMethodDuration, "0.000") & " ms"
  Phosphorus.Log4PStatic.Logger.ExternalInfo "Total Setup Duration: " & Format(TotalSetupDuration, "0.000") & " ms"
  Phosphorus.Log4PStatic.Logger.ExternalInfo "Total Teardown Duration: " & Format(TotalTeardownDuration, "0.000") & " ms"
  QueryPerformanceCounter runEndCount
  runDuration = GetMilliseconds(runStartCount, runEndCount)
  Phosphorus.Log4PStatic.Logger.ExternalInfo "Total Run Duration: " & Format(runDuration, "0.000") & " ms"
  
  'Log test success counts
  Dim testCount As Long
  testCount = TestResults.count
  passCount = CountPassedTests
  skipCount = CountSkippedTests
  Phosphorus.Log4PStatic.Logger.ExternalInfo "Test execution completed."
  Dim intCountOfFailedTests As Integer
  intCountOfFailedTests = testCount - passCount - skipCount
  Phosphorus.Log4PStatic.Logger.ExternalInfo "Total Tests: " & testCount & ", Passed: " & passCount & ", Failed: " & intCountOfFailedTests & ", Skipped: " & skipCount

ExitSub:
  
  'Store the log file path
  LastCreatedPUnitLogFile = Phosphorus.Log4PStatic.Logger.GetFilePath
  
  'Close the logger
  Phosphorus.Log4PStatic.CloseLogger

  'Close the Assert
  Phosphorus.AssertionsStatic.CloseAssert

  Application.StatusBar = False
  VBA.Interaction.DoEvents
  
  'Output a result - non CI/CD only?
  If testCount = 0 Then
    VBA.Interaction.MsgBox "No tests to process!"
  Else
    If intCountOfFailedTests = 0 Then
      If skipCount > 0 Then
        VBA.Interaction.MsgBox "Some tests skipped!"
      Else
        VBA.Interaction.MsgBox "All tests passed!"
      End If
    Else
      VBA.Interaction.MsgBox "Some tests failed!"
    End If
  End If
  
End Sub

Public Function CheckForAndRaiseErrors()
  
  If Not gpUnitError Is Nothing Then
  
    'Store the error details before we erase the object
    Dim lngNumber As Long
    Dim strDescription As String
    lngNumber = gpUnitError.Number
    strDescription = gpUnitError.Description
  
    'We have to erase the error object BEFORE the error is raised!
    Set gpUnitError = Nothing
  
    'Finally, raise the error
    Err.Raise Number:=lngNumber, Description:=strDescription
  
  End If
  
End Function

' Validate filter syntax before execution
Private Function ValidateFilterSyntax(filter As String, filterType As String) As Boolean

  'Allow empty filters
  If Len(filter) = 0 Then
    ValidateFilterSyntax = True
    Exit Function
  End If
  
  'Initialise the string to bel evaluated by the Excel VBA Application.Evaluate function
  Dim evalString As String
  evalString = filter
  
  'Prepare a regex patter match
  Dim regex As Object
  Set regex = CreateObject("VBScript.RegExp")
  regex.pattern = "[A-Za-z0-9_"" ]+"
  regex.Global = True
  
  'Get a list of all matching text substrings within the current filter
  Dim matches As Object
  Set matches = regex.Execute(evalString)
   
  'Loop through all matches
  Dim match As Variant
  For Each match In matches
    'Ignore any allowable Excel functional keyworks ... this list can be expanded!
    If Not (match.value = "TRUE" Or match.value = "FALSE" Or match.value = "AND" Or match.value = "OR" Or match.value = "NOT") Then
      'Replace each matching substring with TRUE so we can run a test evaluation on the filter string for valid syntax
      evalString = VBA.Strings.Replace(evalString, match.value, "TRUE", 1, -1, vbTextCompare)
    End If
  Next match
  'Convert any NOT entries to prevent incorrect failures
  evalString = VBA.Strings.Replace(evalString, "NOT", "", 1, -1, vbTextCompare)

  'Run the test evaluation
  On Error GoTo EvalError
  Dim boolEvaluated As Boolean
  boolEvaluated = Application.Evaluate(evalString)
  'The evaluation will return true if no errors
  If boolEvaluated Then
    ValidateFilterSyntax = True
  End If
  Exit Function

'Log any evaluation errors and return false
EvalError:
  Phosphorus.Log4PStatic.Logger.ExternalError "Invalid " & filterType & " filter syntax: " & filter & " (Error: " & Err.Description & ")"
  ValidateFilterSyntax = False

End Function

' Check if a module is annotated with '@TestModule' or '@Tag(TestModule)'
Private Function IsTestModule(vbComp As VBComponent) As Boolean

  Dim codeMod As CodeModule
  Dim lineNum As Long
  Dim lineText As String
  Dim tags As Variant
  Dim tag As Variant
    
  'Store the module and get the number of lines of code in the module
  Set codeMod = vbComp.CodeModule
  'lineNum = codeMod.CountOfDeclarationLines + 1
  'Start at the first line!
  lineNum = 1
  
  'Loop through all code lines in the code module
  Do While lineNum <= codeMod.CountOfLines
    
    'Get the next line of code
    lineText = Trim(codeMod.Lines(lineNum, 1))
    
    'Stop once we have reached any Option declarations
    If lineText Like "Option *" Then
      Exit Do
    End If
    
    'Stop once we have reached any standard method declaration lines
    If lineText Like "Sub *" Or lineText Like "Public Sub *" Or lineText Like "Private Sub *" Or lineText Like "Function *" Or lineText Like "Public Function *" Or lineText Like "Private Function *" Then
      Exit Do
    End If
        
    'Is the current line an annotation line?
    If lineText Like "'@*" Then
      'Parse all tags in the line
      tags = ParseAnnotationTags(lineText)
      For Each tag In tags
        If UCase(tag) = "TESTMODULE" Then
          IsTestModule = True
          Exit Function
        End If
      Next tag
    End If
    
    'Go to the next line
    lineNum = lineNum + 1
    
  Loop
    
  IsTestModule = False

End Function

' Parse annotation line to extract and flatten tags
Private Function ParseAnnotationTags(annotation As String) As Variant
  
  'Remove any leading single quote
  If VBA.Strings.Left$(annotation, 1) = "'" Then
    annotation = VBA.Strings.Mid$(annotation, 2)
  End If
   
  'Trim the line of spaces
  annotation = VBA.Strings.Trim(annotation)

  'Start a collection of tags in the current line
  Dim tags As collection
  Set tags = New collection
  Dim tagString As String

  'Is the line a straight single tag ... and has no comma separator or brackets - check this after processing complex tags
  If annotation Like "@[A-Za-z0-9_]*" And VBA.Strings.InStr(1, annotation, ",") = 0 And VBA.Strings.InStr(1, annotation, "(") = 0 And VBA.Strings.InStr(1, annotation, ")") = 0 Then
      
    'Remove the leading @
    tagString = VBA.Strings.Mid$(annotation, 2)
    'Do we have any text left?
    If VBA.Strings.Len(VBA.Strings.Trim(tagString)) > 0 Then
      'Add the tag to the collection
      tags.Add tagString
    End If
    'Return the tag as an array
    ParseAnnotationTags = CollectionToArray(tags)
    Exit Function
    
  End If
 
  'Is this a complex tag?
  If annotation Like "@Tag(*)" Or annotation Like "@TestData(*)" Then
        
    'Get the text between the brackets
    tagString = VBA.Strings.Mid$(annotation, VBA.Strings.InStr(annotation, "(") + 1, VBA.Strings.InStrRev(annotation, ")") - VBA.Strings.InStr(annotation, "(") - 1)
    tagString = VBA.Strings.Trim(tagString)
        
    Dim currentTag As String
    currentTag = ""
    
    Dim inQuotes As Boolean
    inQuotes = False
    
    Dim inBraces As Long
    inBraces = 0
    
    Dim inParens As Long
    inParens = 0
    
    Dim isKey As Boolean
    isKey = False
    
    'Loop through each character of the string
    Dim i As Long
    For i = 1 To VBA.Strings.Len(tagString)
    
      'Get the current character
      Dim c As String
      c = VBA.Strings.Mid$(tagString, i, 1)
            
      Select Case c
      
        Case """"
          If inQuotes Then
            inQuotes = False
          Else
            inQuotes = True
          End If
          currentTag = currentTag & c
                    
        Case "{"
          If Not inQuotes Then
            inBraces = inBraces + 1
            If inBraces = 1 Then
              currentTag = ""
            End If
          Else
            currentTag = currentTag & c
          End If
                    
        Case "}"
          If Not inQuotes Then
            inBraces = inBraces - 1
            If inBraces = 0 Then
              Dim nestedTags As Variant
              nestedTags = SplitNestedTags(currentTag)
              Dim nestedTag As Variant
              For Each nestedTag In nestedTags
                If VBA.Strings.Len(VBA.Strings.Trim(nestedTag)) > 0 Then
                  tags.Add Trim(nestedTag)
                End If
              Next nestedTag
              currentTag = ""
            ElseIf inBraces < 0 Then
              Phosphorus.Log4PStatic.Logger.ExternalWarning "Malformed annotation (unmatched brace): " & annotation
              ParseAnnotationTags = CollectionToArray(tags)
              Exit Function
            End If
          Else
            currentTag = currentTag & c
          End If
                    
        Case "("
          If Not inQuotes Then
            inParens = inParens + 1
          End If
          currentTag = currentTag & c
                    
        Case ")"
          If Not inQuotes Then
            inParens = inParens - 1
            If inParens < 0 Then
              Phosphorus.Log4PStatic.Logger.ExternalWarning "Malformed annotation (unmatched parenthesis): " & annotation
              ParseAnnotationTags = CollectionToArray(tags)
              Exit Function
            End If
          End If
          currentTag = currentTag & c
                    
        Case ","
          If Not inQuotes And inBraces = 0 And inParens = 0 Then
            If VBA.Strings.Len(VBA.Strings.Trim(currentTag)) > 0 Then
              If Not isKey Then
                tags.Add VBA.Strings.Trim(currentTag)
              End If
              currentTag = ""
              isKey = False
            End If
          Else
            currentTag = currentTag & c
          End If
                    
        Case "="
          If Not inQuotes And inBraces = 0 And inParens = 0 Then
            isKey = True
            currentTag = currentTag & c
          Else
            currentTag = currentTag & c
          End If
                    
        Case Else
          currentTag = currentTag & c
      
      End Select
    
    Next i
        
    If VBA.Strings.Len(VBA.Strings.Trim(currentTag)) > 0 Then
      If Not isKey Then
        tags.Add VBA.Strings.Trim(currentTag)
      End If
    End If
        
    If inQuotes Then
      Phosphorus.Log4PStatic.Logger.ExternalWarning "Malformed annotation (unclosed quote): " & annotation
    End If
    
    If inBraces > 0 Then
      Phosphorus.Log4PStatic.Logger.ExternalWarning "Malformed annotation (unclosed brace): " & annotation
    End If
    
    If inParens > 0 Then
      Phosphorus.Log4PStatic.Logger.ExternalWarning "Malformed annotation (unclosed parenthesis): " & annotation
    End If
  
  Else
      
    Phosphorus.Log4PStatic.Logger.ExternalWarning "Invalid annotation format: " & annotation
      
  End If
    
  ParseAnnotationTags = CollectionToArray(tags)

End Function

' Split nested tags within {} into individual tags
Private Function SplitNestedTags(nestedTagString As String) As Variant
    Dim tags As collection
    Dim currentTag As String
    Dim i As Long
    Dim inQuotes As Boolean
    Dim inBraces As Long
    
    Set tags = New collection
    currentTag = ""
    inQuotes = False
    inBraces = 0
    
    For i = 1 To Len(nestedTagString)
        Dim c As String
        c = Mid$(nestedTagString, i, 1)
        
        Select Case c
            Case """"
                If inQuotes Then inQuotes = False Else inQuotes = True
                currentTag = currentTag & c
            Case "{"
                If Not inQuotes Then inBraces = inBraces + 1
                currentTag = currentTag & c
            Case "}"
                If Not inQuotes Then
                    inBraces = inBraces - 1
                    If inBraces < 0 Then
                        Phosphorus.Log4PStatic.Logger.ExternalWarning "Malformed nested tag (unmatched brace): " & nestedTagString
                        SplitNestedTags = CollectionToArray(tags)
                        Exit Function
                    End If
                End If
                currentTag = currentTag & c
            Case ","
                If Not inQuotes And inBraces = 0 Then
                    If Len(Trim(currentTag)) > 0 Then
                        tags.Add Trim(currentTag)
                        currentTag = ""
                    End If
                Else
                    currentTag = currentTag & c
                End If
            Case Else
                currentTag = currentTag & c
        End Select
    Next i
    
    If Len(Trim(currentTag)) > 0 Then
        tags.Add Trim(currentTag)
    End If
    
    SplitNestedTags = CollectionToArray(tags)
End Function

' Convert a Collection to a Variant array
Private Function CollectionToArray(col As collection) As Variant

  Dim result() As Variant
  Dim i As Long
  
  'Is the collection empty?
  If col Is Nothing Or col.count = 0 Then
    'Return an empty array
    CollectionToArray = Array()
    Exit Function
  End If
  
  'Size a return array & loop through all items in the collection
  ReDim result(0 To col.count - 1)
  For i = 1 To col.count
    'Add the current item to the return array
    result(i - 1) = col.item(i)
  Next i
  'Return the non-empty array
  CollectionToArray = result

End Function

' Check if a module should be included based on moduleFilter
Private Function ShouldIncludeTestModule(vbComp As VBComponent, moduleNameFilter As String, allTestModuleNames As collection) As Boolean

  'Include all test mode
  If Len(moduleNameFilter) = 0 Then
    ShouldIncludeTestModule = True
    Exit Function
  End If
    
  Dim currentModule As String
  currentModule = vbComp.Name
  
  Dim evalString As String
  evalString = moduleNameFilter
  
  'Loop through all test modules
  Dim moduleName As Variant
  For Each moduleName In allTestModuleNames
    If UCase(moduleName) = UCase(currentModule) Then
      'Replace the current test module name with TRUE
      evalString = Replace(evalString, moduleName, "TRUE", 1, -1, vbTextCompare)
    Else
      'Replace the any other test module name with FALSE
      evalString = Replace(evalString, moduleName, "FALSE", 1, -1, vbTextCompare)
    End If
  Next moduleName
    
  'Evaluate the expression
  On Error GoTo EvalError
  ShouldIncludeTestModule = Application.Evaluate(evalString)
  Exit Function
    
EvalError:
  Phosphorus.Log4PStatic.Logger.ExternalError "Invalid module filter expression: " & moduleNameFilter & " (Evaluated as: " & evalString & ")"
  ShouldIncludeTestModule = False
    
End Function

' Check if a module has any tests that pass the annotation filter
Private Function HasTestsToRun(projectName As String, vbTestModule As VBComponent, annotationFilter As String, ByRef executableTests As collection) As Boolean

  'Store a reference to the test module's code module
  Dim testModuleCode As CodeModule
  Set testModuleCode = vbTestModule.CodeModule
  
  'Initialise a collection of test method annotations for the whole module
  Dim currentModuleMethodAnnotations As collection
  Set currentModuleMethodAnnotations = New collection
  Dim moduleLevelAnnotation As Variant
  
  'Initialise a collection of test method annotations for each test method
  Dim currentTestMethodAnnotations As collection
'  Set currentTestMethodAnnotations = New collection
    
  Dim currentLineNum As Long
  currentLineNum = 1
  
  Dim intCountOfDeclarationLines As Integer
  intCountOfDeclarationLines = testModuleCode.CountOfDeclarationLines
  
  'Loop through each line of code in the test module
  Dim currentLineText As String
  
  While currentLineNum <= testModuleCode.CountOfLines
    
    'Get the trimmed current line text
    currentLineText = VBA.Strings.Trim(testModuleCode.Lines(currentLineNum, 1))
        
    'Store the line if it begins with @
    If currentLineText Like "'@*" Then
      If (VBA.Strings.InStr(1, VBA.Strings.UCase(currentLineText), "'@FOLDER") <> 1) And _
         (VBA.Strings.InStr(1, VBA.Strings.UCase(currentLineText), "'@TESTMODULE") <> 1) And _
         (VBA.Strings.InStr(1, VBA.Strings.UCase(currentLineText), "'@TESTMETHOD") <> 1) And _
         (VBA.Strings.InStr(1, VBA.Strings.UCase(currentLineText), "'@TESTDATA") <> 1) _
      Then
        'Is this a module or test level tag?
        If currentLineNum <= intCountOfDeclarationLines Then
          currentModuleMethodAnnotations.Add currentLineText
        Else
          currentTestMethodAnnotations.Add currentLineText
        End If
      End If
    End If
        
    'Is this the last line of the declations section?
    If currentLineNum = intCountOfDeclarationLines Then
    
      'Add the module level annotations for the first test
      Set currentTestMethodAnnotations = New collection
      For Each moduleLevelAnnotation In currentModuleMethodAnnotations
        currentTestMethodAnnotations.Add moduleLevelAnnotation
      Next
    
    End If
    
    'Is the current line an indicator that the next line is a test method?
    If currentLineText Like "'@TestMethod*" Then
      
      Dim procedureName As String
      Dim procedureKind As vbext_ProcKind
      
      'The @TestMethod tag must immediately preceed the test method
      procedureName = GetProcedureName(testModuleCode, currentLineNum + 1, procedureKind)
      
      'Is the annotation filter just this test name
      If procedureName = annotationFilter Then
      
        'Add the test to the collection of executable tests
        executableTests.Add procedureName
        
      'Does the current test method annotations match the annotation filter
      ElseIf ShouldIncludeTest(currentTestMethodAnnotations, annotationFilter) Then
        
        'Add the test to the collection of executable tests
        If Len(procedureName) > 0 Then
          executableTests.Add procedureName
        End If
      
      Else
        
        If Len(procedureName) > 0 Then
          
          Dim strJoinedAnnotations As String
          strJoinedAnnotations = VBA.Strings.Join(CollectionToArray(currentTestMethodAnnotations), ", ")
          If strJoinedAnnotations <> "" Then
            strJoinedAnnotations = " " & strJoinedAnnotations
          End If
          Phosphorus.Log4PStatic.Logger.ExternalInfo "Test skipped due to annotations at line " & currentLineNum & ": " & projectName & "." & testModuleCode.Name & "." & procedureName & " (Annotations:" & strJoinedAnnotations & " or Procedure Name: " & procedureName & ")"
        
        End If
      
      End If
      
      'Start a new collections of annotations that will apply to the next test method based on the module annotations
      Set currentTestMethodAnnotations = New collection
      For Each moduleLevelAnnotation In currentModuleMethodAnnotations
        currentTestMethodAnnotations.Add moduleLevelAnnotation
      Next
        
    End If
        
    'Move to the next line
    currentLineNum = currentLineNum + 1
    
  Wend
    
  HasTestsToRun = executableTests.count > 0

End Function

' Determine if a test should be included based on its annotations
Private Function ShouldIncludeTest(testMethodAnnotations As collection, testAnnotationsFilter As String) As Boolean
  
  'Include all tests if the test annotations filter is empty
  If Len(testAnnotationsFilter) = 0 Then
    ShouldIncludeTest = True
    Exit Function
  End If
  
  'Prepare a list of all tags, including expanded ones
  Dim allTags As collection
  Set allTags = New collection
  Dim annotation As Variant
  
  'Loop through all annotations for the current test method
  For Each annotation In testMethodAnnotations
  
    'Build a list of all tags for the current test method
    Dim tags As Variant
    tags = ParseAnnotationTags(VBA.Conversion.CStr(annotation))
    Dim tag As Variant
    For Each tag In tags
      On Error Resume Next
      allTags.Add tag, UCase(tag)
      On Error GoTo 0
    Next tag
  Next annotation

  'Initialise the evaluation string
  Dim evalString As String
  evalString = testAnnotationsFilter
  
  'Replace all matching tags in test annotations filter with TRUE
  Dim uniqueTag As Variant
  For Each uniqueTag In allTags
    evalString = VBA.Strings.Replace(evalString, uniqueTag, "TRUE", 1, -1, vbTextCompare)
  Next uniqueTag
    
  'Get a list of remaining text substrings in the evaluation string
  Dim regex As Object
  Set regex = CreateObject("VBScript.RegExp")
  regex.pattern = "[A-Za-z0-9_"" ]+"
  regex.Global = True
  Dim matches As Object
  Set matches = regex.Execute(evalString)
  Dim match As Variant
  For Each match In matches
    'Do not replace any Excel keywords
    If Not (match.value = "TRUE" Or match.value = "FALSE" Or match.value = "AND" Or match.value = "OR" Or match.value = "NOT") Then
     'Replace any other test substrings with FALSE
      evalString = Replace(evalString, match.value, "FALSE", 1, -1, vbTextCompare)
    End If
  Next match

  'Execute the evaluation
  On Error GoTo EvalError
  ShouldIncludeTest = Application.Evaluate(evalString)
  Exit Function

'Trap any errors in the evaulation
EvalError:
  Phosphorus.Log4PStatic.Logger.ExternalError "Invalid annotation filter expression: " & testAnnotationsFilter & " (Evaluated as: " & evalString & ")"
  ShouldIncludeTest = False

End Function

Private Function GetProcedureName(codeMod As CodeModule, startLine As Long, ByRef procedureKind As vbext_ProcKind) As String
    
  'Store the starting line
  Dim currentLineNumber As Long
  currentLineNumber = startLine
  
  'Loop through all code lines until we find the procedure definition line
  Do While currentLineNumber <= codeMod.CountOfLines
    'Get the current line of text
    Dim currentLineText As String
    currentLineText = VBA.Strings.Trim(codeMod.Lines(currentLineNumber, 1))
    'Test Methods must be subroutines!
    If currentLineText Like "Sub *" Or currentLineText Like "Public Sub *" Or currentLineText Like "Private Sub *" Then
      Dim procedureName As String
      procedureName = codeMod.ProcOfLine(currentLineNumber, vbext_pk_Proc)
      procedureKind = vbext_pk_Proc
      GetProcedureName = procedureName
      Exit Function
    End If
    'Try the next line
    currentLineNumber = currentLineNumber + 1
  Loop
  GetProcedureName = ""

End Function

Private Function ProcedureExists(codeMod As CodeModule, procName As String) As Boolean
  Dim lineNum As Long
  On Error Resume Next
  lineNum = codeMod.ProcBodyLine(procName, vbext_pk_Proc)
  On Error GoTo 0
  ProcedureExists = (lineNum > 0)
End Function

Private Function ExecuteModuleSetup(projectName As String, vbProj As VBProject, vbComp As VBComponent, ByRef setupDuration As Double) As Boolean

  Dim startCount As Currency
  Dim endCount As Currency
  Dim duration As Double
    
  'Check to see if there is a BeforeModule method to be run
  If ProcedureExists(vbComp.CodeModule, "BeforeModule") Then
  
    'Start a counter
    QueryPerformanceCounter startCount
    
    'Jump to SetupError if the BeforeMethod errors
    'See: https://stackoverflow.com/questions/42124252/application-run-with-error-trapping
    ClearAllErrors
    On Error GoTo SetupError
    
    'Run the BeforeModule method
    Dim strProcedureName As String
    strProcedureName = vbProj.Name & "." & vbComp.Name & ".BeforeModule"
    Phosphorus.Log4PStatic.Logger.ExternalInfo "Running " & strProcedureName
    Application.Run strProcedureName
    
    'Check for errors and raise if necessary
    CheckForAndRaiseErrors
    
    'Module setup executed successfully if we reach here
    
    'Stop the counter
    QueryPerformanceCounter endCount
    
    'Store the execution times
    duration = GetMilliseconds(startCount, endCount)
    setupDuration = setupDuration + duration
    TotalSetupDuration = TotalSetupDuration + duration
    
    'Method executed successfully
    ExecuteModuleSetup = True
    Exit Function
        
SetupError:
   'Log the failure
    If Err.Number <> 0 Then
      Phosphorus.Log4PStatic.Logger.ExternalError "BeforeModule failed in module " & projectName & "." & vbComp.Name & ": Error #" & Err.Number & ": " & Err.Description
    Else
      Phosphorus.Log4PStatic.Logger.ExternalError "BeforeModule failed in module " & projectName & "." & vbComp.Name & ": Error #" & gpUnitError.Number & ": " & gpUnitError.Description
    End If
    
    'Stop the counter
    QueryPerformanceCounter endCount
    
    'Store the execution times
    duration = GetMilliseconds(startCount, endCount)
    setupDuration = setupDuration + duration
    TotalSetupDuration = TotalSetupDuration + duration
    
    'Method didn't execute
    ExecuteModuleSetup = False
  
  Else
    'Nothing to run, so ok!
    ExecuteModuleSetup = True
  End If

End Function

Private Sub ExecuteModuleTeardown(vbProj As VBProject, vbComp As VBComponent, ByRef teardownDuration As Double)

  Dim startCount As Currency
  Dim endCount As Currency
  Dim duration As Double
    
  'Check to see if there is a ModuleTeardown method to be run
  If ProcedureExists(vbComp.CodeModule, "AfterModule") Then
    
    'Start a counter
    QueryPerformanceCounter startCount
    
    'Jump to SetupError if the ModuleTeardown errors
    'See: https://stackoverflow.com/questions/42124252/application-run-with-error-trapping
    ClearAllErrors
    On Error GoTo TeardownError
    
    'Run the AfterModule method
    Dim strProcedureName As String
    strProcedureName = vbProj.Name & "." & vbComp.Name & ".AfterModule"
    Phosphorus.Log4PStatic.Logger.ExternalInfo "Running " & strProcedureName
    Application.Run strProcedureName
    
    'Check for errors and raise if necessary
    CheckForAndRaiseErrors
    
    'Module teardown executed successfully if we reach here
    
    'Stop the counter
    QueryPerformanceCounter endCount
    'Calulate duration
    duration = GetMilliseconds(startCount, endCount)
    teardownDuration = teardownDuration + duration
    TotalTeardownDuration = TotalTeardownDuration + duration
    
    'Method executed successfully
    Exit Sub
        
TeardownError:
    
    'Log that Module teardown didn't execute successfully - doesn't cause test failure or skip
    If Err.Number <> 0 Then
      Phosphorus.Log4PStatic.Logger.ExternalError "AfterModule failed in module " & vbProj.Name & "." & vbComp.Name & ": Error #" & Err.Number & ": " & Err.Description
    Else
      Phosphorus.Log4PStatic.Logger.ExternalError "AfterModule failed in module " & vbProj.Name & "." & vbComp.Name & ": Error #" & gpUnitError.Number & ": " & gpUnitError.Description
    End If
    
    'End counter
    QueryPerformanceCounter endCount
    
    'Calulate duration
    duration = GetMilliseconds(startCount, endCount)
    teardownDuration = teardownDuration + duration
    TotalTeardownDuration = TotalTeardownDuration + duration
  
  End If

End Sub

Private Function ExecuteTestSetup(TestName As String, vbProj As VBProject, vbComp As VBComponent, ByRef setupDuration As Double) As Boolean
    
  Dim startCount As Currency
  Dim endCount As Currency
  Dim duration As Double
    
  'Check to see if there is a TestSetup method to be run
  If ProcedureExists(vbComp.CodeModule, "BeforeTest") Then
  
    'Start a counter
    QueryPerformanceCounter startCount
    
    'Jump to SetupError if the TestSetup errors
    'See: https://stackoverflow.com/questions/42124252/application-run-with-error-trapping
    ClearAllErrors
    On Error GoTo SetupError
    
    'Run the BeforeTest method
    Dim strProcedureName As String
    strProcedureName = vbProj.Name & "." & vbComp.Name & ".BeforeTest"
    Phosphorus.Log4PStatic.Logger.ExternalInfo "Running " & strProcedureName & "(" & TestName & ")"
    Application.Run strProcedureName
    
    'Check for errors and raise if necessary
    CheckForAndRaiseErrors
   
    'Module teardown executed successfully if we reach here
       
    'Stop the counter
    QueryPerformanceCounter endCount
    
    'Store the execution times
    duration = GetMilliseconds(startCount, endCount)
    setupDuration = setupDuration + duration
    TotalSetupDuration = TotalSetupDuration + duration
    
    'TestSetup executed successfully
    ExecuteTestSetup = True
    Exit Function
       
SetupError:
    'Log the failure
    If Err.Number <> 0 Then
      Phosphorus.Log4PStatic.Logger.ExternalError "BeforeTest(" & TestName & ") failed in module " & vbProj.Name & "." & vbComp.Name & ": Error #" & Err.Number & ": " & Err.Description
    Else
      Phosphorus.Log4PStatic.Logger.ExternalError "BeforeTest(" & TestName & ") failed in module " & vbProj.Name & "." & vbComp.Name & ": Error #" & gpUnitError.Number & ": " & gpUnitError.Description
    End If
    
    'Stop the counter
    QueryPerformanceCounter endCount
    
    'Store the execution times
    duration = GetMilliseconds(startCount, endCount)
    setupDuration = setupDuration + duration
    TotalSetupDuration = TotalSetupDuration + duration
    
    'Method didn't execute
    ExecuteTestSetup = False
  Else
    'Nothing to run, so ok!
    ExecuteTestSetup = True
  End If

End Function

Private Sub ExecuteTestTeardown(TestName As String, vbProj As VBProject, vbComp As VBComponent, ByRef teardownDuration As Double)
    
  Dim startCount As Currency
  Dim endCount As Currency
  Dim duration As Double
    
  'Check to see if there is a TestTeardown method to be run
  If ProcedureExists(vbComp.CodeModule, "AfterTest") Then
        
    'Start a counter
    QueryPerformanceCounter startCount
    
    'Jump to TeardownError if the TestTeardown errors
    'See: https://stackoverflow.com/questions/42124252/application-run-with-error-trapping
    ClearAllErrors
    On Error GoTo TeardownError
    
    'Run the AfterTest method
    Dim strProcedureName As String
    strProcedureName = vbProj.Name & "." & vbComp.Name & ".AfterTest"
    Phosphorus.Log4PStatic.Logger.ExternalInfo "Running " & strProcedureName & "(" & TestName & ")"
    Application.Run strProcedureName
    
    'Check for errors and raise if necessary
    CheckForAndRaiseErrors
   
    'Test teardown executed successfully if we reach here
       
    'Stop the counter
    QueryPerformanceCounter endCount
    
    'Calulate duration
    duration = GetMilliseconds(startCount, endCount)
    teardownDuration = teardownDuration + duration
    TotalTeardownDuration = TotalTeardownDuration + duration
    
    'TestTeardown executed successfully
    Exit Sub
        
TeardownError:
    'Log the failure
    If Err.Number <> 0 Then
      Phosphorus.Log4PStatic.Logger.ExternalError "AfterTest(" & TestName & ") failed in module " & vbComp.Name & ": Error #" & Err.Number & ": " & Err.Description
    Else
      Phosphorus.Log4PStatic.Logger.ExternalError "AfterTest(" & TestName & ") failed in module " & vbComp.Name & ": Error #" & gpUnitError.Number & ": " & gpUnitError.Description
    End If
    
    'End counter
    QueryPerformanceCounter endCount
    
    'Store the execution times
    duration = GetMilliseconds(startCount, endCount)
    teardownDuration = teardownDuration + duration
    TotalTeardownDuration = TotalTeardownDuration + duration
  
  End If

End Sub

' Execute a test, supporting up to 30 parameters
Private Function ExecuteTest(vbProj As VBProject, vbComp As VBComponent, TestName As String, Optional parameters As Variant) As pUnit.TestResult
    
  Dim TestResult As New pUnit.TestResult
  'Start a new assertions class
  Phosphorus.AssertionsStatic.CloseAssert
  Phosphorus.AssertionsStatic.GetAssert
  
  ' Build test name with parameters
  TestResult.TestName = vbProj.Name & "." & vbComp.Name & "." & TestName

  'Print to VB Editor Immediate screen for when running tests manually
'Removed - this interferes the the Excel base unit tests
'  Application.StatusBar = "Processing " & TestResult.TestName
  VBA.Interaction.DoEvents
  
  ' Build test parameters string
  If Not IsMissing(parameters) Then
    Dim param As Variant
    Dim paramStrings As collection
    Set paramStrings = New collection
    For Each param In parameters
      paramStrings.Add StringifyParameter(param)
    Next param
    Dim paramString As String
    paramString = VBA.Strings.Join(CollectionToArray(paramStrings), ", ")
    TestResult.parameters = paramString
    TestResult.TestName = TestResult.TestName & "(" & paramString & ")"
  End If
    
  'Start a counter
  Dim startCount As Currency
  QueryPerformanceCounter startCount
    
  'Jump to TestError if the Test errors
  'See: https://stackoverflow.com/questions/42124252/application-run-with-error-trapping
  On Error GoTo TestError
  ClearAllErrors
  
  Dim strProcedureName As String
  strProcedureName = vbProj.Name & "." & vbComp.Name & "." & TestName
  
  ' Are there any parameters?
  If IsMissing(parameters) Then
    
    'Run test with no parameters
    Phosphorus.Log4PStatic.Logger.ExternalInfo "Running " & strProcedureName
    Application.Run strProcedureName
  
  Else
    
    'Run test with parameters
    
    'Check for too many parameter passed
    If UBound(parameters) + 1 > 30 Then
      Phosphorus.Log4PStatic.Logger.ExternalError "Too many parameters for test: " & TestResult.TestName & " (Max supported: 30)"
      TestResult.Status = skipped
      TestResult.ErrorMessage = "Too many parameters (Max supported: 30)"
      GoTo TestCleanup
    End If
   
    'Call test method with the right number of parameters
    strProcedureName = vbProj.Name & "." & vbComp.Name & "." & TestName
    strProcedureName = strProcedureName & "("
    Dim i As Integer
    For i = 0 To UBound(parameters)
      If i > 0 Then
        strProcedureName = strProcedureName & ", "
      End If
      strProcedureName = strProcedureName & parameters(i)
    Next i
    strProcedureName = strProcedureName & ")"
    Phosphorus.Log4PStatic.Logger.ExternalInfo "Running " & strProcedureName
    
    Select Case UBound(parameters) + 1
      Case 1
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0)
      Case 2
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1)
      Case 3
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2)
      Case 4
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3)
      Case 5
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4)
      Case 6
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5)
      Case 7
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6)
      Case 8
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7)
      Case 9
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8)
      Case 10
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9)
      Case 11
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10)
      Case 12
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11)
      Case 13
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12)
      Case 14
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13)
      Case 15
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14)
      Case 16
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15)
      Case 17
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16)
      Case 18
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), _
          parameters(17)
      Case 19
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), _
          parameters(17), parameters(18)
      Case 20
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), _
          parameters(17), parameters(18), parameters(19)
      Case 21
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), _
          parameters(17), parameters(18), parameters(19), parameters(20)
      Case 22
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), _
          parameters(17), parameters(18), parameters(19), parameters(20), parameters(21)
      Case 23
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), _
          parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22)
      Case 25
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), _
          parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22), parameters(23), _
          parameters(24)
      Case 26
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), _
          parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22), parameters(23), _
          parameters(24), parameters(25)
      Case 27
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), _
          parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22), parameters(23), _
          parameters(24), parameters(25), parameters(26)
      Case 28
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), _
          parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22), parameters(23), _
          parameters(24), parameters(25), parameters(26), parameters(27)
      Case 29
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), _
          parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22), parameters(23), _
          parameters(24), parameters(25), parameters(26), parameters(27), parameters(28)
      Case 30
        Application.Run vbProj.Name & "." & vbComp.Name & "." & TestName, parameters(0), parameters(1), parameters(2), _
          parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), _
          parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), _
          parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22), parameters(23), _
          parameters(24), parameters(25), parameters(26), parameters(27), parameters(28), parameters(29)
    End Select
  End If
  
  'Check for errors and raise if necessary
  CheckForAndRaiseErrors

  'Test executed successfully if we reach here, but check for critical assertion failures
  If Phosphorus.AssertionsStatic.pAssert.HasCriticalFailure Then
    TestResult.Status = FAILED
'    TestResult.ErrorMessage = Phosphorus.AssertionsStatic.pAssert.FailedTests & "Failed Assertions, including at least 1 critical"
'    TestResults.Add TestResult
    Phosphorus.Log4PStatic.Logger.ExternalInfo "Test failed due to assertion failures, including at least 1 critical: " & TestResult.TestName
    TestResult.ErrorMessage = "Test failed due to assertion failures, including at least 1 critical: " & TestResult.TestName
  Else
  
    If Phosphorus.AssertionsStatic.pAssert.FailedTests > 0 Then
      Phosphorus.Log4PStatic.Logger.ExternalInfo "Test passed but there were assertion failures: " & TestResult.TestName
    End If
    
    'Test executed successfully if we reach here
    TestResult.Status = Passed
  
  End If
  
TestCleanup:
  'Stop the counter
  Dim endCount As Currency
  QueryPerformanceCounter endCount

 'Calulate duration
  Dim duration As Double
  duration = GetMilliseconds(startCount, endCount)
  TestResult.duration = duration
  TotalTestMethodDuration = TotalTestMethodDuration + duration
  
  'TestTeardown executed successfully
  Set ExecuteTest = TestResult
  Exit Function
  
TestError:
  'Log test failure
  TestResult.Status = FAILED
  If Err.Number <> 0 Then
    TestResult.ErrorMessage = "Error #" & Err.Number & ": " & Err.Description
  Else
    TestResult.ErrorMessage = "Error #" & gpUnitError.Number & ": " & gpUnitError.Description
  End If
  Phosphorus.Log4PStatic.Logger.ExternalError "Test failed: " & TestResult.TestName & " - " & TestResult.ErrorMessage
  Set ExecuteTest = TestResult

End Function

' Get test data for a specific test method
Private Function GetTestData(vbComp As VBComponent, TestName As String) As collection

  Dim annotation As Variant
    
  Dim testData As collection
  Set testData = New collection
  
  Dim codeMod As CodeModule
  Set codeMod = vbComp.CodeModule
  
  Dim annotations As collection
  Set annotations = New collection
  
  'Start at line 1
  Dim currentLineNumber As Long
  currentLineNumber = 1
    
  'Loop through each line
  While currentLineNumber <= codeMod.CountOfLines
  
    Dim currentLineText As String
    currentLineText = VBA.Strings.Trim(codeMod.Lines(currentLineNumber, 1))
    
    'Store any annotations
    If currentLineText Like "'@*" Then
      annotations.Add currentLineText
    End If
    
    'Have we reached a test method definition line?
    If currentLineText Like "Sub *" Or currentLineText Like "Public Sub *" Or currentLineText Like "Private Sub *" Then
      
      'Have we reached the target test method?
      If codeMod.ProcOfLine(currentLineNumber, vbext_pk_Proc) = TestName Then
        
        'Process all the annotations preceding this test
        For Each annotation In annotations
          
          'Is this a line of test data?
          If annotation Like "'@TestData*" Then
            
            'Skip errors as we parse the data line
            Dim dataSets As Variant
            On Error Resume Next
            'Parse the test data line into an array
            dataSets = ParseTestData(VBA.Conversion.CStr(annotation), vbComp.Name & "." & TestName)
            On Error GoTo 0
            
            'Add the new data sets to a collection
            If Not IsEmpty(dataSets) Then
              Dim dataSet As Variant
              For Each dataSet In dataSets
                testData.Add dataSet
              Next dataSet
            End If
        
          End If
        
        Next annotation
        
        'We've found the Test Data for the target test method, so exit function here!
        Set GetTestData = testData
        Exit Function
        
      End If
      
      'Start a new collection of annotations after each test method
      Set annotations = New collection
    
    End If
        
    currentLineNumber = currentLineNumber + 1
  
  Wend
        
End Function

' Parse test data from @TestData annotation
Private Function ParseTestData(annotation As String, TestName As String) As Variant
    
  ' Use StringBuilder for optimization
  Dim dataString As StringBuilder
  Set dataString = New StringBuilder
  
  Dim currentData As StringBuilder
  Set currentData = New StringBuilder
    
  Dim isValid As Boolean
  isValid = True
    
  'Trim off any leading ' character
  If VBA.Strings.Left$(annotation, 1) = "'" Then
    annotation = VBA.Strings.Mid$(annotation, 2)
  End If
  
  'Trim off any leading spaces
  annotation = VBA.Strings.Trim(annotation)
        
  'Get the position of the parameter parentheses
  Dim startPos As Long
  Dim endPos As Long
  startPos = InStr(annotation, "(")
  endPos = InStrRev(annotation, ")")

  If Not annotation Like "@TestData?*" Then
    Phosphorus.Log4PStatic.Logger.ExternalWarning "Invalid @TestData format for test " & TestName & ": " & annotation & " (Expected '@TestData(...)')"
    ParseTestData = Empty
    Exit Function
  End If

  If startPos = 0 Or endPos = 0 Or endPos <= startPos Then
    Phosphorus.Log4PStatic.Logger.ExternalWarning "Malformed @TestData for test " & TestName & ": missing or unbalanced parentheses in " & annotation
    ParseTestData = Empty
    Exit Function
  End If
    
  dataString.Append VBA.Strings.Mid$(annotation, startPos + 1, endPos - startPos - 1)
  If dataString.Length = 0 Then
    Phosphorus.Log4PStatic.Logger.ExternalWarning "Empty @TestData for test " & TestName & ": " & annotation
    ParseTestData = Empty
    Exit Function
  End If
      
  Dim i As Long
  Dim c As String
  Dim inQuotes As Boolean
  Dim inBraces As Long
  Dim inParens As Long
  Dim dataSet As collection
  
  Dim testData As collection
  Set testData = New collection

  For i = 1 To dataString.Length
  
    c = dataString.ToString(i, 1)
    Select Case c
            
      Case """"
        If inQuotes Then
          inQuotes = False
        Else
          inQuotes = True
        End If
        currentData.Append c
      Case "{"
        If Not inQuotes Then
          inBraces = inBraces + 1
          If inBraces = 1 Then
            Set currentData = New StringBuilder
          End If
        Else
          currentData.Append c
        End If
      Case "}"
        If Not inQuotes Then
          inBraces = inBraces - 1
          'Have we finished a first level nested set
          If inBraces = 1 Then
            Set dataSet = New collection
            Dim params As Variant
            On Error Resume Next
            params = ParseComplexParameters(currentData.ToString, TestName)
            On Error GoTo 0
            If Not IsEmpty(params) Then
              Dim param As Variant
              For Each param In params
                dataSet.Add param
              Next param
              If dataSet.count > 0 Then
                testData.Add CollectionToArray(dataSet)
              Else
                Phosphorus.Log4PStatic.Logger.ExternalWarning "Empty data set skipped for test " & TestName & ": {" & currentData.ToString & "}"
              End If
            Else
              Phosphorus.Log4PStatic.Logger.ExternalWarning "Invalid data set skipped for test " & TestName & ": {" & currentData.ToString & "}"
            End If
            Set currentData = New StringBuilder
          ElseIf inBraces < 0 Then
            Phosphorus.Log4PStatic.Logger.ExternalWarning "Unmatched closing brace in @TestData for test " & TestName & ": " & annotation
            isValid = False
            Exit For
          End If
        Else
          currentData.Append c
        End If
      Case "("
        If Not inQuotes Then
          inParens = inParens + 1
          currentData.Append c
        Else
          currentData.Append c
        End If
      Case ")"
        If Not inQuotes Then
          inParens = inParens - 1
          If inParens < 0 Then
            Phosphorus.Log4PStatic.Logger.ExternalWarning "Unmatched closing parenthesis in @TestData for test " & TestName & ": " & annotation
            isValid = False
            Exit For
          End If
          currentData.Append c
        Else
          currentData.Append c
        End If
      Case ","
        If Not inQuotes And inBraces = 1 And inParens = 0 Then
          Set dataSet = New collection
          Dim params2 As Variant
          'Parse the test data only if we have data - the comma's between sets will otherwise be interpreted as empty data sets
          If VBA.Strings.Len(currentData.ToString) <> 0 Then
            On Error Resume Next
            params2 = ParseComplexParameters(currentData.ToString, TestName)
            On Error GoTo 0
            If Not IsEmpty(params2) Then
              Dim param2 As Variant
              For Each param2 In params2
                dataSet.Add param2
              Next param2
              If dataSet.count > 0 Then
                testData.Add CollectionToArray(dataSet)
              Else
                Phosphorus.Log4PStatic.Logger.ExternalWarning "Empty data set skipped for test " & TestName & ": {" & currentData.ToString & "}"
              End If
            Else
              Phosphorus.Log4PStatic.Logger.ExternalWarning "Invalid data set skipped for test " & TestName & ": {" & currentData.ToString & "}"
            End If
          End If
          Set currentData = New StringBuilder
        Else
          currentData.Append c
        End If
      Case Else
        currentData.Append c
    End Select
    
  Next i
    
  If inQuotes Then
    Phosphorus.Log4PStatic.Logger.ExternalWarning "Unclosed quote in @TestData for test " & TestName & ": " & annotation
    isValid = False
  End If
    
  If inBraces > 0 Then
    Phosphorus.Log4PStatic.Logger.ExternalWarning "Unclosed brace in @TestData for test " & TestName & ": " & annotation
    isValid = False
  End If
    
  If inParens > 0 Then
    Phosphorus.Log4PStatic.Logger.ExternalWarning "Unclosed parenthesis in @TestData for test " & TestName & ": " & annotation
    isValid = False
  End If
    
  If currentData.Length > 0 And inBraces = 1 And isValid Then
    Set dataSet = New collection
    Dim finalParams As Variant
    On Error Resume Next
    finalParams = ParseComplexParameters(currentData.ToString, TestName)
    On Error GoTo 0
    If Not IsEmpty(finalParams) Then
      Dim finalParam As Variant
      For Each finalParam In finalParams
        dataSet.Add finalParam
      Next finalParam
      If dataSet.count > 0 Then
        testData.Add CollectionToArray(dataSet)
      Else
        Phosphorus.Log4PStatic.Logger.ExternalWarning "Empty final data set skipped for test " & TestName & ": {" & currentData.ToString & "}"
      End If
    Else
      Phosphorus.Log4PStatic.Logger.ExternalWarning "Invalid final data set skipped for test " & TestName & ": {" & currentData.ToString & "}"
    End If
  End If
    
  If Not isValid Or testData.count = 0 Then
    Phosphorus.Log4PStatic.Logger.ExternalWarning "No valid test data parsed for test " & TestName & ": " & annotation
    ParseTestData = Empty
    Exit Function
  End If
    
  ParseTestData = CollectionToArray(testData)

End Function

' Get parameter count and types for a procedure
Private Function GetProcedureSignature(codeMod As CodeModule, procedureName As String) As Variant
        
  'Initialse the dictionary of test signatures if necessary
  If MethodSignatures Is Nothing Then
    Set MethodSignatures = CreateObject("Scripting.Dictionary")
  End If
    
  'If the method signature already exists in the collection just return it
  Dim key As String
  key = codeMod.Parent.Name & "." & procedureName
  If MethodSignatures.exists(key) Then
    GetProcedureSignature = MethodSignatures(key)
    Exit Function
  End If

  'Get the line number of the desired procedure
  Dim currentLineNumber As Long
  currentLineNumber = codeMod.ProcBodyLine(procedureName, vbext_pk_Proc)
  
  'If the procedure isn't found just return an empty key & value
  If currentLineNumber = 0 Then
    GetProcedureSignature = Array(0, Array())
    Exit Function
  End If
  
  'Get the defiinition line of the desired procedure
  Dim currentLineText As String
  currentLineText = VBA.Strings.Trim(codeMod.Lines(currentLineNumber, 1))
  'Ignore any test methods not defined as subroutines
  If Not (currentLineText Like "Sub *" Or currentLineText Like "Public Sub *" Or currentLineText Like "Private Sub *") Then
    '... just return an empty key & value
    GetProcedureSignature = Array(0, Array())
    Exit Function
  End If
    
  'Find the start and end of the parameters
  Dim startPos As Long
  Dim endPos As Long
  startPos = VBA.Strings.InStr(currentLineText, "(")
  endPos = VBA.Strings.InStrRev(currentLineText, ")")
  
  ''Ignore ill-defined parameters
  If startPos = 0 Or endPos = 0 Or endPos <= startPos Then
    '... just return an empty key & value
    GetProcedureSignature = Array(0, Array())
    Exit Function
  End If
    
  'Get the parameters between the brackets
  Dim parametersString As String
  parametersString = VBA.Strings.Mid$(currentLineText, startPos + 1, endPos - startPos - 1)
  parametersString = VBA.Strings.Trim(parametersString)
  
  'If there are no parameters ...
  If VBA.Strings.Len(parametersString) = 0 Then
    '... just return an empty key & value
    GetProcedureSignature = Array(0, Array())
    Exit Function
  End If
    
  'Prepare a collecion of parameters
  Dim parameters As collection
  Set parameters = New collection
  
  'Current parameter
  Dim currentParameterString As String
  currentParameterString = ""
  
  'Parameter count
  Dim parameterCount As Long
  parameterCount = 0
    
  Dim inQuotes As Boolean
  inQuotes = False
  
  'Loop through each character of the parameter string
  Dim characterPosition As Long
  For characterPosition = 1 To Len(parametersString)
    Dim character As String
    character = Mid$(parametersString, characterPosition, 1)
    'The action depends on the next character
    Select Case character
      Case """"
        'Start/end of quotes
        inQuotes = Not inQuotes
        currentParameterString = currentParameterString & character
      Case ","
        'Commas indicate new parameters unless we are inside quotes
        If Not inQuotes Then
          parameters.Add VBA.Strings.Trim(currentParameterString)
          'Initialise the next parameter string
          currentParameterString = ""
          parameterCount = parameterCount + 1
        Else
          'Otherwise add the character to the currentParameter
          currentParameterString = currentParameterString & character
        End If
      Case Else
        'In all other cases add the character to the currentParameter
        currentParameterString = currentParameterString & character
    End Select
  Next characterPosition
  
  'Add the last parameter after we have reached the end of the parameters string
  If VBA.Strings.Len(VBA.Strings.Trim(currentParameterString)) > 0 Then
    parameters.Add VBA.Strings.Trim(currentParameterString)
    parameterCount = parameterCount + 1
  End If
    
  'Prepare an array of the type of each parameter
  Dim parameterTypes() As String
  ReDim parameterTypes(0 To parameterCount - 1)
  
  'Loop through the array of parameters
  Dim parameterCounter As Long
  For parameterCounter = 1 To parameters.count
    Dim parameter As String
    parameter = parameters(parameterCounter)
    Dim asPos As Long
    asPos = InStrRev(parameter, " As ")
    If asPos > 0 Then
      Dim typeName As String
      typeName = VBA.Strings.Trim(VBA.Strings.Mid$(parameter, asPos + 4))
      If typeName Like "*()" Then
        typeName = "Array"
      End If
      parameterTypes(parameterCounter - 1) = typeName
    Else
      parameterTypes(parameterCounter - 1) = "Variant"
    End If
  Next parameterCounter
  
  'Prepare an array to return as the method signature
  Dim signature As Variant
  signature = VBA.Array(parameterCount, parameterTypes)
  
  'Add the signature to a collection to save retireving it again
  MethodSignatures.Add key, signature
  
  'Return the current method signature
  GetProcedureSignature = signature
    
End Function

' Helper function to get milliseconds from performance counter
Private Function GetMilliseconds(startCount As Currency, endCount As Currency) As Double
  Dim freq As Currency
  If QueryPerformanceFrequency(freq) Then
    If freq <> 0 Then
      GetMilliseconds = ((endCount - startCount) / freq) * 1000
    Else
      GetMilliseconds = 0
    End If
  Else
    GetMilliseconds = 0
  End If
End Function

' Validate parameter types against method signature
Private Function ValidateParameterTypes(dataSet As Variant, paramTypes As Variant, TestName As String) As Boolean
  
  Dim i As Long
  Dim isValid As Boolean
  isValid = True
    
  'Check that the number of data items match the number of parameters
  If UBound(dataSet) <> UBound(paramTypes) Then
    Phosphorus.Log4PStatic.Logger.ExternalWarning "Parameter count mismatch in ValidateParameterTypes for test " & TestName
    ValidateParameterTypes = False
    Exit Function
  End If
  
  'Loop through all items in the dataset
  For i = LBound(dataSet) To UBound(dataSet)
    
    Dim param As Variant
    param = dataSet(i)
    
    Dim expectedType As String
    expectedType = paramTypes(i)
        
    Dim actualType As String
    actualType = typeName(param)
    If actualType Like "Variant()" Then
      actualType = "Array"
    End If
    
    Select Case UCase(expectedType)
      Case "VARIANT"
        ' Variant accepts any type
      Case "STRING"
        If actualType <> "String" Then
           Phosphorus.Log4PStatic.Logger.ExternalWarning "Type mismatch for parameter " & (i + 1) & " in test " & TestName & ": expected String, got " & actualType
           isValid = False
         End If
      Case "INTEGER", "LONG" 'These should accept Integer or Long but NOT Double
         'If Not (actualType = "Integer" Or actualType = "Long" Or actualType = "Double") Then
         If Not (actualType = "Integer" Or actualType = "Long") Then
           Phosphorus.Log4PStatic.Logger.ExternalWarning "Type mismatch for parameter " & (i + 1) & " in test " & TestName & ": expected Integer/Long, got " & actualType
           isValid = False
         End If
      Case "DOUBLE"
         If Not (actualType = "Double" Or actualType = "Integer" Or actualType = "Long") Then
           Phosphorus.Log4PStatic.Logger.ExternalWarning "Type mismatch for parameter " & (i + 1) & " in test " & TestName & ": expected Double, got " & actualType
           isValid = False
         End If
      Case "BOOLEAN"
         If actualType <> "Boolean" Then
           Phosphorus.Log4PStatic.Logger.ExternalWarning "Type mismatch for parameter " & (i + 1) & " in test " & TestName & ": expected Boolean, got " & actualType
           isValid = False
         End If
      Case "COLLECTION"
         If actualType <> "Collection" Then
           Phosphorus.Log4PStatic.Logger.ExternalWarning "Type mismatch for parameter " & (i + 1) & " in test " & TestName & ": expected Collection, got " & actualType
           isValid = False
         End If
      Case "DICTIONARY"
         If actualType <> "Dictionary" Then
           Phosphorus.Log4PStatic.Logger.ExternalWarning "Type mismatch for parameter " & (i + 1) & " in test " & TestName & ": expected Dictionary, got " & actualType
           isValid = False
         End If
      Case "ARRAY"
         If Not IsArray(param) Then
           Phosphorus.Log4PStatic.Logger.ExternalWarning "Type mismatch for parameter " & (i + 1) & " in test " & TestName & ": expected Array, got " & actualType
           isValid = False
         End If
      Case Else
        Phosphorus.Log4PStatic.Logger.ExternalWarning "Unsupported parameter type for parameter " & (i + 1) & " in test " & TestName & ": " & expectedType
        isValid = False
    End Select
  Next i
    
  ValidateParameterTypes = isValid
  
End Function

' Stringify complex data types for reporting
Private Function StringifyParameter(param As Variant) As String
    On Error Resume Next
    If IsArray(param) Then
        Dim arrStr As String
        arrStr = "{"
        Dim i As Long
        For i = LBound(param) To UBound(param)
            arrStr = arrStr & StringifyParameter(param(i))
            If i < UBound(param) Then arrStr = arrStr & ", "
        Next i
        arrStr = arrStr & "}"
        StringifyParameter = arrStr
    ElseIf typeName(param) = "Collection" Then
        Dim colStr As String
        colStr = "{"
        Dim item As Variant
        Dim first As Boolean
        first = True
        For Each item In param
            If Not first Then colStr = colStr & ", "
            colStr = colStr & StringifyParameter(item)
            first = False
        Next item
        colStr = colStr & "}"
        StringifyParameter = colStr
    ElseIf typeName(param) = "Dictionary" Then
        Dim dictStr As String
        dictStr = "{"
        Dim key As Variant
        first = True
        For Each key In param.Keys
            If Not first Then dictStr = dictStr & ", "
            dictStr = dictStr & key & "=" & StringifyParameter(param(key))
            first = False
        Next key
        dictStr = dictStr & "}"
        StringifyParameter = dictStr
    ElseIf IsObject(param) Then
        StringifyParameter = "[Object " & typeName(param) & "]"
    Else
        StringifyParameter = CStr(param)
    End If
    On Error GoTo 0
End Function

' Parse parameters, including complex types
Private Function ParseComplexParameters(paramString As String, TestName As String) As Variant

    Dim params As collection
    Dim currentParam As StringBuilder
    Dim i As Long
    Dim inQuotes As Boolean
    Dim inBraces As Long
    Dim isKey As Boolean
    Dim keyString As String
    Dim dict As Object
    Dim isValid As Boolean
    
  Set params = New collection
  Set currentParam = New StringBuilder
    
  inQuotes = False
  inBraces = 0
  isKey = False
  isValid = True
    
  paramString = Trim(paramString)
  If Len(paramString) = 0 Then
    Phosphorus.Log4PStatic.Logger.ExternalWarning "Empty parameter string in @TestData for test " & TestName & ": {}"
    ParseComplexParameters = Empty
    Exit Function
  End If
    
  For i = 1 To Len(paramString)
    
        Dim c As String
    c = Mid$(paramString, i, 1)
        
        Select Case c
            Case """"
                If inQuotes Then inQuotes = False Else inQuotes = True
                currentParam.Append c
            Case "{"
                If Not inQuotes Then
                    inBraces = inBraces + 1
                    currentParam.Append c
                Else
                    currentParam.Append c
                End If
            Case "}"
                If Not inQuotes Then
                    inBraces = inBraces - 1
                    If inBraces < 0 Then
                        Phosphorus.Log4PStatic.Logger.ExternalWarning "Unmatched closing brace in parameter for test " & TestName & ": " & paramString
                        isValid = False
                        Exit For
                    End If
                    currentParam.Append c
                Else
                    currentParam.Append c
                End If
            Case ","
                If Not inQuotes And inBraces = 0 Then
                    If currentParam.Length > 0 Then
                        If Not isKey Then
                            Dim paramValue As Variant
                            On Error Resume Next
                            paramValue = ParseParameterValue(currentParam.ToString, TestName)
                            On Error GoTo 0
                            If Not IsEmpty(paramValue) Then
                              params.Add paramValue
                            Else
                              Phosphorus.Log4PStatic.Logger.ExternalWarning "Invalid parameter skipped for test " & TestName & ": " & currentParam.ToString
                            End If
                        End If
                        Set currentParam = New StringBuilder
                        isKey = False
                    End If
                Else
                    currentParam.Append c
                End If
            Case "="
                If Not inQuotes And inBraces = 0 Then
                    If currentParam.Length = 0 Then
                        Phosphorus.Log4PStatic.Logger.ExternalWarning "Empty key in dictionary parameter for test " & TestName & ": " & paramString
                        isValid = False
                        Exit For
                    End If
                    isKey = True
                    keyString = Trim(currentParam.ToString)
                    Set currentParam = New StringBuilder
                Else
                    currentParam.Append c
                End If
            Case Else
                currentParam.Append c
        End Select
    Next i
    
    If inQuotes Then
        Phosphorus.Log4PStatic.Logger.ExternalWarning "Unclosed quote in parameter for test " & TestName & ": " & paramString
        isValid = False
    End If
    
    If inBraces > 0 Then
        Phosphorus.Log4PStatic.Logger.ExternalWarning "Unclosed brace in parameter for test " & TestName & ": " & paramString
        isValid = False
    End If
    
    If currentParam.Length > 0 And isValid Then
        If isKey Then
            If Len(Trim(keyString)) = 0 Then
                Phosphorus.Log4PStatic.Logger.ExternalWarning "Empty key in dictionary parameter for test " & TestName & ": " & paramString
                isValid = False
            Else
                Set dict = CreateObject("Scripting.Dictionary")
                Dim dictValue As Variant
                On Error Resume Next
                dictValue = ParseParameterValue(currentParam.ToString, TestName)
                On Error GoTo 0
                If Not IsEmpty(dictValue) Then
                    dict.Add keyString, dictValue
                    params.Add dict
                Else
                    Phosphorus.Log4PStatic.Logger.ExternalWarning "Invalid dictionary value skipped for test " & TestName & ": " & keyString & "=" & currentParam.ToString
                End If
            End If
        Else
            Dim finalValue As Variant
            On Error Resume Next
            finalValue = ParseParameterValue(currentParam.ToString, TestName)
            On Error GoTo 0
            If Not IsEmpty(finalValue) Then
                params.Add finalValue
            Else
                Phosphorus.Log4PStatic.Logger.ExternalWarning "Invalid final parameter skipped for test " & TestName & ": " & currentParam.ToString
            End If
        End If
    End If
    
    If Not isValid Or params.count = 0 Then
        Phosphorus.Log4PStatic.Logger.ExternalWarning "No valid parameters parsed for test " & TestName & ": " & paramString
        ParseComplexParameters = Empty
        Exit Function
    End If
    
    ParseComplexParameters = CollectionToArray(params)
End Function

' Parse a single parameter value, handling complex types safely
Private Function ParseParameterValue(paramString As String, TestName As String) As Variant
  
  paramString = Trim(paramString)
    
  If Len(paramString) = 0 Then
    Phosphorus.Log4PStatic.Logger.ExternalWarning "Empty parameter value in @TestData for test " & TestName
    ParseParameterValue = Empty
    Exit Function
  End If
    
  If paramString Like "{*}" Then
    Dim innerContent As String
    innerContent = Mid$(paramString, 2, Len(paramString) - 2)
    If Len(Trim(innerContent)) = 0 Then
      Phosphorus.Log4PStatic.Logger.ExternalWarning "Empty nested structure in @TestData for test " & TestName & ": {}"
      ParseParameterValue = Array()
      Exit Function
    End If
    Dim innerParams As Variant
    On Error Resume Next
    innerParams = ParseComplexParameters(innerContent, TestName)
    On Error GoTo 0
    If Not IsEmpty(innerParams) Then
      ParseParameterValue = innerParams
    Else
      Phosphorus.Log4PStatic.Logger.ExternalWarning "Invalid nested structure in @TestData for test " & TestName & ": {" & innerContent & "}"
      ParseParameterValue = Empty
    End If
  ElseIf InStr(paramString, "=") > 0 And Not (paramString Like """*""") Then
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim pairs As Variant
    On Error Resume Next
    pairs = ParseComplexParameters(paramString, TestName)
    On Error GoTo 0
    If Not IsEmpty(pairs) Then
      Dim pair As Variant
      Dim validPairs As Boolean
      validPairs = False
      For Each pair In pairs
        If typeName(pair) = "Dictionary" Then
          Dim key As Variant
          For Each key In pair.Keys
            dict.Add key, pair(key)
            validPairs = True
          Next key
        End If
      Next pair
      If validPairs Then
        Set ParseParameterValue = dict
      Else
        Phosphorus.Log4PStatic.Logger.ExternalWarning "Invalid dictionary structure in @TestData for test " & TestName & ": " & paramString
        ParseParameterValue = Empty
      End If
    Else
      Phosphorus.Log4PStatic.Logger.ExternalWarning "Invalid dictionary structure in @TestData for test " & TestName & ": " & paramString
      ParseParameterValue = Empty
    End If
  Else
    If paramString Like """*""" Then
      ParseParameterValue = VBA.Strings.Mid$(paramString, 2, Len(paramString) - 2)
    ElseIf VBA.Information.IsNumeric(paramString) Then
      If InStr(paramString, ".") > 0 Then
        ParseParameterValue = VBA.Conversion.CDbl(paramString)
      Else
        If Val(paramString) >= -32768 And Val(paramString) <= 32767 Then
          ParseParameterValue = VBA.Conversion.CInt(paramString)
        Else
          ParseParameterValue = VBA.Conversion.CLng(paramString)
        End If
      End If
    ElseIf UCase(paramString) = "TRUE" Then
      ParseParameterValue = True
    ElseIf UCase(paramString) = "FALSE" Then
      ParseParameterValue = False
    Else
      Phosphorus.Log4PStatic.Logger.ExternalWarning "Unrecognized parameter value in @TestData for test " & TestName & ": " & paramString
      ParseParameterValue = Empty
    End If
  End If
End Function

Private Function CountPassedTests() As Long
  Dim TestResult As pUnit.TestResult
  Dim count As Long
  For Each TestResult In TestResults
    If TestResult.Status = Passed Then
      count = count + 1
    End If
  Next TestResult
  CountPassedTests = count
End Function

Private Function CountSkippedTests() As Long
  Dim TestResult As pUnit.TestResult
  Dim count As Long
  For Each TestResult In TestResults
    If TestResult.Status = skipped Then
      count = count + 1
    End If
  Next TestResult
  CountSkippedTests = count
End Function

' Get annotations for a specific test method
Private Function GetTestAnnotations(vbComp As VBComponent, TestName As String) As String

  Dim codeMod As CodeModule
  Set codeMod = vbComp.CodeModule
  
  Dim annotations As collection
  Set annotations = New collection
  
  Dim lineNum As Long
  lineNum = 1
  
  Dim strResult As String
  
  'Loop through all the lines of code
  While lineNum <= codeMod.CountOfLines
    Dim lineText As String
    lineText = Trim(codeMod.Lines(lineNum, 1))
    'Store any annotations in tehir original format (not parsed!)
    If lineText Like "'@*" Then
      annotations.Add VBA.Strings.Mid(lineText, 3, VBA.Strings.Len(lineText))
    End If
    'Have we reached the test method?
    If lineText Like "Sub *" Or lineText Like "Public Sub *" Or lineText Like "Private Sub *" Then
      'We have reached the test method
      Dim strCurrentMethodName As String
      strCurrentMethodName = codeMod.ProcOfLine(lineNum, vbext_pk_Proc)
      If strCurrentMethodName = TestName Then
        'Return an array of the current collection of annotations
        If annotations.count = 0 Then
          strResult = "none"
        Else
          strResult = VBA.Strings.Join(CollectionToArray(annotations), ", ")
        End If
        GetTestAnnotations = strResult
        Exit Function
      End If
      'Clear down the collection of annotations every time have passed any test method
      Set annotations = New collection
    End If
    lineNum = lineNum + 1
  Wend
  
  GetTestAnnotations = "none"

End Function
