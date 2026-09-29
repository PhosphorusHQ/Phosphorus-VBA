Attribute VB_Name = "pPathExamples"
'@Folder pPath
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Private pPathPage As pPathExamplesPage

Public Sub pPathExamples()

  'pPath always requires the Logger class
  Phosphorus.Log4PStatic.GetLogger

  On Error GoTo ErrorHandler
    
  'Always HighlightElements for pPath examples
  Window.HighlightElements = True

  Set pPathPage = New pPathExamplesPage
  
  'Test all cases with the screen in landscape mode - portrait won't work for the Excel based tests!
  RunPreValidationTests
  RunEvaluationTests
  RunEvaluationExcelTests
'Phosphorus.Utils.PJGDebugMode = False
'pPathPage.Evaluation_TestExcel008
'pPathPage.Evaluation_TestExcel008
'Phosphorus.Utils.PJGDebugMode = False
'pPathPage.Evaluation_Test053
  GoTo ExitSub
  
ErrorHandler:
  MsgBox Err.Description & " (Error Number #" & Err.Number & ")"
  GoTo ExitSub
  
ExitSub:
  If Not RunningAllExamples Then
    Window.HighlightElements = False
  End If
  Set pPathPage = Nothing

  Phosphorus.Log4PStatic.CloseLogger

End Sub

Private Sub RunPreValidationTests()
  With pPathPage
    .PreValidation_Test01
    .PreValidation_Test02
    .PreValidation_Test03
    .PreValidation_Test04
    .PreValidation_Test05
    .PreValidation_Test06
    .PreValidation_Test07
    .PreValidation_Test08
    .PreValidation_Test09
    .PreValidation_Test10
    .PreValidation_Test11
    .PreValidation_Test12
    .PreValidation_Test13
    .PreValidation_Test14
    .PreValidation_Test15
  End With
End Sub

Private Sub RunEvaluationTests()
  With pPathPage
    .Evaluation_Test001
    .Evaluation_Test002
    .Evaluation_Test003
    .Evaluation_Test004
    .Evaluation_Test005
    .Evaluation_Test006
    .Evaluation_Test007
    .Evaluation_Test008
    .Evaluation_Test009
    .Evaluation_Test010
    .Evaluation_Test011
    .Evaluation_Test012
    .Evaluation_Test013
    .Evaluation_Test014
    .Evaluation_Test015
    .Evaluation_Test016
    .Evaluation_Test017
    .Evaluation_Test018
    .Evaluation_Test019
    .Evaluation_Test020
    .Evaluation_Test021
    .Evaluation_Test022
    .Evaluation_Test023
    .Evaluation_Test024
    .Evaluation_Test025
    .Evaluation_Test026
    .Evaluation_Test027
    .Evaluation_Test028
    .Evaluation_Test029
    .Evaluation_Test030
    .Evaluation_Test031
    .Evaluation_Test032
    .Evaluation_Test033
    .Evaluation_Test034
    .Evaluation_Test035
    .Evaluation_Test036
    .Evaluation_Test037
    .Evaluation_Test038
    .Evaluation_Test039
    .Evaluation_Test040
    .Evaluation_Test041
    .Evaluation_Test042
    .Evaluation_Test043
    .Evaluation_Test044
    .Evaluation_Test045
    .Evaluation_Test046
    .Evaluation_Test047
    .Evaluation_Test048
    .Evaluation_Test049
    .Evaluation_Test050
    .Evaluation_Test051
    .Evaluation_Test052
'PJG    .Evaluation_Test053
    .Evaluation_Test054
    .Evaluation_Test055
    .Evaluation_Test056
    .Evaluation_Test057
    .Evaluation_Test058
    .Evaluation_Test059
    .Evaluation_Test060
    .Evaluation_Test061
    .Evaluation_Test062
    .Evaluation_Test063
    .Evaluation_Test064
    .Evaluation_Test065
    .Evaluation_Test066
    .Evaluation_Test067
    .Evaluation_Test068
    .Evaluation_Test069
    .Evaluation_Test070
    .Evaluation_Test071
    .Evaluation_Test072
    .Evaluation_Test073
    .Evaluation_Test074
    .Evaluation_Test075
    .Evaluation_Test076
    .Evaluation_Test077
    .Evaluation_Test078
    .Evaluation_Test079
    .Evaluation_Test080
    .Evaluation_Test081a
    .Evaluation_Test081b
    .Evaluation_Test081c
    .Evaluation_Test081d
    .Evaluation_Test081e
    .Evaluation_Test082
    .Evaluation_Test083
    .Evaluation_Test084
    .Evaluation_Test085
    .Evaluation_Test086
    .Evaluation_Test087
    .Evaluation_Test088
    .Evaluation_Test089
    .Evaluation_Test090
    .Evaluation_Test091
    .Evaluation_Test092
    .Evaluation_Test093
    .Evaluation_Test094
    .Evaluation_Test095
    .Evaluation_Test096
    .Evaluation_Test097
    .Evaluation_Test098
    .Evaluation_Test099
    .Evaluation_Test100
    .Evaluation_Test101
    .Evaluation_Test102
    .Evaluation_Test103a
    .Evaluation_Test103b
    .Evaluation_Test104a
    .Evaluation_Test104b
    .Evaluation_Test105
    .Evaluation_Test106
    .Evaluation_Test107
    .Evaluation_Test108
    .Evaluation_Test109
    .Evaluation_Test110
    .Evaluation_Test111
    .Evaluation_Test112
    .Evaluation_Test113
    .Evaluation_Test114
    .Evaluation_Test115
    .Evaluation_Test116
    .Evaluation_Test117
    .Evaluation_Test118
    .Evaluation_Test119
    .Evaluation_Test120
    .Evaluation_Test121
    .Evaluation_Test122
    .Evaluation_Test123
    .Evaluation_Test124
    .Evaluation_Test125
    .Evaluation_Test126
    .Evaluation_Test127
    .Evaluation_Test128
    .Evaluation_Test129
    .Evaluation_Test131
    .Evaluation_Test132a
    .Evaluation_Test132b
    .Evaluation_Test133a
    .Evaluation_Test134
  End With
End Sub

Private Sub RunEvaluationExcelTests()
  With pPathPage
    .Evaluation_TestExcel001
    .Evaluation_TestExcel002
    .Evaluation_TestExcel003
    .Evaluation_TestExcel004
    .Evaluation_TestExcel005
    .Evaluation_TestExcel006
    .Evaluation_TestExcel007
    .Evaluation_TestExcel008
    .Evaluation_TestExcel009
    .Evaluation_TestExcel010
    .Evaluation_TestExcel011
    .Evaluation_TestExcel012
    .Evaluation_TestExcel013
    .Evaluation_TestExcel014
    .Evaluation_TestExcel015
    .Evaluation_TestExcel016
    .Evaluation_TestExcel017
    .Evaluation_TestExcel018
    .Evaluation_TestExcel019
    .Evaluation_TestExcel020
    .Evaluation_TestExcel100
    .Evaluation_TestExcel101
    .Evaluation_TestExcel102
    .Evaluation_TestExcel103
    .Evaluation_TestExcel104
    .Evaluation_TestExcel105
    .Evaluation_TestExcel106
    .Evaluation_TestExcel107
    .Evaluation_TestExcel108
    .Evaluation_TestExcel109
    .Evaluation_TestExcel110
    .Evaluation_TestExcel111
    .Evaluation_TestExcel113
    .Evaluation_TestExcel114
    .Evaluation_TestExcel115
    .Evaluation_TestExcel116
    .Evaluation_TestExcel117
    .Evaluation_TestExcel118
    .Evaluation_TestExcel119
    .Evaluation_TestExcel120
    .Evaluation_TestExcel121
  End With
End Sub

Public Sub RunAllpPathTestsWithMultipleTries()
  RunAllPreValidationTests
  RunAllEvaluationTests
  RunAllEvaluationExcelTests
End Sub

Private Sub RunAllPreValidationTests()
  RunApPathTest "PreValidation_Test01"
  RunApPathTest "PreValidation_Test02"
  RunApPathTest "PreValidation_Test03"
  RunApPathTest "PreValidation_Test04"
  RunApPathTest "PreValidation_Test05"
  RunApPathTest "PreValidation_Test06"
  RunApPathTest "PreValidation_Test07"
  RunApPathTest "PreValidation_Test08"
  RunApPathTest "PreValidation_Test09"
  RunApPathTest "PreValidation_Test10"
  RunApPathTest "PreValidation_Test11"
  RunApPathTest "PreValidation_Test12"
  RunApPathTest "PreValidation_Test13"
  RunApPathTest "PreValidation_Test14"
  RunApPathTest "PreValidation_Test15"
End Sub

Private Sub RunAllEvaluationTests()
  RunApPathTest "Evaluation_Test001"
  RunApPathTest "Evaluation_Test002"
  RunApPathTest "Evaluation_Test003"
  RunApPathTest "Evaluation_Test004"
  RunApPathTest "Evaluation_Test005"
  RunApPathTest "Evaluation_Test006"
  RunApPathTest "Evaluation_Test007"
  RunApPathTest "Evaluation_Test008"
  RunApPathTest "Evaluation_Test009"
  RunApPathTest "Evaluation_Test010"
  RunApPathTest "Evaluation_Test011"
  RunApPathTest "Evaluation_Test012"
  RunApPathTest "Evaluation_Test013"
  RunApPathTest "Evaluation_Test014"
  RunApPathTest "Evaluation_Test015"
  RunApPathTest "Evaluation_Test016"
  RunApPathTest "Evaluation_Test017"
  RunApPathTest "Evaluation_Test018"
  RunApPathTest "Evaluation_Test019"
  RunApPathTest "Evaluation_Test020"
  RunApPathTest "Evaluation_Test021"
  RunApPathTest "Evaluation_Test022"
  RunApPathTest "Evaluation_Test023"
  RunApPathTest "Evaluation_Test024"
  RunApPathTest "Evaluation_Test025"
  RunApPathTest "Evaluation_Test026"
  RunApPathTest "Evaluation_Test027"
  RunApPathTest "Evaluation_Test028"
  RunApPathTest "Evaluation_Test029"
  RunApPathTest "Evaluation_Test030"
  RunApPathTest "Evaluation_Test031"
  RunApPathTest "Evaluation_Test032"
  RunApPathTest "Evaluation_Test033"
  RunApPathTest "Evaluation_Test034"
  RunApPathTest "Evaluation_Test035"
  RunApPathTest "Evaluation_Test036"
  RunApPathTest "Evaluation_Test037"
  RunApPathTest "Evaluation_Test038"
  RunApPathTest "Evaluation_Test039"
  RunApPathTest "Evaluation_Test040"
  RunApPathTest "Evaluation_Test041"
  RunApPathTest "Evaluation_Test042"
  RunApPathTest "Evaluation_Test043"
  RunApPathTest "Evaluation_Test044"
  RunApPathTest "Evaluation_Test045"
  RunApPathTest "Evaluation_Test046"
  RunApPathTest "Evaluation_Test047"
  RunApPathTest "Evaluation_Test048"
  RunApPathTest "Evaluation_Test049"
  RunApPathTest "Evaluation_Test050"
  RunApPathTest "Evaluation_Test051"
  RunApPathTest "Evaluation_Test052"
'PJG
  RunApPathTest "Evaluation_Test053"
  RunApPathTest "Evaluation_Test054"
  RunApPathTest "Evaluation_Test055"
  RunApPathTest "Evaluation_Test056"
  RunApPathTest "Evaluation_Test057"
  RunApPathTest "Evaluation_Test058"
  RunApPathTest "Evaluation_Test059"
  RunApPathTest "Evaluation_Test060"
  RunApPathTest "Evaluation_Test061"
  RunApPathTest "Evaluation_Test062"
  RunApPathTest "Evaluation_Test063"
  RunApPathTest "Evaluation_Test064"
  RunApPathTest "Evaluation_Test065"
  RunApPathTest "Evaluation_Test066"
  RunApPathTest "Evaluation_Test067"
  RunApPathTest "Evaluation_Test068"
  RunApPathTest "Evaluation_Test069"
  RunApPathTest "Evaluation_Test070"
  RunApPathTest "Evaluation_Test071"
  RunApPathTest "Evaluation_Test072"
  RunApPathTest "Evaluation_Test073"
  RunApPathTest "Evaluation_Test074"
  RunApPathTest "Evaluation_Test075"
  RunApPathTest "Evaluation_Test076"
  RunApPathTest "Evaluation_Test077"
  RunApPathTest "Evaluation_Test078"
  RunApPathTest "Evaluation_Test079"
  RunApPathTest "Evaluation_Test080"
  RunApPathTest "Evaluation_Test081a"
  RunApPathTest "Evaluation_Test081b"
  RunApPathTest "Evaluation_Test081c"
  RunApPathTest "Evaluation_Test081d"
  RunApPathTest "Evaluation_Test081e"
  RunApPathTest "Evaluation_Test082"
  RunApPathTest "Evaluation_Test083"
  RunApPathTest "Evaluation_Test084"
  RunApPathTest "Evaluation_Test085"
  RunApPathTest "Evaluation_Test086"
  RunApPathTest "Evaluation_Test087"
  RunApPathTest "Evaluation_Test088"
  RunApPathTest "Evaluation_Test089"
  RunApPathTest "Evaluation_Test090"
  RunApPathTest "Evaluation_Test091"
  RunApPathTest "Evaluation_Test092"
  RunApPathTest "Evaluation_Test093"
  RunApPathTest "Evaluation_Test094"
  RunApPathTest "Evaluation_Test095"
  RunApPathTest "Evaluation_Test096"
  RunApPathTest "Evaluation_Test097"
  RunApPathTest "Evaluation_Test098"
  RunApPathTest "Evaluation_Test099"
  RunApPathTest "Evaluation_Test100"
  RunApPathTest "Evaluation_Test101"
  RunApPathTest "Evaluation_Test102"
  RunApPathTest "Evaluation_Test103a"
  RunApPathTest "Evaluation_Test103b"
  RunApPathTest "Evaluation_Test104a"
  RunApPathTest "Evaluation_Test104b"
  RunApPathTest "Evaluation_Test105"
  RunApPathTest "Evaluation_Test106"
  RunApPathTest "Evaluation_Test107"
  RunApPathTest "Evaluation_Test108"
  RunApPathTest "Evaluation_Test109"
  RunApPathTest "Evaluation_Test110"
  RunApPathTest "Evaluation_Test111"
  RunApPathTest "Evaluation_Test112"
  RunApPathTest "Evaluation_Test113"
  RunApPathTest "Evaluation_Test114"
  RunApPathTest "Evaluation_Test115"
  RunApPathTest "Evaluation_Test116"
  RunApPathTest "Evaluation_Test117"
  RunApPathTest "Evaluation_Test118"
  RunApPathTest "Evaluation_Test119"
  RunApPathTest "Evaluation_Test120"
  RunApPathTest "Evaluation_Test121"
  RunApPathTest "Evaluation_Test122"
  RunApPathTest "Evaluation_Test123"
  RunApPathTest "Evaluation_Test124"
  RunApPathTest "Evaluation_Test125"
  RunApPathTest "Evaluation_Test126"
  RunApPathTest "Evaluation_Test127"
  RunApPathTest "Evaluation_Test128"
  RunApPathTest "Evaluation_Test129"
  RunApPathTest "Evaluation_Test131"
  RunApPathTest "Evaluation_Test132a"
  RunApPathTest "Evaluation_Test132b"
  RunApPathTest "Evaluation_Test133a"
  RunApPathTest "Evaluation_Test134"
End Sub

Private Sub RunAllEvaluationExcelTests()
  RunApPathTest "Evaluation_TestExcel001"
  RunApPathTest "Evaluation_TestExcel002"
  RunApPathTest "Evaluation_TestExcel003"
  RunApPathTest "Evaluation_TestExcel004"
  RunApPathTest "Evaluation_TestExcel005"
  RunApPathTest "Evaluation_TestExcel006"
  RunApPathTest "Evaluation_TestExcel007"
  RunApPathTest "Evaluation_TestExcel008"
  RunApPathTest "Evaluation_TestExcel009"
  RunApPathTest "Evaluation_TestExcel010"
  RunApPathTest "Evaluation_TestExcel011"
  RunApPathTest "Evaluation_TestExcel012"
  RunApPathTest "Evaluation_TestExcel013"
  RunApPathTest "Evaluation_TestExcel014"
  RunApPathTest "Evaluation_TestExcel015"
  RunApPathTest "Evaluation_TestExcel016"
  RunApPathTest "Evaluation_TestExcel017"
  RunApPathTest "Evaluation_TestExcel018"
  RunApPathTest "Evaluation_TestExcel019"
  RunApPathTest "Evaluation_TestExcel020"
  RunApPathTest "Evaluation_TestExcel100"
  RunApPathTest "Evaluation_TestExcel101"
  RunApPathTest "Evaluation_TestExcel102"
  RunApPathTest "Evaluation_TestExcel103"
  RunApPathTest "Evaluation_TestExcel104"
  RunApPathTest "Evaluation_TestExcel105"
  RunApPathTest "Evaluation_TestExcel106"
  RunApPathTest "Evaluation_TestExcel107"
  RunApPathTest "Evaluation_TestExcel108"
  RunApPathTest "Evaluation_TestExcel109"
  RunApPathTest "Evaluation_TestExcel110"
  RunApPathTest "Evaluation_TestExcel111"
  RunApPathTest "Evaluation_TestExcel113"
  RunApPathTest "Evaluation_TestExcel114"
  RunApPathTest "Evaluation_TestExcel115"
  RunApPathTest "Evaluation_TestExcel116"
  RunApPathTest "Evaluation_TestExcel117"
  RunApPathTest "Evaluation_TestExcel118"
  RunApPathTest "Evaluation_TestExcel119"
  RunApPathTest "Evaluation_TestExcel120"
  RunApPathTest "Evaluation_TestExcel121"
End Sub

Private Sub RunApPathTest(TestName As String)
  AllExamples.TryToRunAnExampleMultipleTimes False, Examples.pPathExample_SubTest, TestName
End Sub

Public Function RunASinglepPathTest(TestName As String) As Boolean

  Dim Succeeded As Boolean
  Succeeded = False
  
  On Error GoTo ErrorHandler
  
  'Always HighlightElements for pPath examples
  Window.HighlightElements = True
  
  Set pPathPage = New pPathExamplesPage
  
  With pPathPage
    
    Select Case TestName
    
      'PreValidation_Tests
      Case "PreValidation_Test01"
        .PreValidation_Test01
      Case "PreValidation_Test02"
        .PreValidation_Test02
      Case "PreValidation_Test03"
      .PreValidation_Test03
      Case "PreValidation_Test04"
        .PreValidation_Test04
      Case "PreValidation_Test05"
        .PreValidation_Test05
      Case "PreValidation_Test06"
        .PreValidation_Test06
      Case "PreValidation_Test07"
        .PreValidation_Test07
      Case "PreValidation_Test08"
        .PreValidation_Test08
      Case "PreValidation_Test09"
        .PreValidation_Test09
      Case "PreValidation_Test10"
        .PreValidation_Test10
      Case "PreValidation_Test11"
        .PreValidation_Test11
      Case "PreValidation_Test12"
        pPathPage.PreValidation_Test12
      Case "PreValidation_Test13"
        pPathPage.PreValidation_Test13
      Case "PreValidation_Test14"
        .PreValidation_Test14
      Case "PreValidation_Test15"
        .PreValidation_Test15

      'Evaluation_Tests
      Case "Evaluation_Test001"
        .Evaluation_Test002
      Case "Evaluation_Test002"
        .Evaluation_Test002
      Case "Evaluation_Test003"
        .Evaluation_Test003
      Case "Evaluation_Test004"
        .Evaluation_Test004
      Case "Evaluation_Test005"
        .Evaluation_Test005
      Case "Evaluation_Test006"
        .Evaluation_Test006
      Case "Evaluation_Test007"
        .Evaluation_Test007
      Case "Evaluation_Test008"
        .Evaluation_Test008
      Case "Evaluation_Test009"
        .Evaluation_Test009
      Case "Evaluation_Test010"
        .Evaluation_Test010
      Case "Evaluation_Test011"
        .Evaluation_Test011
      Case "Evaluation_Test012"
        .Evaluation_Test012
      Case "Evaluation_Test013"
        .Evaluation_Test013
      Case "Evaluation_Test014"
        .Evaluation_Test014
      Case "Evaluation_Test015"
        .Evaluation_Test015
      Case "Evaluation_Test016"
        .Evaluation_Test016
      Case "Evaluation_Test017"
        .Evaluation_Test017
      Case "Evaluation_Test018"
        .Evaluation_Test018
      Case "Evaluation_Test019"
        .Evaluation_Test019
      Case "Evaluation_Test020"
        .Evaluation_Test020
      Case "Evaluation_Test021"
        .Evaluation_Test021
      Case "Evaluation_Test022"
        .Evaluation_Test022
      Case "Evaluation_Test023"
        .Evaluation_Test023
      Case "Evaluation_Test024"
        .Evaluation_Test024
      Case "Evaluation_Test025"
        .Evaluation_Test025
      Case "Evaluation_Test026"
        .Evaluation_Test026
      Case "Evaluation_Test027"
        .Evaluation_Test027
      Case "Evaluation_Test028"
        .Evaluation_Test028
      Case "Evaluation_Test029"
        .Evaluation_Test029
      Case "Evaluation_Test030"
        .Evaluation_Test030
      Case "Evaluation_Test031"
        .Evaluation_Test031
      Case "Evaluation_Test032"
        .Evaluation_Test032
      Case "Evaluation_Test033"
        .Evaluation_Test033
      Case "Evaluation_Test034"
        .Evaluation_Test034
      Case "Evaluation_Test035"
        .Evaluation_Test035
      Case "Evaluation_Test036"
        .Evaluation_Test036
      Case "Evaluation_Test037"
        .Evaluation_Test037
      Case "Evaluation_Test038"
        .Evaluation_Test038
      Case "Evaluation_Test001"
        .Evaluation_Test001
      Case "Evaluation_Test001"
        .Evaluation_Test001
      Case "Evaluation_Test039"
        .Evaluation_Test039
      Case "Evaluation_Test040"
        .Evaluation_Test040
      Case "Evaluation_Test041"
        .Evaluation_Test041
      Case "Evaluation_Test042"
        .Evaluation_Test042
      Case "Evaluation_Test043"
        .Evaluation_Test043
      Case "Evaluation_Test044"
        .Evaluation_Test044
      Case "Evaluation_Test045"
        .Evaluation_Test045
      Case "Evaluation_Test046"
        .Evaluation_Test046
      Case "Evaluation_Test047"
        .Evaluation_Test047
      Case "Evaluation_Test048"
        .Evaluation_Test048
      Case "Evaluation_Test049"
        .Evaluation_Test049
      Case "Evaluation_Test050"
        .Evaluation_Test050
      Case "Evaluation_Test051"
        .Evaluation_Test051
      Case "Evaluation_Test052"
        .Evaluation_Test052
      Case "Evaluation_Test053"
        .Evaluation_Test053
      Case "Evaluation_Test054"
        .Evaluation_Test054
      Case "Evaluation_Test055"
        .Evaluation_Test055
      Case "Evaluation_Test056"
        .Evaluation_Test056
      Case "Evaluation_Test057"
        .Evaluation_Test057
      Case "Evaluation_Test058"
        .Evaluation_Test058
      Case "Evaluation_Test059"
        .Evaluation_Test059
      Case "Evaluation_Test060"
        .Evaluation_Test060
      Case "Evaluation_Test061"
        .Evaluation_Test061
      Case "Evaluation_Test062"
        .Evaluation_Test062
      Case "Evaluation_Test063"
        .Evaluation_Test063
      Case "Evaluation_Test064"
        .Evaluation_Test064
      Case "Evaluation_Test065"
        .Evaluation_Test065
      Case "Evaluation_Test066"
        .Evaluation_Test066
      Case "Evaluation_Test067"
        .Evaluation_Test067
      Case "Evaluation_Test068"
        .Evaluation_Test068
      Case "Evaluation_Test069"
        .Evaluation_Test069
      Case "Evaluation_Test070"
        .Evaluation_Test070
      Case "Evaluation_Test071"
        .Evaluation_Test071
      Case "Evaluation_Test072"
        .Evaluation_Test072
      Case "Evaluation_Test073"
        .Evaluation_Test073
      Case "Evaluation_Test074"
        .Evaluation_Test074
      Case "Evaluation_Test075"
        .Evaluation_Test075
      Case "Evaluation_Test076"
        .Evaluation_Test076
      Case "Evaluation_Test077"
        .Evaluation_Test077
      Case "Evaluation_Test078"
        .Evaluation_Test078
      Case "Evaluation_Test079"
        .Evaluation_Test079
      Case "Evaluation_Test080"
        .Evaluation_Test080
      Case "Evaluation_Test081a"
        .Evaluation_Test081a
      Case "Evaluation_Test081b"
        .Evaluation_Test081b
      Case "Evaluation_Test081c"
        .Evaluation_Test081c
      Case "Evaluation_Test081d"
        .Evaluation_Test081d
      Case "Evaluation_Test081e"
        .Evaluation_Test081e
      Case "Evaluation_Test082"
        .Evaluation_Test082
      Case "Evaluation_Test083"
        .Evaluation_Test083
      Case "Evaluation_Test084"
        .Evaluation_Test084
      Case "Evaluation_Test085"
        .Evaluation_Test085
      Case "Evaluation_Test086"
        .Evaluation_Test086
      Case "Evaluation_Test087"
        .Evaluation_Test087
      Case "Evaluation_Test088"
        .Evaluation_Test088
      Case "Evaluation_Test089"
        .Evaluation_Test089
      Case "Evaluation_Test090"
        .Evaluation_Test090
      Case "Evaluation_Test091"
        .Evaluation_Test091
      Case "Evaluation_Test092"
        .Evaluation_Test092
      Case "Evaluation_Test093"
        .Evaluation_Test093
      Case "Evaluation_Test094"
        .Evaluation_Test094
      Case "Evaluation_Test095"
        .Evaluation_Test095
      Case "Evaluation_Test096"
        .Evaluation_Test096
      Case "Evaluation_Test097"
        .Evaluation_Test097
      Case "Evaluation_Test098"
        .Evaluation_Test098
      Case "Evaluation_Test099"
        .Evaluation_Test099
      Case "Evaluation_Test100"
        .Evaluation_Test100
      Case "Evaluation_Test101"
        .Evaluation_Test101
      Case "Evaluation_Test102"
        .Evaluation_Test102
      Case "Evaluation_Test103a"
        .Evaluation_Test103a
      Case "Evaluation_Test103b"
        .Evaluation_Test103b
      Case "Evaluation_Test104a"
        .Evaluation_Test104a
      Case "Evaluation_Test104b"
        .Evaluation_Test104b
      Case "Evaluation_Test105"
        .Evaluation_Test105
      Case "Evaluation_Test106"
        .Evaluation_Test106
      Case "Evaluation_Test107"
        .Evaluation_Test107
      Case "Evaluation_Test108"
        .Evaluation_Test108
      Case "Evaluation_Test109"
        .Evaluation_Test109
      Case "Evaluation_Test110"
        .Evaluation_Test110
      Case "Evaluation_Test111"
        .Evaluation_Test111
      Case "Evaluation_Test112"
        .Evaluation_Test112
      Case "Evaluation_Test113"
        .Evaluation_Test113
      Case "Evaluation_Test114"
        .Evaluation_Test114
      Case "Evaluation_Test115"
        .Evaluation_Test115
      Case "Evaluation_Test116"
        .Evaluation_Test116
      Case "Evaluation_Test117"
        .Evaluation_Test117
      Case "Evaluation_Test118"
        .Evaluation_Test118
      Case "Evaluation_Test119"
        .Evaluation_Test119
      Case "Evaluation_Test120"
        .Evaluation_Test120
      Case "Evaluation_Test121"
        .Evaluation_Test121
      Case "Evaluation_Test122"
        .Evaluation_Test122
      Case "Evaluation_Test123"
        .Evaluation_Test123
      Case "Evaluation_Test124"
        .Evaluation_Test124
      Case "Evaluation_Test125"
        .Evaluation_Test125
      Case "Evaluation_Test126"
        .Evaluation_Test126
      Case "Evaluation_Test127"
        .Evaluation_Test127
      Case "Evaluation_Test128"
        .Evaluation_Test128
      Case "Evaluation_Test129"
        .Evaluation_Test129
      Case "Evaluation_Test130"
        .Evaluation_Test130
      Case "Evaluation_Test131"
        .Evaluation_Test131
      Case "Evaluation_Test132a"
        .Evaluation_Test132a
      Case "Evaluation_Test132b"
        .Evaluation_Test132b
      Case "Evaluation_Test133a"
        .Evaluation_Test133a
      Case "Evaluation_Test133b"
        .Evaluation_Test133b
      Case "Evaluation_Test134"
        .Evaluation_Test134
      
      'Evaluation_Tests
      Case "Evaluation_TestExcel001"
        .Evaluation_TestExcel001
      Case "Evaluation_TestExcel002"
        .Evaluation_TestExcel002
      Case "Evaluation_TestExcel003"
        .Evaluation_TestExcel003
      Case "Evaluation_TestExcel004"
        .Evaluation_TestExcel004
      Case "Evaluation_TestExcel005"
        .Evaluation_TestExcel005
      Case "Evaluation_TestExcel006"
        .Evaluation_TestExcel006
      Case "Evaluation_TestExcel007"
        .Evaluation_TestExcel007
      Case "Evaluation_TestExcel008"
        .Evaluation_TestExcel008
      Case "Evaluation_TestExcel009"
        .Evaluation_TestExcel009
      Case "Evaluation_TestExcel010"
        .Evaluation_TestExcel010
      Case "Evaluation_TestExcel011"
        .Evaluation_TestExcel011
      Case "Evaluation_TestExcel012"
        .Evaluation_TestExcel012
      Case "Evaluation_TestExcel013"
        .Evaluation_TestExcel013
      Case "Evaluation_TestExcel014"
        .Evaluation_TestExcel014
      Case "Evaluation_TestExcel015"
        .Evaluation_TestExcel015
      Case "Evaluation_TestExcel016"
        .Evaluation_TestExcel016
      Case "Evaluation_TestExcel017"
        .Evaluation_TestExcel017
      Case "Evaluation_TestExcel018"
        .Evaluation_TestExcel018
      Case "Evaluation_TestExcel019"
        .Evaluation_TestExcel019
      Case "Evaluation_TestExcel020"
        .Evaluation_TestExcel020
      Case "Evaluation_TestExcel100"
        .Evaluation_TestExcel100
      Case "Evaluation_TestExcel101"
        .Evaluation_TestExcel101
      Case "Evaluation_TestExcel102"
        .Evaluation_TestExcel102
      Case "Evaluation_TestExcel103"
        .Evaluation_TestExcel103
      Case "Evaluation_TestExcel104"
        .Evaluation_TestExcel104
      Case "Evaluation_TestExcel105"
        .Evaluation_TestExcel105
      Case "Evaluation_TestExcel106"
        .Evaluation_TestExcel106
      Case "Evaluation_TestExcel107"
        .Evaluation_TestExcel107
      Case "Evaluation_TestExcel108"
        .Evaluation_TestExcel108
      Case "Evaluation_TestExcel109"
        .Evaluation_TestExcel109
      Case "Evaluation_TestExcel110"
        .Evaluation_TestExcel110
      Case "Evaluation_TestExcel111"
        .Evaluation_TestExcel111
      Case "Evaluation_TestExcel112"
        .Evaluation_TestExcel112
      Case "Evaluation_TestExcel113"
        .Evaluation_TestExcel113
      Case "Evaluation_TestExcel114"
        .Evaluation_TestExcel114
      Case "Evaluation_TestExcel115"
        .Evaluation_TestExcel115
      Case "Evaluation_TestExcel116"
        .Evaluation_TestExcel116
      Case "Evaluation_TestExcel117"
        .Evaluation_TestExcel117
      Case "Evaluation_TestExcel118"
        .Evaluation_TestExcel118
      Case "Evaluation_TestExcel119"
        .Evaluation_TestExcel119
      Case "Evaluation_TestExcel120"
        .Evaluation_TestExcel120
      Case "Evaluation_TestExcel121"
        .Evaluation_TestExcel121
      Case "Evaluation_TestExcel122"
        .Evaluation_TestExcel122

      Case Else
        MsgBox "Unhandled TestName: " & TestName
Stop
    End Select
  
  End With
  
  Succeeded = True
  GoTo ExitSub

ErrorHandler:
  If RunningAllExamples Then
    Debug.Print Err.Description & " (Error Number #" & Err.Number & ") in '" & TestName & "' Example!"
  Else
    MsgBox Err.Description & " (Error Number #" & Err.Number & ") in '" & TestName & "' Example!", vbCritical
  End If
  GoTo ExitSub
  
ExitSub:
  If Not RunningAllExamples Then
    Window.HighlightElements = False
  End If
  Set pPathPage = Nothing
  
  RunASinglepPathTest = Succeeded

End Function
