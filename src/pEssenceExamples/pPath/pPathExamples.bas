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
  
  If Not RunningAllExamples Then
    Window.HighlightElements = True
  End If

  Set pPathPage = New pPathExamplesPage

  'Test all cases with the screen in landscape mode - portrait won't work for the Excel based tests!
  RunPreValidationTests
  RunEvaluationTests
  RunEvaluationExcelTests
'pPathPage.Evaluation_Test109
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
    .Evaluation_Test053
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

