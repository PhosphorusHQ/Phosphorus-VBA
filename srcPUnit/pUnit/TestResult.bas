VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder pUnit
Option Explicit

Public Enum TestStatus
  Passed = 0
  FAILED = 1
  skipped = 2
End Enum

Public TestName As String
Public Status As TestStatus
Public ErrorMessage As String
Public duration As Double
Public parameters As String
