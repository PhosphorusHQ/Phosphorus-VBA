Attribute VB_Name = "Calculator"
'@Folder Calculator
Option Explicit

Public Type CHPElements
  MasterWindow As String
End Type

Public Sub Calculator()

  On Error GoTo ErrorHandler
  
  GetRootDesktopElement
  
  Dim HomePage As CalculatorHomePage
  Set HomePage = New CalculatorHomePage
  HomePage.SelectCalculatorType "Standard Calculator"
  
MsgBox "What next? List all calculate types. Then use buttons!"
  
  Exit Sub
ErrorHandler:
  MsgBox Err.Description & " (Error Number #" & Err.Number & ")"
  Exit Sub
  
ExitSub:
  Set HomePage = Nothing
   
End Sub
