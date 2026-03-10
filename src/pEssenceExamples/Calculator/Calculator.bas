Attribute VB_Name = "Calculator"
'@Folder Calculator
Option Explicit

Public Sub Calculator()

  On Error GoTo ErrorHandler
  
  pEssence.UIACommon.GetRootDesktopElement
  
  Dim HomePage As CalculatorHomePage
  Set HomePage = New CalculatorHomePage
  HomePage.SelectCalculatorType "Standard Calculator"
MsgBox "What next? - Select from Navigation menu"
  
  Exit Sub
ErrorHandler:
  MsgBox Err.Description & " (Error Number #" & Err.Number & ")"
  
End Sub
