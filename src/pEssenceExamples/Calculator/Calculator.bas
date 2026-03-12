Attribute VB_Name = "Calculator"
'@Folder Calculator
Option Explicit

Private Type PublicMenuElementNames
  StandardCalculator As String
End Type

Public HomePageMenuNames As PublicMenuElementNames
Public HomePageMenuElements As Scripting.Dictionary

Public Sub Calculator()

  On Error GoTo ErrorHandler
    
  GetRootDesktopElement
  
  Dim HomePage As CalculatorHomePage
  Set HomePage = New CalculatorHomePage
  
  HomePage.SelectCalculatorType HomePageMenuNames.StandardCalculator

MsgBox "What next? Select each calculate types. Then use buttons!"
  
  Exit Sub
ErrorHandler:
  MsgBox Err.Description & " (Error Number #" & Err.Number & ")"
  Exit Sub
  
ExitSub:
  Set HomePage = Nothing
  Set HomePageMenuElements = Nothing

End Sub

Private Sub SetHomePageMenuNames()
  HomePageMenuNames.StandardCalculator = "StandardCalculator"
End Sub
