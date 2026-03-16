Attribute VB_Name = "Calculator"
'@Folder Calculator
Option Explicit

Dim MenuNames As New Collection

Private Sub InitialiseMenuNames()
  MenuNames.Add "Standard Calculator"
  MenuNames.Add "Scientific Calculator"
  MenuNames.Add "Graphing Calculator"
  MenuNames.Add "Programmer Calculator"
  MenuNames.Add "Date calculation Calculator"
  MenuNames.Add "Currency Converter"
  MenuNames.Add "Volume Converter"
  MenuNames.Add "Length Converter"
  MenuNames.Add "Weight and mass Converter"
  MenuNames.Add "Temperature Converter"
  MenuNames.Add "Energy Converter"
  MenuNames.Add "Area Converter"
  MenuNames.Add "Speed Converter"
  MenuNames.Add "Time Converter"
  MenuNames.Add "Power Converter"
  MenuNames.Add "Data Converter"
  MenuNames.Add "Pressure Converter"
  MenuNames.Add "Angle Converter"
End Sub

Public Sub Calculator()

  On Error GoTo ErrorHandler
  
  GetRootDesktopElement
  InitialiseMenuNames
  
  Dim HomePage As CalculatorHomePage
  Set HomePage = New CalculatorHomePage

  Dim i As Integer
  Dim CurrentMenuName As String
  For i = 1 To MenuNames.Count
    Debug.Print CurrentMenuName
    CurrentMenuName = MenuNames(i)
    HomePage.SelectCalculatorType CurrentMenuName
    On Error Resume Next
    Application.Run "pEssenceExamples.Calculator." & VBA.Strings.Replace(CurrentMenuName, " ", "")
    On Error GoTo 0
  Next

MsgBox "What next? Timeout on wait for property value. Use buttons! Close application window."
  
  Exit Sub
ErrorHandler:
  MsgBox Err.Description & " (Error Number #" & Err.Number & ")"
  Exit Sub
  
ExitSub:
  Set HomePage = Nothing

End Sub

Private Sub StandardCalculator()
'  MsgBox "PJG: StandardCalculator"
  'Squared = asc 253 ² https://www.ascii-code.com/character/%C2%B2
  '{SquareRoot} ???
  '{Reciprocal}
  '{BackSpace}
  '%
  '± alt 241 or {Negate}
End Sub
