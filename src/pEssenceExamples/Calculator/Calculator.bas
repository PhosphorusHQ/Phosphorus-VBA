Attribute VB_Name = "Calculator"
'@Folder Calculator
Option Explicit

Dim MenuNames As New Collection
Dim HomePage As CalculatorHomePage

Private Sub InitialiseMenuNames()
  Set MenuNames = Nothing
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
  
  Set HomePage = New CalculatorHomePage
  Dim i As Integer
  Dim CurrentMenuName As String
  For i = 1 To MenuNames.Count
    Debug.Print CurrentMenuName
    CurrentMenuName = MenuNames(i)
    HomePage.SelectCalculatorType CurrentMenuName
    Select Case CurrentMenuName
      Case "Standard Calculator"
        StandardCalculator
    End Select
  Next
  GoTo ExitSub

ErrorHandler:
  MsgBox Err.Description & " (Error Number #" & Err.Number & ")"
  GoTo ExitSub
  
ExitSub:
  Set HomePage = Nothing
MsgBox "What next? Window class. Timeout on wait for property value? Use buttons!"

End Sub

Private Sub StandardCalculator()

  Dim StandardCalculatorPage As CalculatorStandardPage
  Set StandardCalculatorPage = New CalculatorStandardPage
  StandardCalculatorPage.FindAllControls HomePage
  
  With StandardCalculatorPage
    .Calculate "12345678.09=", "12,345,678.09"
    .Calculate "1+2=", "3"
    .Calculate "=", "5"
    .Calculate "=", "7"
    .Calculate "210-105=", "105"
    .Calculate "21x3=", "63"
    .Calculate "21/3=", "7"
    .Calculate "9{Sqr}", "81"
    .Calculate "9{SqrRt}", "3"
    .Calculate "C", "0"
    .Calculate "12{BS}4-3{Negate}=", "17"
    .Calculate "4{1/x}", "0.25"
  .Calculate "6{M+}7{M+}{MR}", "13"
  End With
  
  Set StandardCalculatorPage = Nothing

End Sub
