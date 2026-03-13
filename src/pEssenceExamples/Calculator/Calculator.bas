Attribute VB_Name = "Calculator"
'@Folder Calculator
Option Explicit

Private Type PublicMenuElementNames
  StandardCalculator As String
  ScientificCalculator As String
  GraphingCalculator As String
  ProgrammerCalculator As String
  DatecalculationCalculator As String
  CurrencyConverter As String
  VolumeConverter As String
  LengthConverter As String
  WeightAndMassConverter As String
  TemperatureConverter As String
  EnergyConverter As String
  AreaConverter As String
  SpeedConverter As String
  TimeConverter As String
  PowerConverter As String
  DataConverter As String
  PressureConverter As String
  AngleConverter As String
End Type

Public HomePageMenuNames As PublicMenuElementNames

Public Sub Calculator()

  On Error GoTo ErrorHandler
    
  GetRootDesktopElement
  SetHomePageMenuNames
  
  Dim HomePage As CalculatorHomePage
  Set HomePage = New CalculatorHomePage
  
  HomePage.SelectCalculatorType HomePageMenuNames.ScientificCalculator
  HomePage.SelectCalculatorType HomePageMenuNames.GraphingCalculator
  HomePage.SelectCalculatorType HomePageMenuNames.ProgrammerCalculator
  HomePage.SelectCalculatorType HomePageMenuNames.DatecalculationCalculator
  
  HomePage.SelectCalculatorType HomePageMenuNames.CurrencyConverter
  HomePage.SelectCalculatorType HomePageMenuNames.VolumeConverter
  HomePage.SelectCalculatorType HomePageMenuNames.LengthConverter
  HomePage.SelectCalculatorType HomePageMenuNames.WeightAndMassConverter
  HomePage.SelectCalculatorType HomePageMenuNames.TemperatureConverter
  HomePage.SelectCalculatorType HomePageMenuNames.AreaConverter
  HomePage.SelectCalculatorType HomePageMenuNames.SpeedConverter
  HomePage.SelectCalculatorType HomePageMenuNames.TimeConverter
  HomePage.SelectCalculatorType HomePageMenuNames.PowerConverter
  HomePage.SelectCalculatorType HomePageMenuNames.DataConverter
  HomePage.SelectCalculatorType HomePageMenuNames.PressureConverter
  HomePage.SelectCalculatorType HomePageMenuNames.AngleConverter
  
'Last
  HomePage.SelectCalculatorType HomePageMenuNames.StandardCalculator

MsgBox "What next? Scroll menu into view. Select each calculate types. Then use buttons!"
  
  Exit Sub
ErrorHandler:
  MsgBox Err.Description & " (Error Number #" & Err.Number & ")"
  Exit Sub
  
ExitSub:
  Set HomePage = Nothing

End Sub

Private Sub SetHomePageMenuNames()
  With HomePageMenuNames
    .StandardCalculator = "Standard Calculator"
    .ScientificCalculator = "Scientific Calculator"
    .GraphingCalculator = "Graphing Calculator"
    .ProgrammerCalculator = "Programmer Calculator"
    .DatecalculationCalculator = "Date calculation Calculator"
    .CurrencyConverter = "Currency Converter"
    .VolumeConverter = "Volume Converter"
    .LengthConverter = "Length Converter"
    .WeightAndMassConverter = "Weight and mass Converter"
    .TemperatureConverter = "Temperature Converter"
    .EnergyConverter = "Energy Converter"
    .AreaConverter = "Area Converter"
    .SpeedConverter = "Speed Converter"
    .TimeConverter = "Time Converter"
    .PowerConverter = "Power Converter"
    .DataConverter = "Data Converter"
    .PressureConverter = "Pressure Converter"
    .AngleConverter = "Angle Converter"
  End With
End Sub
