Attribute VB_Name = "Factory"
'@Folder Common
Option Explicit

Public Function GetNewLogger() As Phosphorus.Log4P
  Set GetNewLogger = New Phosphorus.Log4P
End Function

Public Function GetNewPhosphorusPPath() As Phosphorus.PPath
  Set GetNewPhosphorusPPath = New Phosphorus.PPath
End Function

Public Function GetNewPhosphorusPPathReturnClass(intNumberOfPPathExpressions As Integer) As Phosphorus.PPathReturnClass
  Set GetNewPhosphorusPPathReturnClass = New Phosphorus.PPathReturnClass
  GetNewPhosphorusPPathReturnClass.Initialise intNumberOfPPathExpressions
End Function

Public Function GetNewPhosphorusLog4P() As Phosphorus.Log4P
  Set GetNewPhosphorusLog4P = New Phosphorus.Log4P
End Function

Public Function GetNewPDriver(dType As Phosphorus.pWindowsDriverType) As Phosphorus.pWindowsDriver
  Set GetNewPDriver = New Phosphorus.pWindowsDriver
  GetNewPDriver.SetDriverType dType
  If Phosphorus.pWindowsDriverStatic.gCUIAutomation Is Nothing Then
    Set Phosphorus.pWindowsDriverStatic.gCUIAutomation = New CUIAutomation
    If gUIADesktopUIElement Is Nothing Then
      Set gUIADesktopUIElement = Phosphorus.pWindowsDriverStatic.gCUIAutomation.GetRootElement
    End If
  End If
End Function

