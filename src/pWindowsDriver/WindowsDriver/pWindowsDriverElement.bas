VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pWindowsDriverElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder WindowsDriver
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Private This As WindowsDriverElement

Private Type WindowsDriverElement
  Name As String
  UIAElement As UIAutomationClient.IUIAutomationElement
  FoundBypPath As String
  FoundFromUIAElement As UIAutomationClient.IUIAutomationElement
  ParentWindowsDriver As pWindowsDriver
  Actions As pWindowsDriverElement_Actions
End Type

Public Property Let Name(ByRef Name As String)
  This.Name = Name
End Property

Public Property Get Name() As String
  Name = This.Name
End Property

Public Sub Initialise( _
  ByRef Name As String, _
  ByRef ParentWindowsDriver As pWindowsDriver, _
  ByRef UIAElement As UIAutomationClient.IUIAutomationElement, _
  ByVal FoundBypPath As String, _
  Optional ByRef FoundFromUIAElement As UIAutomationClient.IUIAutomationElement)
  
  This.Name = Name
  Set This.ParentWindowsDriver = ParentWindowsDriver
  Set This.UIAElement = UIAElement
  This.FoundBypPath = FoundBypPath
  If Not FoundFromUIAElement Is Nothing Then
    Set This.FoundFromUIAElement = FoundFromUIAElement
  End If

End Sub

Public Property Get UIAElement() As UIAutomationClient.IUIAutomationElement
  Set UIAElement = This.UIAElement
End Property

Public Property Get FoundBypPath() As String
  FoundBypPath = This.FoundBypPath
End Property

Public Sub WaitForWindowInteractionState( _
  State As UIAutomationClient.WindowInteractionState)
'Wait for given Window Interaction State of this windows element
  
  WaitForElementConditionOrState _
    UIAutomationClient.UIA_PatternIds.UIA_WindowPatternId, _
    State, _
    "Window Interaction State", _
    pWinDriver.pWindowsDriverStatic.GetWindowInteractionStateDescription(State)

End Sub

Private Sub WaitForElementConditionOrState( _
  PatternOrPropertyID As Long, _
  DesiredConditionOrStateID As Long, _
  DescriptionOfPatternOrProperty As String, _
  DescriptionOfDesiredState As String)
'TODO: Do we need to wait for this state after every action on any element, eg click? Or After every element found?
  
  Dim CurrentConditionOrStateID As Long
  Dim ElementIsInCurrentConditionOrStateID As Boolean
  ElementIsInCurrentConditionOrStateID = False
  
  'Pass 1 - test for a matching PatterOrPropertyID
  'TODO: Make this a select case statement
  If PatternOrPropertyID = UIAutomationClient.UIA_PatternIds.UIA_WindowPatternId Then
    Dim WindowPattern As IUIAutomationWindowPattern
    Set WindowPattern = This.UIAElement.GetCurrentPattern(PatternOrPropertyID)
    If WindowPattern Is Nothing Then
      'TODO: Raise Error!
      MsgBox "PJG: Raise No Window Pattern Error Here!"
      Exit Sub
    End If
    CurrentConditionOrStateID = WindowPattern.CurrentWindowInteractionState
  End If

  'Pass 2 - now get the Conditon or State of the Pattern or Property
  'TODO: Make this a select case statement
  If CurrentConditionOrStateID = DesiredConditionOrStateID Then
    ElementIsInCurrentConditionOrStateID = True
  End If
  
  'TODO: Add a wait loop
  
  'Raise an error if the condition hasn't been met in time
  If Not ElementIsInCurrentConditionOrStateID Then
    'Raise Exception
    Phosphorus.pExceptions.Raise _
      Phosphorus.Exceptions.WindowsDriverUIElementCondtionOrStateNotMetBeforeTimeout, _
      DescriptionOfDesiredState, _
      VBA.Conversion.Str(DesiredConditionOrStateID), _
      VBA.Conversion.Str(CurrentConditionOrStateID), _
      VBA.Conversion.Str(This.ParentWindowsDriver.GetDefaultImplicitTimeoutInSeconds)
  End If
  
End Sub

Public Function GetProcessID() As Long
  GetProcessID = This.UIAElement.CurrentProcessId
End Function

'************
'* PATTERNS *
'************

Public Function CloseWindow()
  Dim WindowPattern As UIAutomationClient.IUIAutomationWindowPattern
  Set WindowPattern = GetPattern(UIAutomationClient.UIA_WindowPatternId)
  WindowPattern.Close
End Function

'Use IUnknown to force a casting to correct type correctly - typesafe
Function GetPattern(patternId As Long, Optional CheckOnly As Boolean = False) As IUnknown
  'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controlpattern-idsStop
  Select Case patternId
    Case UIAutomationClient.UIA_WindowPatternId
      Set GetPattern = This.UIAElement.GetCurrentPattern(patternId)
    Case Else
  End Select
  If GetPattern Is Nothing And Not CheckOnly Then
    Phosphorus.pExceptions.Raise _
      WindowsDriverPatternNotHandled, _
      patternId, _
      pWinDriver.UIAutomationStatic.GetPatternName(patternId)
  End If
End Function

Public Function Actions() As pWindowsDriverElement_Actions
  If This.Actions Is Nothing Then
    Set This.Actions = New pWindowsDriverElement_Actions
    This.Actions.Initialise This.UIAElement
    Set Actions = This.Actions
  Else
    Set Actions = This.Actions
  End If
End Function

