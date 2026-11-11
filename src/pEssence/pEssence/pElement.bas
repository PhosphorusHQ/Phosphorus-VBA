VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder pEssence
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public GivenName As String
Public UIAElement As IUIAutomationElement
Public ParentLocator As pLocator
Public ParentElement As pElement

'Declarations for SendInput
#If VBA7 Then
  Private Declare PtrSafe Function SendInput Lib "user32" ( _
    ByVal nInputs As LongPtr, _
    ByRef pInputs As Any, _
    ByVal cbSize As LongPtr) As LongPtr

  Private Declare PtrSafe Function VkKeyScan Lib "user32" Alias "VkKeyScanA" ( _
    ByVal cChar As Byte) As Integer

  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  Private Declare Function SendInput Lib "user32" ( _
    ByVal nInputs As Long, _
    ByRef pInputs As Any, _
    ByVal cbSize As Long) As Long

  Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" ( _
    ByVal cChar As Byte) As Integer

  Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const INPUT_KEYBOARD As Long = 1
Private Const KEYEVENTF_KEYUP As Long = &H2
Private Const KEYEVENTF_UNICODE As Long = &H4
Private Const KEYEVENTF_EXTENDEDKEY As Long = &H1

' Common virtual-key codes
Private Const VK_CONTROL As Integer = &H11
Private Const VK_SHIFT As Integer = &H10
Private Const VK_MENU As Integer = &H12      ' Alt
Private Const VK_RETURN As Integer = &HD
Private Const VK_TAB As Integer = &H9
Private Const VK_ESCAPE As Integer = &H1B
Private Const VK_BACK As Integer = &H8

Private Type KEYBDINPUT
  wVk As Integer
  wScan As Integer
  dwFlags As Long
  time As Long
  dwExtraInfo As LongPtr          ' Long on 32-bit
End Type

Private Type INPUT_TYPE
  dwType As Long
  ki As KEYBDINPUT
  ' Padding is often needed for correct structure size on 64-bit
  extra As Currency               ' or use a byte array / careful sizing
End Type

Private Sub Class_Terminate()
  Set UIAElement = Nothing
  Set ParentLocator = Nothing
  Set ParentElement = Nothing
End Sub

Private Sub AutoFindElement()
  If UIAElement Is Nothing Then
    ParentLocator.Find 10
  End If
End Sub

Public Function AriaRole() As String
  AriaRole = GetProperty(UIAProperties.AriaRole)
End Function

Public Function AriaProperties() As String
  AriaProperties = GetProperty(UIAProperties.AriaProperties)
End Function

'Public Function SureSelect() As pElement
'  Dim Delay As Integer
'  Delay = 10
'  While Me.IsNotSelected
'    Me.Click
'    snooze Delay
'If Me.IsNotSelected And Delay < 100 Then
''Stop
'  Delay = Delay + 10
'End If
'
'  Wend
'  Set SureSelect = Me
'End Function

Public Function Click(Optional PlusShift As Boolean = False, Optional PlusCtrl As Boolean = False, Optional ShiftAnchorElement As pElement) As pElement

  Dim InitialSelectionState As Boolean
  Dim ExpectedExpandedState As Integer
  
  If Me.HasPattern(UIAPatterns.ExpandCollapse) Then
    If Me.IsExpanded Then
      ExpectedExpandedState = ExpandCollapseState.ExpandCollapseState_Collapsed
    Else
      ExpectedExpandedState = ExpandCollapseState.ExpandCollapseState_Expanded
    End If
  Else
    InitialSelectionState = Me.IsSelected
  End If
  
  Actions.Click Me, PlusShift, PlusCtrl, ShiftAnchorElement
    
  If HasPattern(UIAPatterns.SelectionItem) Then
    WaitForPatternState UIAPatterns.SelectionItem, Not InitialSelectionState
    GoTo ExitFunction
  End If
  
  If HasPattern(UIAPatterns.Toggle) Then
    WaitForPatternState UIAPatterns.Toggle, Not InitialSelectionState
    GoTo ExitFunction
  End If
  
  If HasPattern(UIAPatterns.ExpandCollapse) Then
    WaitForPatternState UIAPatterns.ExpandCollapse, ExpectedExpandedState
    GoTo ExitFunction
  End If
  
ExitFunction:
  Set Click = Me

End Function

Public Function CtrlClick() As pElement
  Set CtrlClick = Me.Click(PlusCtrl:=True)
End Function

Public Function ShiftClick(Optional ShiftAnchorElement As pElement) As pElement
  Set ShiftClick = Me.Click(PlusShift:=True, ShiftAnchorElement:=ShiftAnchorElement)
End Function

Public Sub ClickIfEnabled(Optional NonClickItem As pElement)
  If Me.IsEnabled Then
    Me.Click
  Else
    If Not NonClickItem Is Nothing Then
      NonClickItem.Click
    End If
  End If
End Sub

Public Sub RightClick()
  Actions.RightClick Me
End Sub

Public Function ControlType() As Long
  ControlType = GetProperty(UIAProperties.ControlType)
End Function

Public Sub DragAndDrop(ToElement As pElement)
  Actions.DragAndDrop Me, ToElement
End Sub

Public Sub CloseWindow()
  AutoFindElement
  Toaster.Message "Close Window " & Name, Action
  Actions.IsElementReady Me
  If HasProperty(UIAProperties.IsWindowPatternAvailable) Then
    Dim patt As IUIAutomationWindowPattern
    Set patt = GetPattern(UIAPatterns.Window, RaiseError:=True)
    patt.Close
    Set patt = Nothing
    Exit Sub
  End If
End Sub

Public Function IsExpanded() As Boolean
  IsExpanded = CBool(ExpandCollapseState.ExpandCollapseState_Expanded)
End Function

Public Function ExpandedCollapsedState() As Integer
  Dim Pattern As IUIAutomationExpandCollapsePattern
  Set Pattern = Me.UIAElement.GetCurrentPattern(UIAPatterns.ExpandCollapse)
  ExpandedCollapsedState = Pattern.CurrentExpandCollapseState
End Function

'Tools > References > OLE Automation needed for IUnknown type
Public Function GetPattern(PatternId As Long, Optional RaiseError As Boolean) As IUnknown
  AutoFindElement
  On Error Resume Next
  Set GetPattern = UIAElement.GetCurrentPattern(PatternId)
  On Error GoTo 0
  If RaiseError Then
    If GetPattern Is Nothing Then
      ErrorLogging.LogError Errors.PatternFailedForElement, "Pattern failed for element: " & GivenName
      Exit Function
    End If
  End If
End Function

Public Function GetProperty(PropertyId As Long) As Variant
  AutoFindElement
  On Error Resume Next
  GetProperty = UIAElement.GetCurrentPropertyValue(PropertyId)
  On Error GoTo 0
End Function

Public Function FindDescendant(NameOrId As String, Optional TargetControlTypeOrAriaRole As Variant) As pElement
  If (Me.UIAElement.CurrentControlType = UIAControlTypeIDs.List) Or _
     (Me.UIAElement.CurrentControlType = UIAControlTypeIDs.ComboBox) Then
     TargetControlTypeOrAriaRole = UIAControlTypeIDs.ListItem
  End If
  Dim Descendants As IUIAutomationElementArray
  Dim Descendant As IUIAutomationElement
  Set Descendants = Me.UIAElement.FindAll(TreeScope.Descendants, UIA.CreateTrueCondition)
  If Descendants.Length = 0 Then
    ' No children, just exit
  Else
    Dim Counter As Long
    For Counter = 0 To Descendants.Length - Counter
      Set Descendant = Descendants.GetElement(Counter)
      If Descendant.CurrentName = NameOrId Then
        If Not IsMissing(TargetControlTypeOrAriaRole) Then
          If Descendant.CurrentControlType = TargetControlTypeOrAriaRole Then
            Set FindDescendant = Factory.GetNewElement(NameOrId, Descendant, Me)
            Exit Function
          End If
        Else
          Set FindDescendant = Factory.GetNewElement(NameOrId, Descendant)
          Exit Function
        End If
      End If
    Next Counter
  End If
  Set Descendant = Nothing
  Set Descendants = Nothing
End Function

Public Function FindSibling(NameOrId As String, Optional TargetControlTypeOrAriaRole As Variant = "") As pElement
  
  Dim SiblingUIAElement As IUIAutomationElement
  
  If (Me.UIAElement.CurrentControlType = UIAControlTypeIDs.Text) And _
     ((TargetControlTypeOrAriaRole = UIAControlTypeIDs.RadioButton) Or _
      (TargetControlTypeOrAriaRole = UIAControlTypeIDs.CheckBox)) Then
    'Radio Buttons/Checkboxes preceed their labels
    'Look for previous sibling
    Set TreeWalker = UIA.ControlViewWalker
    Set SiblingUIAElement = TreeWalker.GetPreviousSiblingElement(Me.UIAElement)
    While Not SiblingUIAElement Is Nothing
      If (SiblingUIAElement.CurrentAutomationId = NameOrId) Then
        Set FindSibling = Factory.GetNewElement(NameOrId, SiblingUIAElement)
        Exit Function
      End If
      Set SiblingUIAElement = TreeWalker.GetPreviousSiblingElement(SiblingUIAElement)
    Wend
    'Prevent looking for following items!
    Exit Function
  End If
  
  'Else Look for next sibling
  Set TreeWalker = UIA.ControlViewWalker
  Set SiblingUIAElement = TreeWalker.GetNextSiblingElement(Me.UIAElement)
  While Not SiblingUIAElement Is Nothing
    If TargetControlTypeOrAriaRole = "" Then
      If (SiblingUIAElement.CurrentName = NameOrId) Then
        Set FindSibling = Factory.GetNewElement(NameOrId, SiblingUIAElement)
        Exit Function
      End If
    Else
      If (SiblingUIAElement.CurrentName = NameOrId) And _
         ((SiblingUIAElement.CurrentControlType = TargetControlTypeOrAriaRole) Or _
          (SiblingUIAElement.CurrentAriaRole = TargetControlTypeOrAriaRole)) Then
        Set FindSibling = Factory.GetNewElement(NameOrId, SiblingUIAElement)
        Exit Function
      End If
    End If
    Set SiblingUIAElement = TreeWalker.GetNextSiblingElement(SiblingUIAElement)
  Wend

End Function

Public Function FindLabelledElement() As pElement
'Labelled element must be a sibling
  Dim SiblingUIAElement As IUIAutomationElement
  Dim FirstSiblingUIAElement As IUIAutomationElement
  Dim LabelRuntimeID As String
  Dim CurrentRuntimeID As String
  Set TreeWalker = UIA.ControlViewWalker
  LabelRuntimeID = pPath.RuntimeIDs.GetElementRuntimeId(Me.UIAElement)
  Set SiblingUIAElement = TreeWalker.GetNextSiblingElement(Me.UIAElement)
    While Not SiblingUIAElement Is Nothing
      If FirstSiblingUIAElement Is Nothing Then
        'Ignore any text elements - these will not be labelled
        If SiblingUIAElement.CurrentControlType <> UIAControlTypeIDs.Text Then
          Set FirstSiblingUIAElement = SiblingUIAElement
        End If
      End If
      If Not SiblingUIAElement.CurrentLabeledBy Is Nothing Then
        CurrentRuntimeID = pPath.RuntimeIDs.GetElementRuntimeId(SiblingUIAElement.CurrentLabeledBy)
        If (CurrentRuntimeID = LabelRuntimeID) Then
          Set FindLabelledElement = Factory.GetNewElement(Me.UIAElement.CurrentName, SiblingUIAElement, Me.ParentElement)
          Exit Function
        End If
      Else
        If SiblingUIAElement.CurrentName = Me.UIAElement.CurrentName Then
          'Some browsers just link labels by the Name!
          Set FindLabelledElement = Factory.GetNewElement(Me.UIAElement.CurrentName, SiblingUIAElement)
          Exit Function
        End If
      End If
      Set SiblingUIAElement = TreeWalker.GetNextSiblingElement(SiblingUIAElement)
    Wend
    'If no clean linking element is found, just return the first item found
    If FindLabelledElement Is Nothing Then
      Set FindLabelledElement = Factory.GetNewElement(Me.UIAElement.CurrentName, FirstSiblingUIAElement, Me.ParentElement)
    End If
End Function

Public Function Text() As String
  AutoFindElement
  Text = UIAPatts.GetTextValue(Me.UIAElement)
End Function

Public Function ToggleState() As Integer
  AutoFindElement
  If HasPattern(UIAPatterns.Toggle) Then
    Dim CurrentElementTogglePattern As IUIAutomationTogglePattern
    Set CurrentElementTogglePattern = GetPattern(UIAPatterns.Toggle, RaiseError:=True)
    ToggleState = CurrentElementTogglePattern.CurrentToggleState
    Set CurrentElementTogglePattern = Nothing
  End If
End Function

Public Function Value() As String
  AutoFindElement
  If HasPattern(UIAPatterns.Value) Then
    Dim CurrentElementValuePattern As IUIAutomationValuePattern
    Set CurrentElementValuePattern = GetPattern(UIAPatterns.Value, RaiseError:=True)
    Value = CurrentElementValuePattern.CurrentValue
    Set CurrentElementValuePattern = Nothing
  End If
End Function

Public Function HasPattern(PatternId As Long, Optional RaiseError As Boolean) As Boolean
  AutoFindElement
  HasPattern = UIAPatts.HasPattern(UIAElement, PatternId)
End Function

Public Function HasProperty(PropertyId As Long, Optional RaiseError As Boolean) As Boolean
  AutoFindElement
  Dim Property As Variant
  On Error Resume Next
  Property = UIAElement.GetCurrentPropertyValue(PropertyId)
  On Error GoTo 0
  HasProperty = Not IsEmpty(Property)
End Function

Public Function IsAlive() As Boolean
  'DON'T USE AUTOFIND ON THIS METHOD - it will create an infiite loop!
  On Error Resume Next
  Dim PID As Long
  PID = UIAElement.CurrentProcessId  'any property access will fail if stale
  IsAlive = (Err.Number = 0) And (PID > 0)
  On Error GoTo 0
End Function

Public Function IsDisabled() As Boolean
  IsDisabled = Not IsEnabled()
End Function

Public Function IsEnabled() As Boolean
  AutoFindElement
  Window.HighlightElement Me.UIAElement
  IsEnabled = GetProperty(UIAProperties.IsEnabled)
  Window.ReleaseHighlighting
End Function

Public Function IsFocussed() As Boolean
  AutoFindElement
  Dim FocusedElement As IUIAutomationElement
  Set FocusedElement = UIA.GetFocusedElement
  If FocusedElement Is Nothing Then
    IsFocussed = False
  Else
    Dim IsMatch As Long
    IsMatch = UIA.CompareElements(Me.UIAElement, FocusedElement)
    IsFocussed = CBool(IsMatch)
  End If
End Function

Public Function IsSelected() As Boolean
  
  AutoFindElement
            
  Window.HighlightElement Me.UIAElement

  ' Method 1: Preferred - Use SelectionItemPattern
  If HasPattern(UIAPatterns.SelectionItem) Then
    Dim SelectionItemPattern As IUIAutomationSelectionItemPattern
    Set SelectionItemPattern = GetPattern(UIAPatterns.SelectionItem)
    IsSelected = SelectionItemPattern.CurrentIsSelected
    Set SelectionItemPattern = Nothing
    GoTo CleanUp
  End If
  
  ' Method 2: try TogglePattern
  If HasPattern(UIAPatterns.Toggle) Then
    Dim TogglePattern As IUIAutomationTogglePattern
    Set TogglePattern = GetPattern(UIAPatterns.Toggle)
    IsSelected = TogglePattern.CurrentToggleState
    Set TogglePattern = Nothing
    GoTo CleanUp
  End If
  
  ' Method 2: Fallback - Direct property (works on many controls)
'  Dim varValue As Variant
  'varValue = el.GetProperty(UIA_IsSelectedPropertyId)   ' Property ID 30079
'varValue = Actions.GetValue(Me.GivenName, Me.UIAElement)
'  If IsBoolean(varValue) Then
'    IsRadioButtonSelected = CBool(varValue)
'  Else
'    IsRadioButtonSelected = False
'  End If
    
CleanUp:
  Window.ReleaseHighlighting
    
End Function

Public Function IsNotSelected() As Boolean
  IsNotSelected = Not IsSelected
End Function

Public Function Name() As String
  Name = GetProperty(UIAProperties.Name)
End Function

'Highlight only when setting the element state - this is an action, but not for gets, which may be part of another action!?
Public Sub SetValue(Value As String)
  AutoFindElement
  Window.HighlightElement UIAElement
  If HasPattern(UIAPatterns.Value) Then
    Dim CurrentElementValuePattern As IUIAutomationValuePattern
    Set CurrentElementValuePattern = GetPattern(UIAPatterns.Value, RaiseError:=True)
    On Error Resume Next
    CurrentElementValuePattern.SetValue Value
    Set CurrentElementValuePattern = Nothing
    Dim InputOK As Boolean
    InputOK = (Err.Number = 0)
    If Not InputOK Then
      'Use WinAPI as a fallback option
      UIAElement.SetFocus
      SendUnicodeText Value
    End If
  End If
  Window.ReleaseHighlighting
End Sub

Private Sub SendUnicodeText(ByVal Text As String)
  Dim i As Long, n As Long
  Dim inputs() As INPUT_TYPE
  Dim ch As Long

'    n = Len(text) * 2               ' down + up for each character
  n = Len(Text) 'This stops the text being duplicated!
  ReDim inputs(0 To n - 1)

  For i = 1 To Len(Text)
    ch = AscW(Mid$(Text, i, 1))
        
    ' Key down
    inputs((i - 1) * 2).dwType = INPUT_KEYBOARD
    inputs((i - 1) * 2).ki.wVk = 0
    inputs((i - 1) * 2).ki.wScan = ch
    inputs((i - 1) * 2).ki.dwFlags = KEYEVENTF_UNICODE

    ' Key up
    inputs((i - 1) * 2 + 1).dwType = INPUT_KEYBOARD
    inputs((i - 1) * 2 + 1).ki.wVk = 0
    inputs((i - 1) * 2 + 1).ki.wScan = ch
    inputs((i - 1) * 2 + 1).ki.dwFlags = KEYEVENTF_UNICODE Or KEYEVENTF_KEYUP
    
  Next i

  Call SendInput(n, inputs(0), LenB(inputs(0)))

End Sub

Public Sub WaitForPropertyValue( _
  UIAProperty As UIAProperties, _
  UIAPropertyValue As Variant, _
  Optional TimeoutInSeconds As Integer, _
  Optional WaitForNotValue As Boolean = False)
  AutoFindElement
  WaitForPropertyValueOrPatternState UIAProperty:=UIAProperty, UIAPropertyValue:=UIAPropertyValue, TimeoutInSeconds:=TimeoutInSeconds, WaitForNotValue:=WaitForNotValue
End Sub
  
Public Sub WaitForPatternState( _
  UIAPatternID As UIAPatterns, _
  PatternState As Variant, _
  Optional TimeoutInSeconds As Integer, _
  Optional WaitForNotValue As Boolean = False)
  AutoFindElement
  WaitForPropertyValueOrPatternState UIAPatternID:=UIAPatternID, PatternState:=PatternState, TimeoutInSeconds:=TimeoutInSeconds, WaitForNotValue:=WaitForNotValue
End Sub
  
Private Sub WaitForPropertyValueOrPatternState( _
  Optional UIAProperty As UIAProperties, _
  Optional UIAPropertyValue As Variant, _
  Optional UIAPatternID As UIAPatterns, _
  Optional PatternState As Variant, _
  Optional TimeoutInSeconds As Integer, _
  Optional WaitForNotValue As Boolean = False)

'TODO: Allow a wait for milliseconds?
  
  'Calculate the end time
  Dim EndTime As Date
  EndTime = DateAdd("s", TimeoutInSeconds, Now)
  
  'Loop until element(s) found or timed out
  Dim PropertyValuePatternStateFound As Boolean
  PropertyValuePatternStateFound = False
  
  Dim PassedEndTime As Boolean
  PassedEndTime = False
  
  Dim CurrentPropertyValue As Variant
  While Not (PropertyValuePatternStateFound Or PassedEndTime)
  
'TODO: REUSE GET ELEMENT PROPERTIES etc HERE!! See UIAPatterns.Value
    If UIAProperty <> 0 Then
      CurrentPropertyValue = GetProperty(UIAProperty)
      PropertyValuePatternStateFound = (CurrentPropertyValue = UIAPropertyValue)
      If WaitForNotValue Then
        PropertyValuePatternStateFound = Not PropertyValuePatternStateFound
      End If
    Else
      Select Case UIAPatternID
        Case UIAPatterns.ExpandCollapse
          Dim ExpandCollapsePattern As IUIAutomationExpandCollapsePattern
          Set ExpandCollapsePattern = GetPattern(UIAPatterns.ExpandCollapse)
          PropertyValuePatternStateFound = (ExpandCollapsePattern.CurrentExpandCollapseState = PatternState)
        Case UIAPatterns.SelectionItem
          Dim SelectionItemPattern As IUIAutomationSelectionItemPattern
          Set SelectionItemPattern = GetPattern(UIAPatterns.SelectionItem)
          Select Case PatternState
            Case "CurrentIsSelected", True
              PropertyValuePatternStateFound = (SelectionItemPattern.CurrentIsSelected = 1)
            Case "CurrentIsNotSelected", False
              PropertyValuePatternStateFound = (SelectionItemPattern.CurrentIsSelected = 0)
          End Select
          Set SelectionItemPattern = Nothing
        Case UIAPatterns.Toggle
          Dim TogglePattern As IUIAutomationTogglePattern
          Set TogglePattern = GetPattern(UIAPatterns.Toggle)
          Select Case PatternState
            Case "CurrentToggleStateOn"
              PropertyValuePatternStateFound = (TogglePattern.CurrentToggleState = 1)
            Case "CurrentToggleStateOff"
              PropertyValuePatternStateFound = (TogglePattern.CurrentToggleState = 0)
          End Select
          Set TogglePattern = Nothing
        Case UIAPatterns.Value
          PropertyValuePatternStateFound = (Me.Value = PatternState)
        Case Else
          MsgBox "PatternID not handled: " & UIAPatternID
      End Select
      If WaitForNotValue Then
        PropertyValuePatternStateFound = Not PropertyValuePatternStateFound
      End If
    End If
    
    If Not PropertyValuePatternStateFound Then
      PassedEndTime = (Now > EndTime)
      If Not PassedEndTime Then
        Snooze 10
      End If
    End If
  Wend

End Sub
