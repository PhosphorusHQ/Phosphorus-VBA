VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Actions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder UIAutomation
'VB_PredeclaredId - see: https://www.vbaplanet.com/attributes.php
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

' =============================================================
' TYPE DEFINITIONS
' =============================================================
Private Type POINTAPI
  x As Long
  y As Long
End Type

' =============================================================
' CONSTANTS - Mouse Messages
' =============================================================
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_LBUTTONDBLCLK  As Long = &H203
Private Const WM_RBUTTONDOWN    As Long = &H204
Private Const WM_RBUTTONUP      As Long = &H205
Private Const WM_RBUTTONDBLCLK  As Long = &H206
'Private Const MK_RBUTTON       As Long = &H2

Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4

Private Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Private Const MOUSEEVENTF_RIGHTUP   As Long = &H10

Private Const KEYEVENTF_KEYUP = &H2

'Keys
Private Const VK_CONTROL = &H11
Private Const VK_MENU = &H12    ' Alt
Private Const VK_SHIFT = &H10

'' Optional: keyboard messages (often used together)
'Private Const WM_KEYDOWN        As Long = &H100
'Private Const WM_KEYUP          As Long = &H101
'Private Const WM_CHAR           As Long = &H102

' Window style constants
Private Const WS_EX_LAYERED      As Long = &H80000
Private Const WS_EX_TRANSPARENT  As Long = &H20
Private Const WS_EX_TOOLWINDOW   As Long = &H80
Private Const WS_EX_TOPMOST      As Long = &H8
Private Const WS_POPUP           As Long = &H80000000
'Private Const WS_VISIBLE         As Long = &H10000000

'Declarations for mouse clicks
#If VBA7 Then

  ' 64-bit Office

  Private Declare PtrSafe Function SetCursorPos Lib "user32" ( _
    ByVal x As Long, ByVal y As Long) As Long

  Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
      
  Private Declare PtrSafe Sub mouse_event Lib "user32" ( _
    ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, _
    ByVal dwData As Long, ByVal dwExtraInfo As LongPtr)

 Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
        
  Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As Long

  Private Declare PtrSafe Function ScreenToClient Lib "user32" ( _
    ByVal hWnd As LongPtr, lpPoint As POINTAPI) As Long
                  
#Else
  ' 32-bit Office

'  Private Declare Function WindowFromPoint Lib "user32" ( _
'    ByVal xPoint As Long, ByVal yPoint As Long) As Long
'
'  Private Declare Function GetCursorPos Lib "user32" ( _
'    lpPoint As POINTAPI) As Long

  Private Declare Function SetCursorPos Lib "user32" ( _
    ByVal x As Long, ByVal y As Long) As Long

  Private Declare Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    
  Private Declare Sub mouse_event Lib "user32" ( _
    ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, _
    ByVal dwData As Long, ByVal dwExtraInfo As Long)

   Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
        
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
        
    Private Declare Function ScreenToClient Lib "user32" _
        (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
   
#End If

Private Enum MouseClickType
  LeftClick
  RightClick
  DoubleClick
  LeftClickSynchronous
End Enum

'LegacyIAccessiblePattern Select action flags
Const SELFLAG_NONE = &H0 'No change (sometimes used as "clear selection")
Const SELFLAG_TAKEFOCUS = &H1 'Give keyboard focus to this item, almost always combined with others
Const SELFLAG_TAKESELECTION = &H2 'Make this the only selected item (clears all others), Single-select lists, trees, combo boxes
Const SELFLAG_EXTENDSELECTION = &H4 'Extend the selection from anchor to this item (Shift+click behavior, Range selection in multi-select controls
Const SELFLAG_ADDSELECTION = &H8 'Add this item to the current selection (Ctrl+click behavior), Adding individual items to multi-selection
Const SELFLAG_REMOVESELECTION = &H10 'Remove this item from the current selection, Ctrl+click to deselect in multi-select
Const SELFLAG_VALID = &H1F 'Bitmask of all valid flags (0x10)
'Most used flag combinations:
'Select this item only (single-select): TAKEFOCUS + TAKESELECTION
'Select this item + give focus: SELFLAG_TAKEFOCUS
'Add this item to existing selection: SELFLAG_TAKEFOCUS Or SELFLAG_ADDSELECTION
'Remove this item from selection: SELFLAG_TAKEFOCUS Or SELFLAG_REMOVESELECTION
'Extend selection (range): SELFLAG_TAKEFOCUS Or SELFLAG_EXTENDSELECTION
'Select only this item, no focus change: SELFLAG_TAKESELECTION

'**********
'* Clicks *
'**********
Public Sub Click(Element As pElement)
    
  Toaster.Message "Clicking " & Name, Action
    
  IsElementReady Element
  Window.HighlightElement Element.UIAElement
  MoveMouseToElement Element

  If TryInvokePattern(Element) Then GoTo Cleanup
  If TrySelectionItemPatternSelect(Element) Then GoTo Cleanup
  If TryTogglePattern(Element) Then GoTo Cleanup
  If TryLegacyIAccessibleDefaultAction(Element) Then GoTo Cleanup
  If TryLegacyIAccessiblePatternSelect(Element, SELFLAG_TAKEFOCUS + SELFLAG_TAKESELECTION) Then GoTo Cleanup
  ' Try this last as we don't know if it worked!
  If TryMouseClickByMessage(Element, LeftClick) Then GoTo Cleanup
'This works?
  If TryMouseClicksByEvent(Element, LeftClick) Then GoTo Cleanup
  If TryMouseClickByMessage(Element, LeftClickSynchronous) Then GoTo Cleanup

Cleanup:
  Window.ReleaseHighlighting

End Sub

Private Function TryInvokePattern(Element As pElement) As Boolean
  On Error GoTo Finish
  TryInvokePattern = False
  If GetPatternPreChecks(Element, IsInvokePatternAvailable) Then
    Dim Pattern As IUIAutomationInvokePattern
    Set Pattern = Element.GetPattern(UIAPatterns.InvokePattern)
    Pattern.Invoke
    TryInvokePattern = True
  End If
Finish:
  On Error Resume Next
End Function

Private Function GetPatternPreChecks(Element As pElement, IsPatternAvailableProperty As UIAProperties) As Boolean
  GetPatternPreChecks = False
  If Element.HasProperty(UIAProperties.ControlType) Then
    If Element.HasProperty(IsPatternAvailableProperty) Then
      If Element.GetProperty(IsPatternAvailableProperty) Then
        GetPatternPreChecks = True
        Exit Function
      End If
    End If
  End If
End Function

Private Function TrySelectionItemPatternSelect(Element As pElement) As Boolean
  On Error GoTo Finish
  TrySelectionItemPatternSelect = False
  If GetPatternPreChecks2(Element, ListItem, IsSelectionItemPatternAvailable) Then
    Dim Pattern As IUIAutomationSelectionItemPattern
    Set Pattern = Element.GetPattern(UIAPatterns.SelectionItemPattern)
    Pattern.Select
    TrySelectionItemPatternSelect = True
  End If
Finish:
  On Error Resume Next
End Function

Private Function TryTogglePattern(Element As pElement) As Boolean
  On Error GoTo Finish
  TryTogglePattern = False
  If GetPatternPreChecks2(Element, CheckBox, IsTogglePatternAvailable) Then
    Dim InitialToggleState As Integer
    InitialToggleState = Element.GetToggleState()
    Dim Pattern As IUIAutomationTogglePattern
    Set Pattern = Element.GetPattern(UIAPatterns.TogglePattern)
    Pattern.Toggle
    If InitialToggleState = 0 Then
      Element.WaitForPropertyValue UIAProperties.ToggleToggleState, 1
    Else
      Element.WaitForPropertyValue UIAProperties.ToggleToggleState, 0
    End If
    TryTogglePattern = True
  End If
Finish:
  On Error Resume Next
End Function

Private Function TryLegacyIAccessibleDefaultAction(Element As pElement) As Boolean
  On Error GoTo Finish
  TryLegacyIAccessibleDefaultAction = False
  If GetPatternPreChecks2(Element, ListItem, IsLegacyIAccessiblePatternAvailable) Then
    Dim Pattern As IUIAutomationLegacyIAccessiblePattern
    Set Pattern = Element.GetPattern(UIAPatterns.LegacyIAccessiblePattern)
    Pattern.DoDefaultAction
    TryLegacyIAccessibleDefaultAction = True
  End If
Finish:
  On Error Resume Next
End Function

Private Function TryLegacyIAccessiblePatternSelect(Element As pElement, Flags As Integer)
  On Error GoTo Finish
  TryLegacyIAccessiblePatternSelect = False
  If GetPatternPreChecks2(Element, ListItem, IsLegacyIAccessiblePatternAvailable) Then
    Dim Pattern As IUIAutomationLegacyIAccessiblePattern
    Set Pattern = Element.GetPattern(UIAPatterns.LegacyIAccessiblePattern)
    Pattern.Select Flags
    TryLegacyIAccessiblePatternSelect = True
  End If
Finish:
  On Error Resume Next
End Function

Private Function GetPatternPreChecks2(Element As pElement, ControlType As UIAControlTypeIDs, IsPatternAvailableProperty As UIAProperties) As Boolean
  GetPatternPreChecks2 = False
  If Element.HasProperty(UIAProperties.ControlType) Then
    If Element.GetProperty(UIAProperties.ControlType) = ControlType Then
      If Element.HasProperty(IsPatternAvailableProperty) Then
        GetPatternPreChecks2 = True
        Exit Function
      End If
    End If
  End If
End Function

Private Function TryMouseClickByMessage(Element As pElement, ClickType As MouseClickType) As Boolean
  On Error GoTo Finish
  TryMouseClickByMessage = False
  MouseClickByMessage Element, ClickType
  TryMouseClickByMessage = True
Finish:
  On Error Resume Next
End Function

Private Sub MouseClickByMessage(Element As pElement, ClickType As MouseClickType)
'PostMessage is asynchronous, SendMessage is synchronous (for special synchronous cases)

  If Element.UIAElement Is Nothing Then Exit Sub
    
  Dim rect As Variant
  rect = Element.UIAElement.GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId)
  
  If Not IsArray(rect) Or UBound(rect) < 3 Then Exit Sub
    
  Dim centerX As Long: centerX = CLng(rect(0) + rect(2) \ 2)
  Dim centerY As Long: centerY = CLng(rect(1) + rect(3) \ 2)
    
  Dim hWnd As LongPtr
  hWnd = Element.UIAElement.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
    
  Dim pt As POINTAPI
  pt.x = centerX
  pt.y = centerY
  ScreenToClient hWnd, pt
    
  'Make LPARAM from x,y coordinates
  Dim MakeLParam As LongPtr
  MakeLParam = (CLngPtr(pt.y) * &H10000) Or (pt.x And &HFFFF&)
    
  'Use PostMessage by default - asynchronous, safer for UI Automation
  Select Case ClickType
    Case MouseClickType.LeftClick
      PostMessage hWnd, WM_LBUTTONDOWN, 1, MakeLParam
      WindowsProcesses.Snooze 35
      PostMessage hWnd, WM_LBUTTONUP, 0, MakeLParam
    Case MouseClickType.RightClick
      PostMessage hWnd, WM_RBUTTONDOWN, 0, MakeLParam
      WindowsProcesses.Snooze 40
      PostMessage hWnd, WM_RBUTTONUP, 0, MakeLParam
    Case MouseClickType.DoubleClick
      PostMessage hWnd, WM_LBUTTONDBLCLK, 1, MakeLParam
      WindowsProcesses.Snooze 50
      PostMessage hWnd, WM_LBUTTONUP, 0, MakeLParam
    Case MouseClickType.LeftClickSynchronous
      SendMessage hWnd, WM_LBUTTONDOWN, 1, MakeLParam
      WindowsProcesses.Snooze 500
      SendMessage hWnd, WM_LBUTTONUP, 0, MakeLParam
  End Select
  
End Sub

Private Function TryMouseClicksByEvent(Element As pElement, ClickType As MouseClickType) As Boolean
  Element.UIAElement.SetFocus
  On Error GoTo Finish
  TryMouseClicksByEvent = False
  Select Case ClickType
    Case MouseClickType.LeftClick
      MouseClicksByEvent Element, MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP
    Case MouseClickType.RightClick
      MouseClicksByEvent Element, MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP
  End Select
  TryMouseClicksByEvent = True
Finish:
  On Error Resume Next
End Function

Private Sub MouseClicksByEvent(Element As pElement, Event1 As Long, Event2 As Long)

  Dim clickablePoint As tagPOINT
  Dim hasPoint As Boolean
  hasPoint = Element.UIAElement.GetClickablePoint(clickablePoint)

  Dim pt As POINTAPI
  If hasPoint Then
    pt.x = clickablePoint.x
    pt.y = clickablePoint.y
  Else
    'Fallback: center of bounding rectangle
    Dim rect As tagRECT
    rect = Element.UIAElement.CurrentBoundingRectangle
    pt.x = rect.Left + (rect.Right - rect.Left) \ 2
    pt.y = rect.Top + (rect.Bottom - rect.Top) \ 2
  End If

  SetCursorPos pt.x, pt.y
  WindowsProcesses.Snooze 50
  mouse_event Event1, 0, 0, 0, 0
  WindowsProcesses.Snooze 80
  mouse_event Event2, 0, 0, 0, 0

End Sub

'***************
'* Drag & Drop *
'***************
Public Sub DragAndDrop( _
  SourceElement As pElement, _
  TargetElement As pElement, _
  Optional holdTimeMs As Long = 400, _
  Optional speed As Long = 500, _
  Optional ctrlKey As Boolean = False, _
  Optional shiftKey As Boolean = False, _
  Optional altKey As Boolean = False)
'DragAndDropWithVisualFeedback
'ctrlKey:=True makes a copy

  Toaster.Message "Drag And Drop " & SourceElement.GivenName, Action
  
  If SourceElement.UIAElement Is Nothing Or TargetElement.UIAElement Is Nothing Then Exit Sub
      
  IsElementReady SourceElement
  IsElementReady TargetElement

  ' Highlight source briefly
'  Window.HighlightElement SourceElement, 5, &HFF00, 300      ' Green highlight on source
  Window.HighlightElement SourceElement.UIAElement
  Window.ReleaseHighlighting
    
  ' Move mouse to source
  MoveMouseToElement SourceElement
  WindowsProcesses.Snooze 250
    
  ' Press modifier keys if requested
  If ctrlKey Then keybd_event VK_CONTROL, 0, 0, 0
  If shiftKey Then keybd_event VK_SHIFT, 0, 0, 0
  If altKey Then keybd_event VK_MENU, 0, 0, 0
    
  ' Start drag
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  WindowsProcesses.Snooze holdTimeMs
    
  ' Get target center
  Dim tgtRect As Variant
  tgtRect = TargetElement.UIAElement.GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId)
  Dim tgtX As Long: tgtX = CLng(tgtRect(0) + tgtRect(2) \ 2)
  Dim tgtY As Long: tgtY = CLng(tgtRect(1) + tgtRect(3) \ 2)

  ' Get source center for smooth movement
  Dim srcRect As Variant
  srcRect = SourceElement.UIAElement.GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId)
  Dim srcX As Long: srcX = CLng(srcRect(0) + srcRect(2) \ 2)
  Dim srcY As Long: srcY = CLng(srcRect(1) + srcRect(3) \ 2)
    
  Dim dx As Long: dx = tgtX - srcX
  Dim dy As Long: dy = tgtY - srcY
  Dim steps As Long: steps = Application.Max(25, Abs(dx) \ speed + Abs(dy) \ speed)
    
  ' Create moving highlight
  Dim HighlightHwnd As LongPtr
  HighlightHwnd = Window.CreateDragHighlight(SourceElement.UIAElement, srcX, srcY, 80, 80)
    
  ' Smooth drag with moving highlight
  Dim i As Long
  For i = 1 To steps
    Dim currX As Long: currX = srcX + CLng(dx * i / steps)
    Dim currY As Long: currY = srcY + CLng(dy * i / steps)
    SetCursorPos currX, currY
    Window.UpdateDragHighlight HighlightHwnd, currX - 40, currY - 40   ' Center highlight on cursor
    WindowsProcesses.Snooze 12
  Next i
    
  SetCursorPos tgtX, tgtY
  WindowsProcesses.Snooze 50
    
  ' Drop
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
  ' Release modifiers
  If ctrlKey Then keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
  If shiftKey Then keybd_event VK_SHIFT, 0, KEYEVENTF_KEYUP, 0
  If altKey Then keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
    
  ' Clean up highlight
  If HighlightHwnd <> 0 Then Window.Destroy (HighlightHwnd)
    
  WindowsProcesses.Snooze 300   ' Allow drop to process

End Sub

Public Function IsElementReady(Element As pElement) As Boolean
  Dim ret As Boolean
  ret = True
  ret = ret And Element.IsAlive()
  If ret Then
    TryToScrollItemIntoView Element
  Else
    ErrorLogging.LogError Errors.ElementIsNotAlive, "Element '" & Element.GivenName & "' is not alive!"
  End If
  IsElementReady = ret
End Function

Private Sub MoveMouseToElement(Element As pElement)
  If Element.UIAElement Is Nothing Then Exit Sub
  Dim rect As Variant
  rect = Element.UIAElement.GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId)
  If Not IsArray(rect) Or UBound(rect) < 3 Then Exit Sub
  Dim cx As Long: cx = CLng(rect(0) + rect(2) \ 2)
  Dim cy As Long: cy = CLng(rect(1) + rect(3) \ 2)
  SetCursorPos cx, cy
End Sub

Public Sub TryToScrollItemIntoView(Element As pElement)
  If Element.HasProperty(UIAProperties.IsScrollItemPatternAvailable) Then
    If Element.HasProperty(UIAProperties.IsOffscreen) Then
      If Element.GetProperty(UIAProperties.IsOffscreen) Then
        Dim patt As IUIAutomationScrollItemPattern
        Set patt = Element.GetPattern(UIAPatterns.ScrollItemPattern, RaiseError:=True)
        patt.ScrollIntoView
      End If
    End If
  End If
End Sub

