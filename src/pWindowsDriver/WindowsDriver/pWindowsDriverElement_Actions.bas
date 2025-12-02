VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pWindowsDriverElement_Actions"
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
' =======================================================================Option Explicit
Option Explicit

#If VBA7 Then
    
  Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hwnd As LongPtr, ByVal wMsg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
        
  Private Declare PtrSafe Function WindowFromPoint Lib "user32" ( _
    ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr
       
  Private Declare PtrSafe Function GetCursorPos Lib "user32" ( _
    lpPoint As POINTAPI) As Long
        
  Private Declare PtrSafe Function SetCursorPos Lib "user32" ( _
    ByVal x As Long, ByVal y As Long) As Long
        
  Private Declare PtrSafe Sub mouse_event Lib "user32" ( _
    ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, _
    ByVal dwData As Long, ByVal dwExtraInfo As LongPtr)

  Private Type POINTAPI
    x As Long
    y As Long
  End Type
    
#Else
    
  Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
        
  Private Declare Function WindowFromPoint Lib "user32" ( _
    ByVal xPoint As Long, ByVal yPoint As Long) As Long
        
  Private Declare Function GetCursorPos Lib "user32" ( _
    lpPoint As POINTAPI) As Long
        
  Private Declare Function SetCursorPos Lib "user32" ( _
    ByVal x As Long, ByVal y As Long) As Long
        
  Private Declare Sub mouse_event Lib "user32" ( _
    ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, _
    ByVal dwData As Long, ByVal dwExtraInfo As Long)

  Private Type POINTAPI
    x As Long
    y As Long
  End Type

#End If

Private Const WM_RBUTTONDOWN   As Long = &H204
Private Const WM_RBUTTONUP     As Long = &H205
Private Const MK_RBUTTON       As Long = &H2
Private Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Private Const MOUSEEVENTF_RIGHTUP   As Long = &H10

Private CurrentElement As UIAutomationClient.IUIAutomationElement

Public Sub Initialise(element As UIAutomationClient.IUIAutomationElement)
  Set CurrentElement = element
End Sub

'******************
'* Exposed Actions
'******************
Public Sub Click()
  Click_LegacyPattern
End Sub

Public Sub RightClick()
'TODO - Pass an element's xpath to check which worked!
'  Click_LegacyPattern
'  Click_InvokePattern
  MouseClicks WM_RBUTTONDOWN, MK_RBUTTON, WM_RBUTTONUP, 0, MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP
End Sub

Private Sub Click_LegacyPattern()
  Dim legacyPattern As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern
  ' Method 1: LegacyIAccessible.DoDefaultAction() ? this is THE supported way in Excel
  Set legacyPattern = CurrentElement.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
  If Not legacyPattern Is Nothing Then
    legacyPattern.DoDefaultAction
  End If
End Sub

Private Sub Click_InvokePattern()
  ' Method 2: Fallback ? InvokePattern (works on some elements)
  Dim invokePattern As IUIAutomationInvokePattern
  Set invokePattern = CurrentElement.GetCurrentPattern(UIA_InvokePatternId)
  If Not invokePattern Is Nothing Then
    invokePattern.Invoke
  End If
End Sub

Private Sub MouseClicks( _
  Action1 As Long, Action2 As Long, Action3 As Long, Action4 As Long, _
  Event1 As Long, Event2 As Long)
      
  Dim clickablePoint As tagPOINT
  Dim hasPoint As Boolean
  hasPoint = CurrentElement.GetClickablePoint(clickablePoint)
    
  Dim pt As POINTAPI
  If hasPoint Then
    pt.x = clickablePoint.x
    pt.y = clickablePoint.y
  Else
    ' Fallback: center of bounding rectangle
    Dim rect As tagRECT
    rect = CurrentElement.CurrentBoundingRectangle
    pt.x = rect.Left + (rect.Right - rect.Left) \ 2
    pt.y = rect.Top + (rect.bottom - rect.Top) \ 2
  End If
    
  Dim hwnd As LongPtr
  hwnd = WindowFromPoint(pt.x, pt.y)
    
  If hwnd <> 0 Then
  
    Logger.InternalDebug "Right-click element action sent via PostMessage (background-safe)"
    Dim lParam As LongPtr
    lParam = (pt.y * &H10000) Or (pt.x And &HFFFF&)
    PostMessage hwnd, Action1, Action2, lParam
    Phosphorus.WindowsProcesses.Snooze 50
    PostMessage hwnd, Action3, Action4, lParam
  
  Else
    
    Logger.InternalDebug "Right-click element action sent via mouse_event"
    SetCursorPos pt.x, pt.y
    Phosphorus.WindowsProcesses.Snooze 50
    mouse_event Event1, 0, 0, 0, 0
    Phosphorus.WindowsProcesses.Snooze 80
    mouse_event Event2, 0, 0, 0, 0
    Debug.Print "Right-click sent via mouse_event"
  
  End If
    
End Sub
