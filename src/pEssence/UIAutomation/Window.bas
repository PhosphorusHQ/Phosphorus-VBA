VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Window"
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

#If VBA7 Then
    Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
        ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, _
        ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
        ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal Hdc As LongPtr) As Long
    Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
    Private Declare PtrSafe Function FillRect Lib "user32" (ByVal Hdc As LongPtr, lpRect As RECT, ByVal HBrush As LongPtr) As Long
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
#Else
  Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
    ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, _
    ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long

  Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
  Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
  Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
  Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
  Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
#End If

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

' Window style constants
Private Const WS_EX_LAYERED      As Long = &H80000
Private Const WS_EX_TRANSPARENT  As Long = &H20
Private Const WS_EX_TOOLWINDOW   As Long = &H80
Private Const WS_EX_TOPMOST      As Long = &H8
Private Const WS_POPUP           As Long = &H80000000
Private Const WS_VISIBLE         As Long = &H10000000

Private Const SW_SHOWNA          As Long = 8

Private Const LWA_ALPHA          As Long = &H2

'Highlight element overlay
Private HwndOverlay As LongPtr
Public HighlightElements As Boolean

Public Sub WaitForInteractionState( _
  ElementName As String, _
  CurrentElement As IUIAutomationElement, _
  UIAPropertyValue As pEssence.UIAWindowInteractionStates, _
  Optional TimeoutInMilliseconds As Long)
  
  Actions.WaitForPropertyValue _
    ElementName, _
    CurrentElement, _
    UIAProperties.WindowWindowInteractionState, _
    UIAPropertyValue
  
End Sub

Public Sub CloseWindow(Name As String, Ele As IUIAutomationElement)
  Actions.IsElementReady Name, Ele
  If UIAProps.GetProperty(Ele, UIAProperties.ControlType) = UIAControlTypeIDs.Window Then
    If UIAProps.HasProperty(Name, Ele, UIAProperties.IsWindowPatternAvailable) Then
      Dim patt As IUIAutomationWindowPattern
      Set patt = UIAPatts.GetPattern(Name, Ele, UIA_PatternIds.UIA_WindowPatternId, RaiseError:=True)
      patt.Close
      Exit Sub
    End If
  End If
End Sub

' =============================================================
' Draw red rectangle around element ? wait ? remove
' =============================================================

Public Sub HighlightElement(Ele As IUIAutomationElement, _
  Optional BorderThickness As Long = 6, _
  Optional BorderColor As Long = &HFF0000, _
  Optional DurationMs As Long = 10)
'Default is &HFF0000 Blue (BGR)
'https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/4b7087c6-20fc-4d90-8d83-730a0c6a4aad
'&H808000 421376 Cyan
    
  On Error GoTo ErrHandler
    
  If Not HighlightElements Then
    Exit Sub
  End If
    
  ' Get bounding rect
  Dim RectArray As Variant
  RectArray = Ele.GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId)
    
  If Not IsArray(RectArray) Or UBound(RectArray) < 3 Then
    Debug.Print "Invalid bounding rectangle"
    Exit Sub
  End If
    
  Dim Left   As Long: Left = CLng(RectArray(0)) - BorderThickness
  Dim Top    As Long: Top = CLng(RectArray(1)) - BorderThickness
  Dim width  As Long: width = CLng(RectArray(2)) + BorderThickness * 2
  Dim height As Long: height = CLng(RectArray(3)) + BorderThickness * 2
    
  ' Clamp to prevent off-screen creation issues
  If Left < 0 Then Left = 0
  If Top < 0 Then Top = 0
  If width > 5000 Then width = 5000
  If height > 5000 Then height = 5000
    
'  Debug.Print "Highlight border at: " & Left & "," & Top & "  size " & width & "×" & height

  ' Create layered overlay
  HwndOverlay = CreateWindowEx( _
    WS_EX_LAYERED Or WS_EX_TRANSPARENT Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST, _
    "Static", "HighlightOverlay", _
    WS_POPUP Or WS_VISIBLE, _
    Left, Top, width, height, _
    0, 0, 0, 0)
    
  If HwndOverlay = 0 Then
    Debug.Print "CreateWindowEx failed: " & Err.LastDllError
    Exit Sub
  End If

  If BorderColor = &HFF0000 Then
    ' Set borderColour to fully opaque for Actions (default colour)
    SetLayeredWindowAttributes HwndOverlay, 0, 225, LWA_ALPHA  ' fully opaque
  Else
    ' Set borderColour to almost transparent for other events, e.g. Finds (non-default colour)
    SetLayeredWindowAttributes HwndOverlay, 0, 100, LWA_ALPHA
  End If
  ShowWindow HwndOverlay, SW_SHOWNA
  UpdateWindow HwndOverlay
    
  ' === DRAW BORDER (not fill) ===
  Dim Hdc As LongPtr: Hdc = GetDC(HwndOverlay)
  If Hdc = 0 Then GoTo Cleanup
    
  Dim HBrush As LongPtr: HBrush = CreateSolidBrush(BorderColor)
    
  Dim R As RECT
  R.Left = 0: R.Top = 0
  R.Right = width: R.Bottom = height
    
  ' Draw outer border
  FillRect Hdc, R, HBrush
    
  ' Erase inner area (make border only)
  R.Left = BorderThickness: R.Top = BorderThickness
  R.Right = width - BorderThickness: R.Bottom = height - BorderThickness
  Dim HWhiteBrush As LongPtr: HWhiteBrush = CreateSolidBrush(vbWhite)
  FillRect Hdc, R, HWhiteBrush
    
  DeleteObject HBrush
  DeleteObject HWhiteBrush
  ReleaseDC HwndOverlay, Hdc
    
  WindowsProcesses.Snooze DurationMs
    
Cleanup:
  Exit Sub

ErrHandler:
  Debug.Print "Highlight error: " & Err.Number & " - " & Err.Description
  If HwndOverlay <> 0 Then DestroyWindow HwndOverlay
    
End Sub

Public Sub ReleaseHighlighting(Optional DurationMs As Long = 50)
  
  If Not HighlightElements Then
    Exit Sub
  End If

  'Slight delay for better user experience
  WindowsProcesses.Snooze DurationMs
  If HwndOverlay <> 0 Then DestroyWindow HwndOverlay
  
End Sub

