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
' =======================================================================Option Explicit
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
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
    Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hdc As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
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
    left   As Long
    top    As Long
    Right  As Long
    Bottom As Long
End Type

' Window style constants (add once at module level)
Private Const WS_EX_LAYERED      As Long = &H80000
Private Const WS_EX_TRANSPARENT  As Long = &H20
Private Const WS_EX_TOOLWINDOW   As Long = &H80
Private Const WS_EX_TOPMOST      As Long = &H8

Private Const WS_POPUP           As Long = &H80000000
Private Const WS_VISIBLE         As Long = &H10000000
Private Const SW_SHOWNA          As Long = 8

Private Const LWA_ALPHA          As Long = &H2


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
                            Optional borderThickness As Long = 6, _
                            Optional borderColor As Long = &H808000, _
                            Optional durationMs As Long = 500)
'Default is &HFF0000 Blue (BGR)
'https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/4b7087c6-20fc-4d90-8d83-730a0c6a4aad
'&H808000 421376 Cyan
    On Error GoTo ErrHandler
    
    ' Get bounding rect
    Dim rectArray As Variant
    rectArray = Ele.GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId)
    
    If Not IsArray(rectArray) Or UBound(rectArray) < 3 Then
        Debug.Print "Invalid bounding rectangle"
        Exit Sub
    End If
    
    Dim left   As Long: left = CLng(rectArray(0)) - borderThickness
    Dim top    As Long: top = CLng(rectArray(1)) - borderThickness
    Dim width  As Long: width = CLng(rectArray(2)) + borderThickness * 2
    Dim height As Long: height = CLng(rectArray(3)) + borderThickness * 2
    
    ' Clamp to prevent off-screen creation issues
    If left < 0 Then left = 0
    If top < 0 Then top = 0
    If width > 5000 Then width = 5000
    If height > 5000 Then height = 5000
    
    Debug.Print "Highlight border at: " & left & "," & top & "  size " & width & "×" & height
    
    ' Create layered overlay
    Dim hwndOverlay As LongPtr
    hwndOverlay = CreateWindowEx( _
        WS_EX_LAYERED Or WS_EX_TRANSPARENT Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST, _
        "Static", "HighlightOverlay", _
        WS_POPUP Or WS_VISIBLE, _
        left, top, width, height, _
        0, 0, 0, 0)
    
    If hwndOverlay = 0 Then
        Debug.Print "CreateWindowEx failed: " & Err.LastDllError
        Exit Sub
    End If
    
   ' SetLayeredWindowAttributes hwndOverlay, 0, 255, LWA_ALPHA  ' fully opaque for border
    SetLayeredWindowAttributes hwndOverlay, 0, 50, LWA_ALPHA  ' fully opaque for border
    
    ShowWindow hwndOverlay, SW_SHOWNA
    UpdateWindow hwndOverlay
    
    ' === DRAW RED BORDER (not fill) ===
    Dim hdc As LongPtr: hdc = GetDC(hwndOverlay)
    If hdc = 0 Then GoTo Cleanup
    
    Dim hBrush As LongPtr: hBrush = CreateSolidBrush(borderColor)
    
    Dim r As RECT
    r.left = 0: r.top = 0
    r.Right = width: r.Bottom = height
    
    ' Draw outer border
    FillRect hdc, r, hBrush
    
    ' Erase inner area (make border only)
    r.left = borderThickness: r.top = borderThickness
    r.Right = width - borderThickness: r.Bottom = height - borderThickness
    Dim hWhiteBrush As LongPtr: hWhiteBrush = CreateSolidBrush(vbWhite)
    FillRect hdc, r, hWhiteBrush
    
    DeleteObject hBrush
    DeleteObject hWhiteBrush
    ReleaseDC hwndOverlay, hdc
    
'    Sleep durationMs
    WindowsProcesses.Snooze durationMs
Cleanup:
    If hwndOverlay <> 0 Then DestroyWindow hwndOverlay
    Exit Sub

ErrHandler:
    Debug.Print "Highlight error: " & Err.Number & " - " & Err.Description
    If hwndOverlay <> 0 Then DestroyWindow hwndOverlay
    
End Sub
