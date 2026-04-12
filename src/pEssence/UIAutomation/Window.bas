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
  Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
  Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
  Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
  Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
  Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
  Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hdc As LongPtr, lpRect As rect, ByVal hBrush As LongPtr) As Long
  Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
  Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
  Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
  Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, _
    ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
  Private Declare PtrSafe Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" ( _
    ByVal hWnd As LongPtr, _
    ByVal nIndex As LongPtr) As LongPtr
  Private Declare PtrSafe Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" ( _
    ByVal hWnd As LongPtr, _
    ByVal nIndex As LongPtr, _
    ByVal dwNewLong As LongPtr) As LongPtr
  Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
    ByVal hWnd As LongPtr) As LongPtr
#Else
  Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
    ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, _
    ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
  Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
  Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
  Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
  Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
  Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As Long, ByVal hWndChildAfter As Long, _
    ByVal lpszClass As String, ByVal lpszWindow As String) As Long
  Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long
  Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
  Private Declare Function DrawMenuBar Lib "user32" ( _
    ByVal hWnd As Long) As Long
#End If

    
Private Type rect
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

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW As Long = &H40

Private Const SW_SHOWNA          As Long = 8

'Constants for title bar
Private Const GWL_STYLE As Long = (-16)           'The offset of a window's style
Private Const GWL_EXSTYLE As Long = (-20)         'The offset of a window's extended style
Private Const WS_CAPTION As Long = &HC00000       'Style to add a titlebar
Private Const WS_EX_DLGMODALFRAME As Long = &H1   'Controls if the window has an icon

'Constants for transparency
Private Const LWA_COLORKEY = &H1                  'Chroma key for fading a certain color on your Form
Private Const LWA_ALPHA = &H2                     'Only needed if you want to fade the entire userform
 
'Highlight element overlay
Private HwndOverlay As LongPtr
Public HighlightElements As Boolean

Public Sub WaitForInteractionState( _
  Element As pElement, _
  UIAPropertyValue As UIAWindowInteractionStates, _
  Optional TimeoutInMilliseconds As Long)
  
  Actions.WaitForPropertyValue _
    Element, _
    UIAProperties.WindowWindowInteractionState, _
    UIAPropertyValue
  
End Sub

Public Sub CloseWindow(Element As pElement)
  Toaster.Message "Close Window " & Name, Action
  Actions.IsElementReady Element
  If UIAProps.GetProperty(Element, UIAProperties.ControlType) = UIAControlTypeIDs.Window Then
    If UIAProps.HasProperty(Element, UIAProperties.IsWindowPatternAvailable) Then
      Dim patt As IUIAutomationWindowPattern
      Set patt = Element.GetPattern(UIA_PatternIds.UIA_WindowPatternId, RaiseError:=True)
      patt.Close
      Exit Sub
    End If
  End If
End Sub

Public Sub Destroy(hWnd As LongPtr)
  DestroyWindow hWnd
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
  Dim Width  As Long: Width = CLng(RectArray(2)) + BorderThickness * 2
  Dim Height As Long: Height = CLng(RectArray(3)) + BorderThickness * 2
    
  ' Clamp to prevent off-screen creation issues
  If Left < 0 Then Left = 0
  If Top < 0 Then Top = 0
  If Width > 5000 Then Width = 5000
  If Height > 5000 Then Height = 5000
    
'  Debug.Print "Highlight border at: " & Left & "," & Top & "  size " & Width & "×" & height

  ' Create layered overlay
  HwndOverlay = CreateWindowEx( _
    WS_EX_LAYERED Or WS_EX_TRANSPARENT Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST, _
    "Static", "HighlightOverlay", _
    WS_POPUP Or WS_VISIBLE, _
    Left, Top, Width, Height, _
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
  Dim hdc As LongPtr: hdc = GetDC(HwndOverlay)
  If hdc = 0 Then GoTo Cleanup
    
  Dim hBrush As LongPtr: hBrush = CreateSolidBrush(BorderColor)
    
  Dim r As rect
  r.Left = 0: r.Top = 0
  r.Right = Width: r.Bottom = Height
    
  ' Draw outer border
  FillRect hdc, r, hBrush
    
  ' Erase inner area (make border only)
  r.Left = BorderThickness: r.Top = BorderThickness
  r.Right = Width - BorderThickness: r.Bottom = Height - BorderThickness
  Dim HWhiteBrush As LongPtr: HWhiteBrush = CreateSolidBrush(vbWhite)
  FillRect hdc, r, HWhiteBrush
    
  DeleteObject hBrush
  DeleteObject HWhiteBrush
  ReleaseDC HwndOverlay, hdc
    
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

' =============================================================
' Create moving highlight rectangle during drag
' =============================================================
Public Function CreateDragHighlight(Ele As IUIAutomationElement, x As Long, y As Long, W As Long, h As Long) As LongPtr

  Dim hWnd As LongPtr
  hWnd = CreateWindowEx( _
    WS_EX_LAYERED Or WS_EX_TRANSPARENT Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST, _
    "Static", "DragHighlight", WS_POPUP, x, y, W, h, 0, 0, 0, 0)
    
  If hWnd = 0 Then Exit Function
    
  SetLayeredWindowAttributes hWnd, 0, 160, LWA_ALPHA   ' Semi-transparent red
    
  Dim hdc As LongPtr: hdc = GetDC(hWnd)

  Dim hBrush As LongPtr: hBrush = CreateSolidBrush(&HFF)      ' Red
'  Dim hBrush As LongPtr: hBrush = CreateSolidBrush(&HFF0000)      ' Red?
  Dim r As rect: r.Left = 0: r.Top = 0: r.Right = W: r.Bottom = h
  FillRect hdc, r, hBrush
  
  DeleteObject hBrush
  ReleaseDC hWnd, hdc
    
  ShowWindow hWnd, SW_SHOWNA
  UpdateWindow hWnd
  CreateDragHighlight = hWnd
    
End Function

Public Sub UpdateDragHighlight(hWnd As LongPtr, x As Long, y As Long)
  If hWnd = 0 Then Exit Sub
  SetWindowPos hWnd, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE
End Sub

Public Sub MakeAlwaysOnTopByCaptionAndClassName(Caption As String, ClassName As String)

  Dim hWnd As LongPtr
    
  ' First try by caption
  hWnd = FindWindow(vbNullString, Caption)
    
  ' If that fails, search for UserForm class ("ThunderDFrame" is the class name for VBA UserForms)
  If hWnd = 0 Then
    hWnd = FindWindowEx(0, 0, ClassName, Caption)
  End If
    
  ' Final fallback: search without caption (first ThunderDFrame)
  If hWnd = 0 Then
    hWnd = FindWindowEx(0, 0, ClassName, vbNullString)
  End If
    
  If hWnd = 0 Then
    Debug.Print "Could not find UserForm window handle"
    Exit Sub
  End If
    
  ' Make it Always On Top
  SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    
End Sub

Public Sub HideTopLevelFormTitleBarAndBorder(Caption As String)
  Dim SourceWindow As LongPtr
  SourceWindow = FindWindow(vbNullString, Caption)
  Dim lngWindow As LongPtr
  'Build window and set window until you remove the caption, title bar and frame around the window
  lngWindow = GetWindowLong(SourceWindow, GWL_STYLE)
  lngWindow = lngWindow And (Not WS_CAPTION)
  SetWindowLong SourceWindow, GWL_STYLE, lngWindow
  lngWindow = GetWindowLong(SourceWindow, GWL_EXSTYLE)
  lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
  SetWindowLong SourceWindow, GWL_EXSTYLE, lngWindow
  DrawMenuBar SourceWindow
End Sub

Public Sub MakeWindowTransparent(Caption As String, Color As Variant)
  'set transparencies on userform
  Dim formhandle As LongPtr
  Dim bytOpacity As Byte
 
  formhandle = FindWindow(vbNullString, Caption)
  If IsMissing(Color) Then Color = vbWhite 'default to vbwhite
  bytOpacity = 100 ' variable keeping opacity setting
 
  SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED
  'The following line makes only a certain color transparent so the
  ' background of the form and any object whose BackColor you've set to match
  ' vbColor (default vbWhite) will be transparent.
'  Me.backColor = Color
  SetLayeredWindowAttributes formhandle, Color, bytOpacity, LWA_COLORKEY

End Sub
