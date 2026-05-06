VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmToaster 
   Caption         =   "Phosphorus"
   ClientHeight    =   3996
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4224
   OleObjectBlob   =   "frmToaster.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmToaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder Toaster
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

' API Declarations
#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal hwndInsertAfter As LongPtr, _
         ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
         ByVal uFlags As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, _
         ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
         ByVal uFlags As Long) As Long
#End If

' This event fires AFTER the form is fully visible
Private Sub UserForm_Activate()
  With Window
    .MakeAlwaysOnTopByCaptionAndClassName Me.Caption, "ThunderDFrame"
    .HideTopLevelFormTitleBarAndBorder Me.Caption
    '.MakeWindowTransparent Me.Caption, Me.BackColor
  End With
End Sub

' =============================================================
' Add Line with Dynamic Background Color
' =============================================================
Public Sub AddLine(Text As String, Optional NotificationType As ToasterNotificationType = Info)

  LastActivity = Now   ' Reset inactivity timer
    
  Dim timestamp As String
  timestamp = Format(Now, "hh:nn:ss")
    
  Dim Prefix As String
    
  Select Case NotificationType
    Case Finding
      Prefix = ChrW(10165) & " "
    Case Action
      Prefix = ChrW(9758) & " "
    Case Success
      Prefix = ChrW(10003) & " "
    Case Warning
      Prefix = ChrW(9888) & " "
    Case Error
      Prefix = ChrW(10060) & " "
    Case Info
      Prefix = ChrW(9432) & " "
    Case Else
      Prefix = ChrW(8226) & " "
  End Select
  With txtLog
    .SelStart = Len(.Text)
    .SelText = timestamp & "  " & Prefix & Text & vbCrLf
      
    ' Color the text
    .SelStart = Len(.Text) - Len(Text & Prefix & timestamp & "  ") - 2
    .SelLength = Len(Text & Prefix & timestamp & "  ")
        
    ' Auto-scroll to bottom
    .SelStart = Len(.Text)
    .SelLength = 0

    ' === SMART TRIMMING: Keep only last 80 lines ===
    Const MAX_LINES As Integer = 16 ' Use fewer less that fills one screen!
    Dim lineCount As Integer
    lineCount = UBound(Split(.Text, vbCrLf))
    If lineCount > MAX_LINES Then
      Dim lines() As String
      lines = Split(.Text, vbCrLf)
      ' Rebuild text keeping only the last MAX_LINES
      Dim newText As String
      Dim i As Integer
      For i = lineCount - MAX_LINES + 1 To lineCount
        If i > 0 Then
          If lines(i) <> "" Then
            newText = newText & lines(i) & vbCrLf
          End If
         End If
      Next i
      'Failsafe - de-duplicate LfCr's
'      newText = VBA.Strings.Replace(newText, vbLf, "")
'      newText = VBA.Strings.Replace(newText, vbCrLf & vbCrLf, vbCrLf)
'      newText = VBA.Strings.Replace(newText, vbCrLf & vbCrLf, vbCrLf)
'      newText = VBA.Strings.Replace(newText, vbCrLf & vbCrLf, vbCrLf)
      .Text = newText
      .SelStart = Len(.Text)
      .SelLength = 0
             ' Force update with multiple DoEvents
        .SetFocus
        DoEvents
        DoEvents
        Application.ScreenUpdating = True
        DoEvents
    End If
  
  End With
  
  DoEvents

End Sub

Public Sub PopUp()
  Me.Height = 0
  Me.Show vbModeless
  Pop True
  'Start inactivity timer
  Toaster.LastActivity = Now
  Application.OnTime Now + TimeValue("00:00:05"), "Toaster.CheckInactivity", , True
End Sub

Public Sub PopDown()
  Pop False
  Me.Hide
End Sub

Private Sub Pop(Up As Boolean)
  Dim StepHeight As Integer
  StepHeight = 1
  If Not Up Then
    StepHeight = -StepHeight
  End If
  Dim i As Integer
  Dim ToHeight As Integer
  For i = 1 To 187
    Me.Height = Me.Height + StepHeight
    Toaster.PositionOnPrimaryMonitorBottomRight
  Next i
End Sub

