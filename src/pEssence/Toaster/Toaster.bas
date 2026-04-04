Attribute VB_Name = "Toaster"
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
#Else
  Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1

Public Enum ToasterNotificationType
  Error = 0
  Finding
  Action
  Info
  Success
  Warning
End Enum

Public LastActivity As Date
Public Const INACTIVITY_TIMEOUT As Long = 5   ' seconds

Private Toast As frmToaster

Private Function ToasterIsActive() As Boolean
  On Error GoTo ExitFunction
  Dim E As Boolean
  E = Toast.Visible
  ToasterIsActive = True
  Exit Function
ExitFunction:
  ToasterIsActive = False
End Function

Private Function TurnOn() As Boolean
  If Not ToasterIsActive Then
    Set Toast = New frmToaster
  End If
End Function

Public Sub Message(Text As String, Optional NotificationType As ToasterNotificationType = Info)
  TurnOn
  If Not Toast.Visible Then Toast.PopUp
  Toast.AddLine Text, NotificationType
End Sub

' =============================================================
' Auto-hide after inactivity
' =============================================================
Public Sub CheckInactivity()
  'If Not Toast Is Nothing Then
  If ToasterIsActive Then
    If Toast.Visible Then
      PositionOnPrimaryMonitorBottomRight
    End If
    If DateDiff("s", LastActivity, Now) > INACTIVITY_TIMEOUT Then
      If Toast.Visible Then
        Toast.PopDown
      End If
    Else
       Application.OnTime Now + TimeValue("00:00:05"), "Toaster.CheckInactivity", , True
    End If
  End If
End Sub

Private Sub Test()
  Message "Starting ..."
  Message "Starting ...", Error
  Message "Starting ...", Info
  Message "Starting ...", Success
  Message "Starting ...", Warning
End Sub

Public Sub PositionOnPrimaryMonitorBottomRight()

  Dim ScreenWidth As Long
  Dim ScreenHeight As Long
    
  ' Get PRIMARY monitor dimensions (this is reliable)
  ScreenWidth = GetSystemMetrics(SM_CXSCREEN)
  ScreenHeight = GetSystemMetrics(SM_CYSCREEN)

  ' Add small safety margin and ensure form is fully visible
  Dim NewLeft As Long
  Dim NewTop As Long
    
  NewLeft = (ScreenWidth - Toast.Width) / 2 + 75
  NewTop = ScreenHeight - Toast.Height - 570 ' Leave space above taskbar
  
  ' Final safety: make sure it's not off-screen
  If NewLeft < 0 Then NewLeft = 30
  If NewTop < 0 Then NewTop = 30
    
  Toast.Left = NewLeft
  Toast.Top = NewTop

  ' Force Windows to apply the position
  Toast.Repaint
  DoEvents

  Window.MakeWindowTransparent Toast.Caption, Toast.BackColor

End Sub

