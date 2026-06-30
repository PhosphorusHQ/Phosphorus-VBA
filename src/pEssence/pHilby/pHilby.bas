Attribute VB_Name = "pHilby"
'@Folder pHilby
'Build 026.5
'***************************************************************************
'
' Authors:  JKP Application Development Services, info@jkp-ads.com, https://www.jkp-ads.com
'           Peter Thornton, pmbthornton@gmail.com
'
' (c)2013-2023, all rights reserved to the authors
'
' You are free to use and adapt the code in these modules for
' your own purposes and to distribute as part of your overall project.
' However all headers and copyright notices should remain intact
'
' You may not publish the code in these modules, for example on a web site,
' without the explicit consent of the authors
'***************************************************************************

Option Explicit

#If VBA7 Then
  Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As tagPOINT) As Long
  Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
  Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
#Else
  Private Declare Function GetCursorPos Lib "user32" (lpPoint As tagPOINT) As Long
  Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
  Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
#End If

Private UserForm As frmpHilby
Public Const GCSAPPNAME As String = "pHilby - UI Automation Spy"
Private Caption As String

Private TimerID As LongPtr
Private Const TIMER_INTERVAL As Long = 250   ' milliseconds (800ms is a good balance)
Private PreviousTagPoint As tagPOINT
Private PollingExpiryTime As Date

'In Tools, VBAProject properties, set the Conditional compilation argument
'DebugMode = 1 to go into debugmode.

#If DEBUGMODE = 1 Then
  ' PT counters for class Init & Term events
  Public gClsTreeViewInit As Long
  Public gClsTreeViewTerm As Long
  Public gClsNodeInit As Long
  Public gClsNodeTerm As Long
  Public gFormInit As Long
  Public gFormTerm As Long
#End If

Public Sub AddButtonToPhosphorusToolbar()
  
  If PhosphorusToolbar Is Nothing Then
    Phosphorus.Toolbar.CreatePhosphorusToolbar
  End If
  
  'Need to create at least 1 control
  Dim btn As CommandBarButton
  Set btn = PhosphorusToolbar.Controls.Add(Type:=msoControlButton)
  With btn
    .Caption = "Launch pHilby"
    .Style = msoButtonIconAndCaption
    .OnAction = "LaunchpHilby"          ' Your macro name
    .FaceId = 25             ' Optional icon (smiley face example)
    .TooltipText = "Launch pHilby"
  End With

End Sub

Private Sub LaunchpHilby()
  Dim LevelsString As String
  LevelsString = InputBox("How many tree levels would you like to load into pHilby? (1-3 is Recommended). Leave blank for all levels)", "pHilby")
  If LevelsString = "" Then
    If MsgBox("Loading the FULL DESKTOP! This may take some time ... please be patient!", vbExclamation + vbOKCancel, "Phosphorus - pHilby") = vbOK Then
      pHilby.Start
    End If
  Else
    Dim LevelsInteger As Integer
    If IsNumeric(LevelsString) Then
      LevelsInteger = CInt(LevelsString)
      pHilby.Start , LevelsInteger
    Else
      MsgBox "Sorry, the text """ & LevelsString & """ cannot be converted to an integer!", vbExclamation + vbOKCancel, "Phosphorus - pHilby"
    End If
  End If
End Sub

Public Sub Start(Optional RootUIAElement As IUIAutomationElement, Optional MaxNumberOfLevels As Integer = 0, Optional DelayLoadingInSeconds As Integer = 0)

  Application.Cursor = xlWait
  
  'pPath always requires the Logger class
  Phosphorus.Log4PStatic.GetLogger

  Toaster.PopDown
  If RootUIAElement Is Nothing Then
    Set RootUIAElement = Factory.GetRootDesktopElement
  End If
  
  Set UserForm = New frmpHilby
  With UserForm
    .LoadTreeView RootUIAElement, MaxNumberOfLevels:=MaxNumberOfLevels, DelayLoadingInSeconds:=DelayLoadingInSeconds
     Application.Cursor = xlDefault
    .AppName = GCSAPPNAME
    .Show 'Must always be Modal, set in properties!
    Unload UserForm
    Set UserForm = Nothing
  End With

  #If DEBUGMODE = 1 Then
    ClassCounts
  #End If

  Phosphorus.Log4PStatic.CloseLogger

End Sub

#If DEBUGMODE = 1 Then
  Sub ClassCounts()
  ' PT, If making any code modifications it's important to ensure all class instances have been properly terminated.
  '     Classes are counted when created and when destroyed in respective Initialize & Terminate events,
  '     when done the totals should be the same.
    If gClsTreeViewInit <> gClsTreeViewTerm Or _
      gClsNodeInit <> gClsNodeTerm Or _
      gFormInit <> gFormTerm Then
      Debug.Print "clsTreeView", gClsTreeViewInit, gClsTreeViewTerm, gClsTreeViewInit - gClsTreeViewTerm
      Debug.Print "clsNode", gClsNodeInit, gClsNodeTerm, gClsNodeInit - gClsNodeTerm
      Debug.Print "gFormInit", gFormInit, gFormTerm, gFormInit - gFormTerm
      MsgBox "NOT all Classes were terminated !" & vbCr & "see Immediate window", , GCSAPPNAME
    End If
    gClsTreeViewInit = 0
    gClsTreeViewTerm = 0
    gClsNodeInit = 0
    gClsNodeTerm = 0
    gFormInit = 0
    gFormTerm = 0
  End Sub
#End If

Public Sub ReleaseHighlighting()
  Window.ReleaseHighlighting
End Sub

Public Sub StartpHilbyUIAPolling()
  StoppHilbyUIAUIAPolling
  TimerID = SetTimer(0, 0, TIMER_INTERVAL, AddressOf TimerProc)
'Stop
  If TimerID = 0 Then MsgBox "Failed to start timer", vbCritical, "pHilby"
End Sub

Public Sub StoppHilbyUIAUIAPolling()
  If TimerID <> 0 Then
    KillTimer 0, TimerID
    TimerID = 0
  End If
End Sub

#If VBA7 Then
Private Sub TimerProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal nIDEvent As LongPtr, ByVal dwTime As Long)
#Else
Private Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal nIDEvent As Long, ByVal dwTime As Long)
#End If
  StoppHilbyUIAUIAPolling
  Dim CurrentTagPoint As tagPOINT
  GetCursorPos CurrentTagPoint
  'Don't update pHilby if we haven't moved the cursor!
  If Not ((CurrentTagPoint.X = PreviousTagPoint.X) And (CurrentTagPoint.Y = PreviousTagPoint.Y)) Then
    UserForm.SearchByCursorPoint CurrentTagPoint
    PollingExpiryTime = Now + TimeValue("00:00:02")
  End If
  PreviousTagPoint = CurrentTagPoint
  'Only poll to the expiry time
  If Now > PollingExpiryTime Then
    MsgBox "Element polling has stopped after no activity for 2 seconds.", vbInformation, "pHilby"
    UserForm.UnpushcbUIAPolling
    Exit Sub
  End If
  'Only poll while the pHilby
  If Not UserForm Is Nothing Then
    StartpHilbyUIAPolling
  End If
End Sub

