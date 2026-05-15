Attribute VB_Name = "Toolbar"
'@Folder Common
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Private Const TOOLBAR_NAME As String = "Phosphorus"
Public PhosphorusToolbar As CommandBar

'https://bettersolutions.com/vba/ribbon/face-ids-2003.htm

Public Sub CreatePhosphorusToolbar()
    
  ' Remove old version first to avoid duplicates
  DeletePhosphorusToolbar
    
  ' Create new toolbar (will appear under Add-ins tab)
  Set PhosphorusToolbar = Application.CommandBars.Add( _
    Name:=TOOLBAR_NAME, _
    Position:=msoBarTop, _
    MenuBar:=False, _
    Temporary:=True)          ' True = deleted when Excel closes
  PhosphorusToolbar.Visible = True
    
  'Need to create at least 1 control!
  Dim btn As CommandBarButton
  Set btn = PhosphorusToolbar.Controls.Add(Type:=msoControlButton)
  With btn
    .Caption = "Reset Cursor"
    .Style = msoButtonIconAndCaption
    .OnAction = "ResetCursor"
    .FaceId = 59   'Smiley face
    .TooltipText = "Dummy Test"
  End With

End Sub

Public Sub DeletePhosphorusToolbar()
  On Error Resume Next
  Application.CommandBars(TOOLBAR_NAME).Delete
  On Error GoTo 0
End Sub

Public Sub ResetCursor()
  Application.Cursor = xlDefault
End Sub
