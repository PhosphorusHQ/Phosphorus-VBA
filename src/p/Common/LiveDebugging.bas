Attribute VB_Name = "LiveDebugging"
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

Private LiveBugging As Boolean

Public Sub Start()
  LiveBugging = True
End Sub

Public Sub Finish()
  LiveBugging = False
End Sub

Public Sub Pause()
  If LiveBugging Then Stop
End Sub

Public Function Running() As Boolean
  Running = LiveBugging
End Function
