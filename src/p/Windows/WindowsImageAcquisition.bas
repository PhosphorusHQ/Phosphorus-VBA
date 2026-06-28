Attribute VB_Name = "WindowsImageAcquisition"
'@Folder Windows
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Sub test()
  WindowsImageAcquisition.LoadImageWIA ThisWorkbook.Path & "\images\Logo.png"
End Sub

Function LoadImage(ByVal FilePath As String) As StdPicture
  On Error Resume Next
  Dim img As WIA.ImageFile
  Set img = New WIA.ImageFile
  img.LoadFile FilePath
  Set LoadImage = img.FileData.Picture
  On Error GoTo 0
End Function

Function LoadAndResizePNG( _
  ByVal FilePath As String, _
  ByVal MaxWidth As Long, _
  ByVal MaxHeight As Long) As StdPicture
    
  On Error Resume Next
    
  Dim img As WIA.ImageFile
  Dim proc As WIA.ImageProcess
    
  Set img = New WIA.ImageFile
  img.LoadFile FilePath
    
  Set proc = New WIA.ImageProcess
    
  ' Add scaling filter
  proc.Filters.Add proc.FilterInfos("Scale").FilterID
  With proc.Filters(1)
    .Properties("MaximumWidth") = MaxWidth
    .Properties("MaximumHeight") = MaxHeight
    .Properties("PreserveAspectRatio") = True
  End With
    
  Set LoadAndResizePNG = proc.Apply(img).FileData.Picture
    
  On Error GoTo 0

End Function

