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

Private Sub Test()
  WindowsImageAcquisition.LoadImage ThisWorkbook.Path & "\images\Logo.png"
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

' In a standard module
Function LoadImageSolidBackground(ByVal FilePath As String, _
                                  Optional BackColor As Long = vbWhite) As StdPicture
    
    Dim img As WIA.ImageFile
    Dim proc As WIA.ImageProcess
    
    Set img = New WIA.ImageFile
    img.LoadFile FilePath
    
    Set proc = New WIA.ImageProcess
    
    ' Add a background color filter (removes transparency)
    proc.Filters.Add proc.FilterInfos("Convert").FilterID
    proc.Filters(1).Properties("FormatID") = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}" ' BMP format
    'proc.Filters.Add proc.FilterInfos("Stamp").FilterID   ' Alternative method if needed
    
    Set LoadImageSolidBackground = proc.Apply(img).FileData.Picture
End Function

