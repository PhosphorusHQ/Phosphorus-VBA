VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoundingRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder UIAutomation
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
    
Private Rect As UIAProps.Rectangle

Public Sub GetCurrentBoundingRectangle(Element As IUIAutomationElement)
  Dim RectArray As Variant
  On Error Resume Next
  RectArray = Element.GetCurrentPropertyValue(UIAProperties.BoundingRectangle)
  On Error GoTo 0
  If IsArray(RectArray) Then
    Rect.Left = RectArray(0)
    Rect.Top = RectArray(1)
    Rect.Right = RectArray(0) + RectArray(2)
    Rect.Bottom = RectArray(1) + RectArray(3)
    Rect.Width = RectArray(2) ' RectArray(2) - RectArray(0)
    Rect.Height = RectArray(3) 'RectArray(3) - RectArray(1)
  End If
End Sub

Public Function GetRectangle() As UIAProps.Rectangle
  GetRectangle = Rect
End Function

Public Function Left() As Double
  Left = Rect.Left
End Function

Public Function Top() As Double
  Top = Rect.Top
End Function

Public Function Right() As Double
  Right = Rect.Right
End Function

Public Function Bottom() As Double
  Bottom = Rect.Bottom
End Function

Public Function Width() As Double
  Width = Rect.Width
End Function

Public Function Height() As Double
  Height = Rect.Height
End Function

