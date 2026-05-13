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

Public Const GCSAPPNAME As String = "pHilby - UI Automation Spy"
'Public Const GCSBUILD As String = "26.5"

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
'Public gbCellsChanged As Boolean

Private Sub StartTest()
  pHilby.Start
End Sub

Public Sub Start(Optional RootUIAElement As IUIAutomationElement)

  Toaster.PopDown
 
  Dim MaxNumberOfLevels As Integer
  MaxNumberOfLevels = 0
  If RootUIAElement Is Nothing Then
    Set RootUIAElement = Factory.GetRootDesktopElement
    MaxNumberOfLevels = 0
  End If
  
  Dim ufrmForm As frmpHilby
  Set ufrmForm = New frmpHilby
  With ufrmForm
    .LoadTreeView RootUIAElement, MaxNumberOfLevels:=MaxNumberOfLevels
    .AppName = GCSAPPNAME
    .Show
    Unload ufrmForm
    Set ufrmForm = Nothing
  End With

  #If DEBUGMODE = 1 Then
    ClassCounts
  #End If
  
'  AppActivate ufrmForm.Caption

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
