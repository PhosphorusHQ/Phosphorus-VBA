VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExampleDomainDotComPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder ExampleDomainDotCom
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Const WEB_APP_NAME = "Example.com"
Const TARGET_PAGE_URL = "https://www.example.com/"
Const TARGET_PAGE_TITLE = "Example Domain"

Private WebBrowser As EdgeWebBrowser

Private Type PrivateElementNames
  MasterWindow As String
End Type

Private This As PrivateElementNames
Private PrivateElements As New Scripting.Dictionary

Private Sub Class_Initialize()
  Factory.GetRootDesktopElement
  Set WebBrowser = Factory.GetNewEdgeWebBrowser
End Sub

Private Sub Class_Terminate()
  Set WebBrowser = Nothing
End Sub

Public Sub Initialize()
  With WebBrowser
    .StartEdge WEB_APP_NAME, TARGET_PAGE_URL, TARGET_PAGE_TITLE
  End With
End Sub
