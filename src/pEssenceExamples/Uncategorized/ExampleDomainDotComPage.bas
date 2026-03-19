VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExampleDomainDotComPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const WEB_APP_NAME = "Example.com"
Const TARGET_PAGE_URL = "https://www.example.com/"
Const TARGET_PAGE_TITLE = "Example Domain"

Private WebBrowser As EdgeWebBrowser

Private Sub Class_Initialize()
  Set WebBrowser = Factory.GetNewEdgeWebBrowser
End Sub

Private Sub Class_Terminate()
  Set WebBrowser = Nothing
End Sub

Public Sub Initialize()
  WebBrowser.StartEdge WEB_APP_NAME, TARGET_PAGE_URL
End Sub


