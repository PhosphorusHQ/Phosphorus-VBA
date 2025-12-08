VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pWindowsDriver_Navigate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Private This As NavigateProperties
Private Type NavigateProperties
  ParentWindowsDriver As pWindowsDriver
  SubDriver As Object
End Type

Public Sub Initialise(ByRef ParentWindowsDriver As pWindowsDriver, SubDriver As Object)
  Set This.ParentWindowsDriver = ParentWindowsDriver
  Set This.SubDriver = SubDriver
End Sub

Public Sub Refresh()
  This.SubDriver.IWindowsDriverWebBrowser_RefreshPage
End Sub
