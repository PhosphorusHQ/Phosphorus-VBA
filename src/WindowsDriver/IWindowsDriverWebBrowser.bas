VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWindowsDriverWebBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder WindowsDriver
Option Explicit

'https://bettersolutions.com/vba/class-modules/implements.htm
'An interface contains only method & function signatures, not any properties

Public Function GetParentWindowsDriver() As pWindowsDriver
End Function

Public Sub LaunchApp(ByRef ParentWindowsDriver As pWindowsDriver, WebAppName As String, WebAppTitle As String, Optional URL As String)
End Sub

