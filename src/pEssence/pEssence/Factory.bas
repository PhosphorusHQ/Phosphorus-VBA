VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder pEssence
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public Enum WebBrowserType
  [_First]
  Chrome
  Edge
  Firefox
  Opera 'Not ready yet!
  [_Last]
End Enum

Public CurrentWebBrowserType As WebBrowserType

Public Function GetRootDesktopElement() As IUIAutomationElement
  If RootDesktopUIAElement Is Nothing Then
    Set RootDesktopUIAElement = UIA.GetRootElement
  End If
   Set GetRootDesktopElement = RootDesktopUIAElement
End Function

Public Function GetNewLocator() As pLocator
  Dim Locator As pLocator
  Set Locator = New pLocator
  Set GetNewLocator = Locator
End Function

Public Function GetNewElement(GivenName As String, UIAElement As IUIAutomationElement) As pElement
  Dim Element As pElement
  Set Element = New pElement
  Element.GivenName = GivenName
  Set Element.UIAElement = UIAElement
  Set GetNewElement = Element
End Function

Public Function GetNewWebBrowser() As Object
  Select Case CurrentWebBrowserType
    Case WebBrowserType.Chrome
      Set GetNewWebBrowser = New ChromeWebBrowser
    Case WebBrowserType.Edge
      Set GetNewWebBrowser = New EdgeWebBrowser
    Case WebBrowserType.Firefox
      Set GetNewWebBrowser = New FirefoxWebBrowser
    Case WebBrowserType.Opera
      Set GetNewWebBrowser = New OperaWebBrowser
    Case Else
      pExceptions.Raise pEssenceUnhandledWebBrowserType
  End Select
End Function
