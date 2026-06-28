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
  Brave
  Chrome
  Chromium
  Edge
  Epic
  Firefox
  Opera
  Samsung
  Yandex
  [_Last]
End Enum

Public CurrentWebBrowserType As WebBrowserType

Public Function GetWebBrowserName(TargetWebBrowserType As WebBrowserType) As String
  Dim R As String
  Select Case TargetWebBrowserType
    Case Brave: R = "Brave"
    Case Chrome: R = "Chrome"
    Case Chromium: R = "Chrome"
    Case Edge: R = "Edge"
    Case Epic: R = "Epic"
    Case Firefox: R = "Firefox"
    Case Opera: R = "Opera"
    Case Samsung: R = "Samsung"
    Case Yandex: R = "Yandex"
    Case Else: R = "Unhandled WebBrowser type in GetWebBrowserName #" & TargetWebBrowserType
  End Select
  GetWebBrowserName = R
End Function

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
    Case WebBrowserType.Brave
      Set GetNewWebBrowser = New BraveWebBrowser
    Case WebBrowserType.Chrome
      Set GetNewWebBrowser = New ChromeWebBrowser
    Case WebBrowserType.Chromium
      Set GetNewWebBrowser = New ChromiumWebBrowser
    Case WebBrowserType.Edge
      Set GetNewWebBrowser = New EdgeWebBrowser
    Case WebBrowserType.Epic
      Set GetNewWebBrowser = New EpicWebBrowser
    Case WebBrowserType.Firefox
      Set GetNewWebBrowser = New FirefoxWebBrowser
    Case WebBrowserType.Opera
      Set GetNewWebBrowser = New OperaWebBrowser
    Case WebBrowserType.Samsung
      Set GetNewWebBrowser = New SamsungWebBrowser
    Case WebBrowserType.Yandex
      Set GetNewWebBrowser = New YandexWebBrowser
    Case Else
      pExceptions.Raise pEssenceUnhandledWebBrowserType
  End Select
End Function

Public Function GetNewBoundingRectangle(Element As IUIAutomationElement) As BoundingRectangle
  Dim BR As BoundingRectangle
  Set BR = New BoundingRectangle
  BR.GetCurrentBoundingRectangle Element
  Set GetNewBoundingRectangle = BR
End Function
