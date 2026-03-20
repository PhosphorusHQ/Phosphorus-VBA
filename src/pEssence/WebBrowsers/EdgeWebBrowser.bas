VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EdgeWebBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder WebBrowsers
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Private Type BrowserAttributes
  WebAppName As String
  URL As String
  WebAppPageTitle As String
  MasterWindowElementSearch As pSearch
  MasterWindow As String
End Type

Private This As BrowserAttributes
Private PrivateElements As New Scripting.Dictionary
  
Private Sub Class_Initialize()
  Set This.MasterWindowElementSearch = Factory.GetNewSearch
  This.MasterWindow = "MasterWindow"
End Sub

Private Sub Class_Terminate()
  Window.CloseWindow This.MasterWindow, PrivateElements(This.MasterWindow)
  Set This.MasterWindowElementSearch = Nothing
End Sub

Public Sub StartEdge(WebAppName As String, URL As String, WebAppPageTitle As String)
  This.WebAppName = WebAppName
  This.URL = URL
  This.WebAppPageTitle = WebAppPageTitle
  LaunchCommandByProtocol This.WebAppName, "microsoft-edge:", This.URL, WindowStyle.Maximized
  FindRootWindowElements
End Sub

Private Sub FindRootWindowElements()

  With This.MasterWindowElementSearch
    .Initialise This.MasterWindow, GetRootDesktopElement, TreeScope.Children
    .AddCondition "NameStartsWithExampleDomain", UIAProperties.Name, UIAPropertyComparisons.StartsWith, This.WebAppPageTitle & " - "
    'Use: AscW & ChrW to determine embedded Unicode characters
    .AddCondition "NameEndsWithMicrosoftEdge", UIAProperties.Name, UIAPropertyComparisons.EndsWith, "Microsoft" & ChrW(8203) & " Edge"
    .AddCondition "ControlTypeIsWindow", UIAProperties.ControlType, UIAPropertyComparisons.Equals, pEssence.UIAControlTypeIDs.Window
    .AddCondition "ClassNameIsChromeWidgetWin1", UIAProperties.ClassName, UIAPropertyComparisons.Equals, "Chrome_WidgetWin_1"
    .AddCondition "WindowInteractionStateIsReadyForUserInteraction", UIAProperties.WindowWindowInteractionState, UIAPropertyComparisons.Equals, UIAWindowInteractionStates.ReadyForUserInteraction
    .Locator By.pConditions, "AND(NameStartsWithExampleDomain, NameEndsWithMicrosoftEdge, ControlTypeIsWindow, ClassNameIsChromeWidgetWin1, WindowInteractionStateIsReadyForUserInteraction)"
     PrivateElements.Add This.MasterWindow, .Find(10)
  End With
  
End Sub
'BrowserRootViewElement
'//Pane[@ClassName="BrowserRootView"]

'BrowserRootWebAreaElement
'//Document[And(@AutomationId="RootWebArea",@Name="Example Domain")]
