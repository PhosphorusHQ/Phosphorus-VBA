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
  MasterWindowName As String
  BrowserRootViewElementSearch As pSearch
  BrowserRootViewElementName As String
'  SidebarContentsSplitViewSearch As pSearch
'  SidebarContentsSplitViewName As String
  RootWebAreaElementSearch As pSearch
  BrowserRootWebAreaElementName As String
End Type

Private This As BrowserAttributes
Private PrivateElements As New Scripting.Dictionary
  
Private Sub Class_Initialize()
  Set This.MasterWindowElementSearch = Factory.GetNewSearch
  Set This.BrowserRootViewElementSearch = Factory.GetNewSearch
  Set This.RootWebAreaElementSearch = Factory.GetNewSearch
  This.MasterWindowName = "MasterWindow"
  This.BrowserRootViewElementName = "BrowserRootViewElement"
'  This.SidebarContentsSplitViewName = "SidebarContentsSplitView"
  This.BrowserRootWebAreaElementName = "BrowserRootWebArea"
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  Window.CloseWindow This.MasterWindowName, PrivateElements(This.MasterWindowName)
  On Error GoTo 0
  Set This.MasterWindowElementSearch = Nothing
  Set This.BrowserRootViewElementSearch = Nothing
'  Set This.SidebarContentsSplitViewSearch = Nothing
  Set This.RootWebAreaElementSearch = Nothing
End Sub

Public Sub StartEdge(WebAppName As String, URL As String, WebAppPageTitle As String)
  This.WebAppName = WebAppName
  This.URL = URL
  This.WebAppPageTitle = WebAppPageTitle
  LaunchCommandByProtocol This.WebAppName, "microsoft-edge:", This.URL, WindowStyle.Maximized
  FindRootWindowElements
End Sub

Private Sub FindRootWindowElements()

  'Use: AscW & ChrW to determine embedded Unicode characters
  With This.MasterWindowElementSearch
    .Initialise This.MasterWindowName, RootDesktopUIAElement, Children, pConditions, "AND(NameStartsWith, NameEndsWith, ControlType, ClassName, WindowInteractionState  )"
    .Condition "NameStartsWith", Name, StartsWith, This.WebAppPageTitle & " - "
    .Condition "NameEndsWith", Name, EndsWith, "Microsoft" & ChrW(8203) & " Edge"
    .Condition "ControlType", ControlType, Equals, UIAControlTypeIDs.Window
    .Condition "ClassName", ClassName, Equals, "Chrome_WidgetWin_1"
    .Condition "WindowInteractionState", WindowWindowInteractionState, Equals, UIAWindowInteractionStates.ReadyForUserInteraction
    PrivateElements.Add This.MasterWindowName, .Find(10)
  End With

  With This.BrowserRootViewElementSearch
    .Initialise This.BrowserRootViewElementName, PrivateElements(This.MasterWindowName), TreeScope.Children, By.pConditions, "AND(ControlType, ClassName)"
    .Condition "ControlType", ControlType, Equals, UIAControlTypeIDs.Pane
    .Condition "ClassName", ClassName, Equals, "BrowserRootView"
    PrivateElements.Add This.BrowserRootViewElementName, .Find(10)
  End With
'More elelement here:
'ClassName "NonClientView"
'ClassName "BrowserFrameViewWin"
'ClassName "BrowserView"

'Then:
'ClassName: "SidebarContentsSplitView"
'ControlType: Pane
'FrameworkID: "Chrome"
'  With This.SidebarContentsSplitViewSearch
'    .Initialise This.SidebarContentsSplitViewName, PrivateElements(This.BrowserRootViewElementName), TreeScope.Children
'    .AddCondition "ControlTypeIsPane", UIAProperties.ControlType, UIAPropertyComparisons.Equals, UIAControlTypeIDs.Pane
'  End With


'-------------

'Then finally the RootWebArea ... this needs to return after finding the first match!
'  With This.RootWebAreaElementSearch
'    .Initialise This.BrowserRootWebAreaElementName, PrivateElements(This.BrowserRootViewElementName), TreeScope.Children
'    .AddCondition "NameIsWebAppPageTitle", UIAProperties.Name, UIAPropertyComparisons.Equals, This.WebAppPageTitle
'    .AddCondition "ControlTypeIsDocument", UIAProperties.ControlType, UIAPropertyComparisons.Equals, UIAControlTypeIDs.Document
'    .AddCondition "AutomationIdIsRootWebArea", UIAProperties.AutomationId, UIAPropertyComparisons.Equals, "RootWebArea"
'    .Locator By.pConditions, "AND(NameIsWebAppPageTitle, ControlTypeIsDocument, AutomationIdIsRootWebArea)"
'    PrivateElements.Add This.BrowserRootWebAreaElementName, .Find(10)
'  End With

'BrowserRootWebAreaElement
'//Document[And(@AutomationId="RootWebArea",@Name="Example Domain")]
  
End Sub
