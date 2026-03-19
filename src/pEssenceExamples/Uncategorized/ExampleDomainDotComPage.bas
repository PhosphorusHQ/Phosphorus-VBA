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
Private MasterWindowElementSearch As pSearch

Private Type PrivateElementNames
  MasterWindow As String
End Type

Private This As PrivateElementNames
Private PrivateElements As New Scripting.Dictionary

Private Sub Class_Initialize()
  Set WebBrowser = Factory.GetNewEdgeWebBrowser
  Set MasterWindowElementSearch = Factory.GetNewSearch
  This.MasterWindow = "MasterWindow"
End Sub

Private Sub Class_Terminate()
  Window.CloseWindow This.MasterWindow, PrivateElements(This.MasterWindow)
  Set WebBrowser = Nothing
  Set MasterWindowElementSearch = Nothing
End Sub

Public Sub Initialize()
  WebBrowser.StartEdge WEB_APP_NAME, TARGET_PAGE_URL
  FindMasterWindowElement
End Sub

Private Sub FindMasterWindowElement()

  With MasterWindowElementSearch
    .Initialise This.MasterWindow, GetRootDesktopElement, TreeScope.Children
    .AddCondition "NameStartsWithExampleDomain", UIAProperties.Name, UIAPropertyComparisons.StartsWith, "Example Domain - "
    .AddCondition "NameEndsWithMicrosoftEdge", UIAProperties.Name, UIAPropertyComparisons.EndsWith, "Microsoft? Edge"
    .AddCondition "ControlTypeIsWindow", UIAProperties.ControlType, UIAPropertyComparisons.Equals, pEssence.UIAControlTypeIDs.Window
    .AddCondition "ClassNameIsChromeWidgetWin1", UIAProperties.ClassName, UIAPropertyComparisons.Equals, "Chrome_WidgetWin_1"
    .AddCondition "WindowInteractionStateIsReadyForUserInteraction", UIAProperties.WindowWindowInteractionState, UIAPropertyComparisons.Equals, UIAWindowInteractionStates.ReadyForUserInteraction
    .Locator by.pConditions, "AND(NameStartsWithExampleDomain, NameEndsWithMicrosoftEdge, ControlTypeIsWindow, ClassNameIsChromeWidgetWin1, WindowInteractionStateIsReadyForUserInteraction)"
    PrivateElements.Add This.MasterWindow, .Find(10)
  End With
      
End Sub

