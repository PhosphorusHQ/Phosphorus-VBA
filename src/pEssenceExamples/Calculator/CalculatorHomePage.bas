VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalculatorHomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Calculator
Option Explicit

Private MasterWindowElementSearch As pSearch
Private MainCalculatorSubWindowElementSearch As pSearch
Private NavigationViewRootElementSearch As pSearch
Private OpenCloseNavigationMenuButtonElementSearch As pSearch
Private NavigationMenuRootPaneWindowElementSearch As pSearch

Private Type PrivateElementNames
  MasterWindow As String
  MainCalculatorSubWindow As String
  NavigationViewRoot As String
  OpenCloseNavigationMenuButton As String
  NavigationMenuRootPaneWindow As String
End Type

Private This As PrivateElementNames
Private PrivateElements As New Scripting.Dictionary

Private Sub Class_Initialize()
  OpenCalculator
  InitializeAllElements
  FindMasterWindowElement
  FindNavigationElements
End Sub

Private Sub Class_Terminate()
  Actions.CloseWindow This.MasterWindow, PrivateElements(This.MasterWindow)
  Set MasterWindowElementSearch = Nothing
  Set MainCalculatorSubWindowElementSearch = Nothing
  Set NavigationViewRootElementSearch = Nothing
  Set OpenCloseNavigationMenuButtonElementSearch = Nothing
  Set NavigationMenuRootPaneWindowElementSearch = Nothing
  Set PrivateElements = Nothing
End Sub

Private Sub OpenCalculator()
  WindowsProcesses.RunShellExecuteToStartNewProcess "Microsoft Windows Calculator", "open", "shell:appsFolder\Microsoft.WindowsCalculator_8wekyb3d8bbwe!App", VBA.Constants.vbNullString, VBA.Constants.vbNullString, WindowStyle.Normal
End Sub

Private Sub InitializeAllElements()
  Set MasterWindowElementSearch = UIACommon.GetNewSearch
  Set MainCalculatorSubWindowElementSearch = UIACommon.GetNewSearch
  Set NavigationViewRootElementSearch = UIACommon.GetNewSearch
  Set OpenCloseNavigationMenuButtonElementSearch = UIACommon.GetNewSearch
  Set NavigationMenuRootPaneWindowElementSearch = UIACommon.GetNewSearch
  This.MasterWindow = "MasterWindow"
  This.MainCalculatorSubWindow = "MainCalculatorSubWindow"
  This.NavigationViewRoot = "NavigationView"
  This.OpenCloseNavigationMenuButton = "OpenCloseNavigationMenuButton"
  This.NavigationMenuRootPaneWindow = "NavigationMenuRootPaneWindow"
End Sub

Private Sub FindMasterWindowElement()
  
  With MasterWindowElementSearch
    .Initialise This.MasterWindow, GetRootDesktopElement, TreeScope.Children
    .AddCondition "NameIsCalculator", UIAProperties.Name, UIAPropertyComparisons.Equals, "Calculator"
    .AddCondition "ControlTypeIsWindow", UIAProperties.ControlType, UIAPropertyComparisons.Equals, pEssence.UIAControlTypeIDs.Window
    .AddCondition "ClassNameIsApplicationFrameWindow", UIAProperties.ClassName, UIAPropertyComparisons.Equals, "ApplicationFrameWindow"
    .Locator By.pConditions, "AND(NameIsCalculator, ControlTypeIsWindow, ClassNameIsApplicationFrameWindow)"
    PrivateElements.Add This.MasterWindow, .Find(10)
  End With

  Actions.WaitForPropertyValue _
    This.MasterWindow, _
    PrivateElements(This.MasterWindow), _
    UIAProperties.WindowWindowInteractionState, _
    pEssence.UIAWindowInteractionStates.ReadyForUserInteraction
    
End Sub

Private Sub FindNavigationElements()
  
  With MainCalculatorSubWindowElementSearch
    .Initialise This.MainCalculatorSubWindow, PrivateElements(This.MasterWindow), TreeScope.Children
    .AddCondition "NameIsCalculator", UIAProperties.Name, UIAPropertyComparisons.Equals, "Calculator"
    .AddCondition "ControlTypeIsWindow", UIAProperties.ControlType, UIAPropertyComparisons.Equals, pEssence.UIAControlTypeIDs.Window
    .AddCondition "ClassNameIsWindowsUICoreCoreWindow", UIAProperties.ClassName, UIAPropertyComparisons.Equals, "Windows.UI.Core.CoreWindow"
    .Locator By.pConditions, "AND(NameIsCalculator, ControlTypeIsWindow, ClassNameIsWindowsUICoreCoreWindow)"
    PrivateElements.Add This.MainCalculatorSubWindow, .Find(10)
  End With
  
  With NavigationViewRootElementSearch
    .Initialise This.NavigationViewRoot, PrivateElements(This.MainCalculatorSubWindow), pEssence.TreeScope.Children
    .Locator By.AutomationId, "NavView"
    PrivateElements.Add This.NavigationViewRoot, .Find(10)
  End With
    
  With OpenCloseNavigationMenuButtonElementSearch
    .Initialise This.OpenCloseNavigationMenuButton, PrivateElements(This.NavigationViewRoot), TreeScope.Children
    .AddCondition "NameIsOpenNavigation", UIAProperties.Name, UIAPropertyComparisons.Equals, "Open Navigation"
    .AddCondition "NameIsCloseNavigation", UIAProperties.Name, UIAPropertyComparisons.Equals, "Close Navigation"
    .AddCondition "ControlTypeIsButton", UIAProperties.ControlType, UIAPropertyComparisons.Equals, UIAControlTypeIDs.Button
    .AddCondition "ClassNameIsButton", UIAProperties.ClassName, UIAPropertyComparisons.Equals, "Button"
    .Locator By.pConditions, "AND(OR(NameIsOpenNavigation,NameIsCloseNavigation), ControlTypeIsButton, ClassNameIsButton)"
    PrivateElements.Add This.OpenCloseNavigationMenuButton, .Find(10)
  End With
        
End Sub

Public Sub SelectCalculatorType(CalculatorType As String)
  OpenNavigationMenu
  Actions.Click CalculatorType, GetMenuElement(CalculatorType)
End Sub

Private Sub OpenNavigationMenu()
  Actions.Click OpenCloseNavigationMenuButtonElementSearch.ElementName, PrivateElements(This.OpenCloseNavigationMenuButton)
End Sub

Private Function GetMenuElement(CalculatorType As String) As IUIAutomationElement
  
  'We need to find this element each time the menu is opened
  With NavigationMenuRootPaneWindowElementSearch
  
     If NavigationMenuRootPaneWindowElementSearch.ElementName = "" Then
      .Initialise This.NavigationMenuRootPaneWindow, PrivateElements(This.NavigationViewRoot), TreeScope.Children
      .AddCondition "AutomationIdIsPaneRoot", UIAProperties.AutomationId, UIAPropertyComparisons.Equals, "PaneRoot"
      .Locator By.pConditions, "AutomationIdIsPaneRoot"
     Else
      PrivateElements.Remove This.NavigationMenuRootPaneWindow
     End If
    PrivateElements.Add This.NavigationMenuRootPaneWindow, .Find
  End With

  Dim CurrentNavigationMenuItemElementSearch As pSearch
  Set CurrentNavigationMenuItemElementSearch = Nothing
  Set CurrentNavigationMenuItemElementSearch = UIACommon.GetNewSearch
  
  With CurrentNavigationMenuItemElementSearch
    .Initialise CalculatorType, PrivateElements(This.NavigationMenuRootPaneWindow), TreeScope.Descendants
    .Locator By.pConditions, "AND(NameIs, ControlTypeIsListItem)"
    .AddCondition "NameIs", UIAProperties.Name, UIAPropertyComparisons.Equals, CalculatorType
    .AddCondition "ControlTypeIsListItem", UIAProperties.ControlType, UIAPropertyComparisons.Equals, UIAControlTypeIDs.ListItem
    Set GetMenuElement = .Find
  End With

End Function

