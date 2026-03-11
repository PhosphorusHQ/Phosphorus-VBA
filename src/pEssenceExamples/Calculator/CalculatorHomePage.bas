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
Private NavViewRootElementSearch As pSearch
Private OpenCloseNavigationMenuButtonElementSearch As pSearch
Private NavigationMenuRootPaneWindowElementSearch As pSearch
Private StandardCalculatorNavigationMenuItemElementSearch As pSearch

Public Elements As New Scripting.Dictionary
Private PrivateElements As New Scripting.Dictionary

Private Type PrivateElementNames
  MasterWindow As String
  MainCalculatorSubWindow As String
  NavigationViewRoot As String
  OpenCloseNavigationMenuButton As String
  NavigationMenuRootPaneWindow As String
  StandardCalculatorNavigationMenuItem As String
End Type

Private This As PrivateElementNames

Private Sub Class_Initialize()
  OpenCalculator
  InitializeAllElements
  pEssence.WindowsProcesses.Snooze 2000
  FindMasterWindowElement
  FindNavigationElements
End Sub

Private Sub Class_Terminate()
  Set MasterWindowElementSearch = Nothing
  Set MainCalculatorSubWindowElementSearch = Nothing
  Set NavViewRootElementSearch = Nothing
  Set OpenCloseNavigationMenuButtonElementSearch = Nothing
  Set NavigationMenuRootPaneWindowElementSearch = Nothing
  Set StandardCalculatorNavigationMenuItemElementSearch = Nothing
  Set Elements = Nothing
  Set PrivateElements = Nothing
End Sub

Private Sub OpenCalculator()
  WindowsProcesses.RunShellExecuteToStartNewProcess "Microsoft Windows Calculator", "open", "shell:appsFolder\Microsoft.WindowsCalculator_8wekyb3d8bbwe!App", VBA.Constants.vbNullString, VBA.Constants.vbNullString, WindowStyle.Normal
End Sub

Private Sub InitializeAllElements()
  Set MasterWindowElementSearch = UIACommon.GetNewSearch
  Set MainCalculatorSubWindowElementSearch = UIACommon.GetNewSearch
  Set NavViewRootElementSearch = UIACommon.GetNewSearch
  Set OpenCloseNavigationMenuButtonElementSearch = UIACommon.GetNewSearch
  Set NavigationMenuRootPaneWindowElementSearch = UIACommon.GetNewSearch
  Set StandardCalculatorNavigationMenuItemElementSearch = UIACommon.GetNewSearch
  This.MasterWindow = "MasterWindow"
  This.MainCalculatorSubWindow = "MainCalculatorSubWindow"
  This.NavigationViewRoot = "NavigationView"
  This.OpenCloseNavigationMenuButton = "OpenCloseNavigationMenuButton"
  This.NavigationMenuRootPaneWindow = "NavigationMenuRootPaneWindow"
  This.StandardCalculatorNavigationMenuItem = "StandardCalculatorNavigationMenuItem"
End Sub

Private Sub FindMasterWindowElement()
  
  With MasterWindowElementSearch
    .Initialise This.MasterWindow, GetRootDesktopElement, TreeScope.Children
    .AddCondition "NameIsCalculator", UIAProperties.Name, UIAPropertyComparisons.Equals, "Calculator"
    .AddCondition "ControlTypeIsWindow", UIAProperties.ControlType, UIAPropertyComparisons.Equals, pEssence.UIAControlTypeIDs.Window
    .AddCondition "ClassNameIsApplicationFrameWindow", UIAProperties.ClassName, UIAPropertyComparisons.Equals, "ApplicationFrameWindow"
    .Locator By.pConditions, "AND(NameIsCalculator, ControlTypeIsWindow, ClassNameIsApplicationFrameWindow)"
    PrivateElements.Add This.MasterWindow, .Find
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
    PrivateElements.Add This.MainCalculatorSubWindow, .Find
  End With
  
  With NavViewRootElementSearch
    .Initialise This.NavigationViewRoot, PrivateElements(This.MainCalculatorSubWindow), pEssence.TreeScope.Children
    .Locator By.AutomationId, "NavView"
    PrivateElements.Add This.NavigationViewRoot, .Find
  End With
    
  With OpenCloseNavigationMenuButtonElementSearch
    .Initialise This.OpenCloseNavigationMenuButton, PrivateElements(This.NavigationViewRoot), TreeScope.Children
    .AddCondition "NameIsOpenNavigation", UIAProperties.Name, UIAPropertyComparisons.Equals, "Open Navigation"
    .AddCondition "NameIsCloseNavigation", UIAProperties.Name, UIAPropertyComparisons.Equals, "Close Navigation"
    .AddCondition "ControlTypeIsButton", UIAProperties.ControlType, UIAPropertyComparisons.Equals, UIAControlTypeIDs.Button
    .AddCondition "ClassNameIsButton", UIAProperties.ClassName, UIAPropertyComparisons.Equals, "Button"
    .Locator By.pConditions, "AND(OR(NameIsOpenNavigation,NameIsCloseNavigation), ControlTypeIsButton, ClassNameIsButton)"
    PrivateElements.Add This.OpenCloseNavigationMenuButton, .Find
  End With
        
End Sub

Public Sub SelectCalculatorType(CalculatorType As String)
  OpenNavigationMenu
  WindowsProcesses.Snooze 2000
  InitialiseExpandedNavigationMenuElements
  Select Case CalculatorType
    Case "Standard Calculator"
      Actions.Click This.StandardCalculatorNavigationMenuItem, PrivateElements(This.StandardCalculatorNavigationMenuItem)
    Case Else
      MsgBox "?"
  End Select
End Sub

Private Sub OpenNavigationMenu()
  Actions.Click OpenCloseNavigationMenuButtonElementSearch.ElementName, PrivateElements(This.OpenCloseNavigationMenuButton)
  'OpenCloseNavigationMenuButtonElementSearch.FoundUIAElement
End Sub

Private Sub InitialiseExpandedNavigationMenuElements()
  
  With NavigationMenuRootPaneWindowElementSearch
    .Initialise This.NavigationMenuRootPaneWindow, PrivateElements(This.NavigationViewRoot), TreeScope.Children
    .AddCondition "AutomationIdIsPaneRoot", UIAProperties.AutomationId, UIAPropertyComparisons.Equals, "PaneRoot"
    .Locator By.pConditions, "AutomationIdIsPaneRoot"
    PrivateElements.Add This.NavigationMenuRootPaneWindow, .Find
  End With

  With StandardCalculatorNavigationMenuItemElementSearch
    .Initialise This.StandardCalculatorNavigationMenuItem, PrivateElements(This.NavigationMenuRootPaneWindow), TreeScope.Descendants
    .Locator By.pConditions, "AND(NameIsStandardCalculator, ControlTypeIsListItem)"
    .AddCondition "NameIsStandardCalculator", UIAProperties.Name, UIAPropertyComparisons.Equals, "Standard Calculator"
    .AddCondition "ControlTypeIsListItem", UIAProperties.ControlType, UIAPropertyComparisons.Equals, pEssence.UIAControlTypeIDs.ListItem
    PrivateElements.Add This.StandardCalculatorNavigationMenuItem, .Find
  End With

End Sub
