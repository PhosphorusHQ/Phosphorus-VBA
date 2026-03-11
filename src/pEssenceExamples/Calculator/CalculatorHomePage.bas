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

Private MasterWindowElement As pSearch
Private MainCalculatorSubWindowElement As pSearch
Private NavViewElement As pSearch
Private OpenCloseNavigationMenuButtonElement As pSearch
Private NavigationMenuRootPaneWindowElement As pSearch
Private StandardCalculatorNavigationMenuItemElement As pSearch

Private Sub OpenCalculator()
  pEssence.WindowsProcesses.RunShellExecuteToStartNewProcess "Microsoft Windows Calculator", "open", "shell:appsFolder\Microsoft.WindowsCalculator_8wekyb3d8bbwe!App", VBA.Constants.vbNullString, VBA.Constants.vbNullString, WindowStyle.Normal
End Sub

Private Sub Class_Initialize()
  OpenCalculator
  InitializeAllElements
  pEssence.WindowsProcesses.Snooze 2000
  FindMasterWindowElement
  FindNavigationElements
End Sub

Private Sub InitializeAllElements()
  Set MasterWindowElement = pEssence.UIACommon.GetNewSearch
  Set MainCalculatorSubWindowElement = pEssence.UIACommon.GetNewSearch
  Set NavViewElement = pEssence.UIACommon.GetNewSearch
  Set OpenCloseNavigationMenuButtonElement = pEssence.UIACommon.GetNewSearch
  Set NavigationMenuRootPaneWindowElement = pEssence.UIACommon.GetNewSearch
  Set StandardCalculatorNavigationMenuItemElement = pEssence.UIACommon.GetNewSearch
End Sub

Private Sub FindMasterWindowElement()
  
  With MasterWindowElement
    .Initialise _
      "MasterWindowElement", _
      UIACommon.GetRootDesktopElement, _
      pEssence.TreeScope.Children
    .AddCondition "NameIsCalculator", UIAProperties.Name, UIAPropertyComparisons.Equals, "Calculator"
    .AddCondition "ControlTypeIsWindow", UIAProperties.ControlType, UIAPropertyComparisons.Equals, pEssence.UIAControlTypeIDs.Window
    .AddCondition "ClassNameIsApplicationFrameWindow", UIAProperties.ClassName, UIAPropertyComparisons.Equals, "ApplicationFrameWindow"
    .Locator By.pConditions, "AND(NameIsCalculator, ControlTypeIsWindow, ClassNameIsApplicationFrameWindow)"
    .Find
  End With

  MasterWindowElement.WaitForPropertyValue _
    UIAProperties.WindowWindowInteractionState, _
    pEssence.UIAWindowInteractionStates.ReadyForUserInteraction

End Sub

Private Sub FindNavigationElements()
  
  With MainCalculatorSubWindowElement
    .Initialise _
      "MainCalculatorSubWindowElement", _
      MasterWindowElement.FoundUIAElement, _
      pEssence.TreeScope.Children
    .AddCondition "NameIsCalculator", UIAProperties.Name, UIAPropertyComparisons.Equals, "Calculator"
    .AddCondition "ControlTypeIsWindow", UIAProperties.ControlType, UIAPropertyComparisons.Equals, pEssence.UIAControlTypeIDs.Window
    .AddCondition "ClassNameIsWindowsUICoreCoreWindow", UIAProperties.ClassName, pEssence.UIAPropertyComparisons.Equals, "Windows.UI.Core.CoreWindow"
    .Locator By.pConditions, "AND(NameIsCalculator, ControlTypeIsWindow, ClassNameIsWindowsUICoreCoreWindow)"
    .Find
  End With
  
  With NavViewElement
    .Initialise _
      "NavViewElement", _
      MainCalculatorSubWindowElement.FoundUIAElement, _
      pEssence.TreeScope.Children
    .Locator By.AutomationId, "NavView"
    .Find
  End With
    
  With OpenCloseNavigationMenuButtonElement
    .Initialise _
      "OpenCloseNavigationMenuButtonElement", _
      NavViewElement.FoundUIAElement, _
      pEssence.TreeScope.Children
    .AddCondition "NameIsOpenNavigation", UIAProperties.Name, UIAPropertyComparisons.Equals, "Open Navigation"
    .AddCondition "NameIsCloseNavigation", UIAProperties.Name, UIAPropertyComparisons.Equals, "Close Navigation"
    .AddCondition "ControlTypeIsButton", UIAProperties.ControlType, UIAPropertyComparisons.Equals, pEssence.UIAControlTypeIDs.Button
    .AddCondition "ClassNameIsButton", UIAProperties.ClassName, pEssence.UIAPropertyComparisons.Equals, "Button"
    .Locator By.pConditions, "AND(OR(NameIsOpenNavigation,NameIsCloseNavigation), ControlTypeIsButton, ClassNameIsButton)"
    .Find
  End With
        
End Sub

Private Sub OpenNavigationMenu()
  Actions.Click OpenCloseNavigationMenuButtonElement.ElementName, OpenCloseNavigationMenuButtonElement.FoundUIAElement
End Sub

Private Sub InitialiseExpandedNavigationMenuElements()
  
  With NavigationMenuRootPaneWindowElement
    .Initialise _
      "NavigationMenuRootPaneWindowElement", _
      NavViewElement.FoundUIAElement, _
      pEssence.TreeScope.Children
    .AddCondition "AutomationIdIsPaneRoot", UIAProperties.AutomationId, UIAPropertyComparisons.Equals, "PaneRoot"
    .Locator By.pConditions, "AutomationIdIsPaneRoot"
    .Find
  End With
  
  With StandardCalculatorNavigationMenuItemElement
    .Initialise _
      "StandardCalculatorNavigationMenuItemElement", _
      NavigationMenuRootPaneWindowElement.FoundUIAElement, _
      pEssence.TreeScope.Descendants
    .Locator By.pConditions, "AND(NameIsStandardCalculator, ControlTypeIsListItem)"
    .AddCondition "NameIsStandardCalculator", UIAProperties.Name, UIAPropertyComparisons.Equals, "Standard Calculator"
    .AddCondition "ControlTypeIsListItem", UIAProperties.ControlType, UIAPropertyComparisons.Equals, pEssence.UIAControlTypeIDs.ListItem
    .Find
  End With

End Sub

Public Sub SelectCalculatorType(CalculatorType As String)
  OpenNavigationMenu
  pEssence.WindowsProcesses.Snooze 2000
  InitialiseExpandedNavigationMenuElements
  Select Case CalculatorType
    Case "Standard Calculator"
      StandardCalculatorNavigationMenuItemElement.Find
      Actions.Click StandardCalculatorNavigationMenuItemElement.ElementName, StandardCalculatorNavigationMenuItemElement.FoundUIAElement
    Case Else
      MsgBox "?"
  End Select
End Sub
