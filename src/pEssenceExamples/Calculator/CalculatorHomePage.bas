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
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Private MasterWindowElementSearch As pSearch
Private MainCalculatorSubWindowElementSearch As pSearch
Private NavigationViewRootCustomControlElementSearch As pSearch
Private OpenCloseNavigationMenuButtonElementSearch As pSearch
Private NavigationMenuRootPaneWindowElementSearch As pSearch
Private CurrentCalculatorLandmarkGroupControlElementSearch As pSearch

Private Type PrivateElementNames
  MasterWindow As String
  MainCalculatorSubWindow As String
  NavigationViewRootCustomControl As String
  OpenCloseNavigationMenuButton As String
  NavigationMenuRootPaneWindow As String
  CurrentCalculatorLandmarkGroupControl As String
End Type

Private This As PrivateElementNames
Private PrivateElements As New Scripting.Dictionary

Private Sub Class_Initialize()
  Factory.GetRootDesktopElement
  OpenCalculator
  InitializeAllElements
  FindMasterWindowElement
  FindNavigationElements
End Sub

Private Sub Class_Terminate()
  Window.CloseWindow This.MasterWindow, PrivateElements(This.MasterWindow)
  Set MasterWindowElementSearch = Nothing
  Set MainCalculatorSubWindowElementSearch = Nothing
  Set NavigationViewRootCustomControlElementSearch = Nothing
  Set OpenCloseNavigationMenuButtonElementSearch = Nothing
  Set NavigationMenuRootPaneWindowElementSearch = Nothing
  Set CurrentCalculatorLandmarkGroupControlElementSearch = Nothing
  Set PrivateElements = Nothing
End Sub

Private Sub OpenCalculator()
  WindowsProcesses.RunShellExecuteToStartNewProcess "Microsoft Windows Calculator", "open", "shell:appsFolder\Microsoft.WindowsCalculator_8wekyb3d8bbwe!App", VBA.Constants.vbNullString, VBA.Constants.vbNullString, WindowStyle.Normal
End Sub

Private Sub InitializeAllElements()
  
  Set MasterWindowElementSearch = Factory.GetNewSearch
  Set MainCalculatorSubWindowElementSearch = Factory.GetNewSearch
  Set NavigationViewRootCustomControlElementSearch = Factory.GetNewSearch
  Set OpenCloseNavigationMenuButtonElementSearch = Factory.GetNewSearch
  Set NavigationMenuRootPaneWindowElementSearch = Factory.GetNewSearch
  Set CurrentCalculatorLandmarkGroupControlElementSearch = Factory.GetNewSearch
  
  This.MasterWindow = "MasterWindow"
  This.MainCalculatorSubWindow = "MainCalculatorSubWindow"
  This.NavigationViewRootCustomControl = "NavigationView"
  This.OpenCloseNavigationMenuButton = "OpenCloseNavigationMenuButton"
  This.NavigationMenuRootPaneWindow = "NavigationMenuRootPaneWindow"
  This.CurrentCalculatorLandmarkGroupControl = "CurrentCalculatorLandmarkGroupControlElement"

End Sub

Private Sub FindMasterWindowElement()
  
  With MasterWindowElementSearch
    .Initialise This.MasterWindow, Factory.RootDesktopUIAElement, TreeScope.Children, By.pConditions, "AND(NameIsCalculator, ControlTypeIsWindow, ClassNameIsApplicationFrameWindow, WindowInteractionStateIsReadyForUserInteraction)"
    .Condition "NameIsCalculator", UIAProperties.Name, UIAPropertyComparisons.Equals, "Calculator"
    .Condition "ControlTypeIsWindow", UIAProperties.ControlType, UIAPropertyComparisons.Equals, pEssence.UIAControlTypeIDs.Window
    .Condition "ClassNameIsApplicationFrameWindow", UIAProperties.ClassName, UIAPropertyComparisons.Equals, "ApplicationFrameWindow"
    .Condition "WindowInteractionStateIsReadyForUserInteraction", UIAProperties.WindowWindowInteractionState, UIAPropertyComparisons.Equals, UIAWindowInteractionStates.ReadyForUserInteraction
    PrivateElements.Add This.MasterWindow, .Find(10)
  End With
      
End Sub

Private Sub FindNavigationElements()
  
  With MainCalculatorSubWindowElementSearch
    .Initialise This.MainCalculatorSubWindow, PrivateElements(This.MasterWindow), TreeScope.Children, By.pConditions, "AND(NameIsCalculator, ControlTypeIsWindow, ClassNameIsWindowsUICoreCoreWindow)"
    .Condition "NameIsCalculator", UIAProperties.Name, UIAPropertyComparisons.Equals, "Calculator"
    .Condition "ControlTypeIsWindow", UIAProperties.ControlType, UIAPropertyComparisons.Equals, pEssence.UIAControlTypeIDs.Window
    .Condition "ClassNameIsWindowsUICoreCoreWindow", UIAProperties.ClassName, UIAPropertyComparisons.Equals, "Windows.UI.Core.CoreWindow"
    PrivateElements.Add This.MainCalculatorSubWindow, .Find(10)
  End With
  
  With NavigationViewRootCustomControlElementSearch
    .Initialise This.NavigationViewRootCustomControl, PrivateElements(This.MainCalculatorSubWindow), pEssence.TreeScope.Children, By.AutomationId, "NavView"
    PrivateElements.Add This.NavigationViewRootCustomControl, .Find(10)
  End With
    
  With OpenCloseNavigationMenuButtonElementSearch
    .Initialise This.OpenCloseNavigationMenuButton, PrivateElements(This.NavigationViewRootCustomControl), TreeScope.Children, By.pConditions, "AND(OR(NameIsOpenNavigation,NameIsCloseNavigation), ControlTypeIsButton, ClassNameIsButton)"
    .Condition "NameIsOpenNavigation", UIAProperties.Name, UIAPropertyComparisons.Equals, "Open Navigation"
    .Condition "NameIsCloseNavigation", UIAProperties.Name, UIAPropertyComparisons.Equals, "Close Navigation"
    .Condition "ControlTypeIsButton", UIAProperties.ControlType, UIAPropertyComparisons.Equals, UIAControlTypeIDs.Button
    .Condition "ClassNameIsButton", UIAProperties.ClassName, UIAPropertyComparisons.Equals, "Button"
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
      .Initialise This.NavigationMenuRootPaneWindow, PrivateElements(This.NavigationViewRootCustomControl), TreeScope.Children, By.pConditions, "AutomationIdIsPaneRoot"
      .Condition "AutomationIdIsPaneRoot", UIAProperties.AutomationId, UIAPropertyComparisons.Equals, "PaneRoot"
     Else
      PrivateElements.Remove This.NavigationMenuRootPaneWindow
     End If
    PrivateElements.Add This.NavigationMenuRootPaneWindow, .Find
  End With

  Dim CurrentNavigationMenuItemElementSearch As pSearch
  Set CurrentNavigationMenuItemElementSearch = Nothing
  Set CurrentNavigationMenuItemElementSearch = Factory.GetNewSearch
  
  With CurrentNavigationMenuItemElementSearch
    .Initialise CalculatorType, PrivateElements(This.NavigationMenuRootPaneWindow), TreeScope.Descendants, By.pConditions, "AND(NameIs, ControlTypeIsListItem)"
    .Condition "NameIs", UIAProperties.Name, UIAPropertyComparisons.Equals, CalculatorType
    .Condition "ControlTypeIsListItem", UIAProperties.ControlType, UIAPropertyComparisons.Equals, UIAControlTypeIDs.ListItem
    Set GetMenuElement = .Find(10)
  End With

End Function

Public Function GetCurrentCalculatorLandmarkElement() As IUIAutomationElement
  
  'We need to find the NavViewRoot Element again
  PrivateElements(This.NavigationViewRootCustomControl) = NavigationViewRootCustomControlElementSearch.Find(10)
  
  With CurrentCalculatorLandmarkGroupControlElementSearch
    .Initialise This.CurrentCalculatorLandmarkGroupControl, PrivateElements(This.NavigationViewRootCustomControl), TreeScope.Children, By.pConditions, "AND(ControlTypeIsGroup, ClassNameIsLandmarkTarget)"
    .Condition "ControlTypeIsGroup", UIAProperties.ControlType, UIAPropertyComparisons.Equals, UIAControlTypeIDs.Group
    .Condition "ClassNameIsLandmarkTarget", UIAProperties.ClassName, UIAPropertyComparisons.Equals, "LandmarkTarget"
    PrivateElements.Add This.CurrentCalculatorLandmarkGroupControl, .Find(10)
  End With
  Set GetCurrentCalculatorLandmarkElement = PrivateElements(This.CurrentCalculatorLandmarkGroupControl)
  
End Function


