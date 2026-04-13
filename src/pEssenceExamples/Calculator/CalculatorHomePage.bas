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

Private Type PageAttributes
  MasterWindow As pLocator
  MainCalculatorSubWindow As pLocator
  NavigationViewRootCustomControl As pLocator
  OpenCloseNavigationMenuButton As pLocator
  NavigationMenuRootPaneWindow As pLocator
  CurrentCalculatorLandmarkGroupControl As pLocator
End Type

Private This As PageAttributes

Private Sub Class_Initialize()
  OpenCalculator
  InitialiseAllLocators
End Sub

Private Sub Class_Terminate()
  This.MasterWindow.Element.CloseWindow
  Set This.MasterWindow = Nothing
  Set This.MainCalculatorSubWindow = Nothing
  Set This.NavigationViewRootCustomControl = Nothing
  Set This.OpenCloseNavigationMenuButton = Nothing
  Set This.NavigationMenuRootPaneWindow = Nothing
  Set This.CurrentCalculatorLandmarkGroupControl = Nothing
End Sub

Private Sub OpenCalculator()
  WindowsProcesses.RunShellExecuteToStartNewProcess "Microsoft Windows Calculator", "open", "shell:appsFolder\Microsoft.WindowsCalculator_8wekyb3d8bbwe!App", VBA.Constants.vbNullString, VBA.Constants.vbNullString, WindowStyle.Normal
End Sub

Private Sub InitialiseAllLocators()
  
  Set This.MasterWindow = Factory.GetNewLocator
  Set This.MainCalculatorSubWindow = Factory.GetNewLocator
  Set This.NavigationViewRootCustomControl = Factory.GetNewLocator
  Set This.OpenCloseNavigationMenuButton = Factory.GetNewLocator
  Set This.NavigationMenuRootPaneWindow = Factory.GetNewLocator
  Set This.CurrentCalculatorLandmarkGroupControl = Factory.GetNewLocator
  
  With This.MasterWindow
    .Initialise "MasterWindow", Nothing, Children, pConditions, "AND(NameIs, ControlType, ClassName, WindowInteractionState)"
    'We can't use just 'Name' as this appears in 'ClassName'!
    .Condition "NameIs", Name, IsTheString, "Calculator"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Window
    .Condition "ClassName", ClassName, IsTheString, "ApplicationFrameWindow"
    .Condition "WindowInteractionState", WindowWindowInteractionState, EqualsNumber, UIAWindowInteractionStates.ReadyForUserInteraction
  End With
        
  With This.MainCalculatorSubWindow
    .Initialise "MainCalculatorSubWindow", This.MasterWindow, Children, pConditions, "AND(NameIs, ControlType, ClassName)"
    .Condition "NameIs", Name, IsTheString, "Calculator"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Window
    .Condition "ClassName", ClassName, IsTheString, "Windows.UI.Core.CoreWindow"
  End With

  With This.NavigationViewRootCustomControl
    .Initialise "NavigationView", This.MainCalculatorSubWindow, Children, By.AutomationId, "NavView"
  End With
    
  With This.OpenCloseNavigationMenuButton
    .Initialise "OpenCloseNavigationMenuButton", This.NavigationViewRootCustomControl, Children, pConditions, _
      "AND(OR(NameIsOpen,NameIsClose), ControlType, ClassName)"
    .Condition "NameIsOpen", Name, IsTheString, "Open Navigation"
    .Condition "NameIsClose", Name, IsTheString, "Close Navigation"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Button
    .Condition "ClassName", ClassName, IsTheString, "Button"
  End With
        
End Sub

Public Sub SelectCalculatorType(CalculatorType As String)
  Toaster.Message "Selecting " & CalculatorType
  OpenNavigationMenu
  Actions.Click GetMenuElement(CalculatorType)
End Sub

Private Sub OpenNavigationMenu()
  'Always find the element before we take action on it!
  This.OpenCloseNavigationMenuButton.Find 10
  Actions.Click This.OpenCloseNavigationMenuButton.Element
End Sub

Private Function GetMenuElement(CalculatorType As String) As pElement
  
  'We need to find this element each time the menu is opened
  With This.NavigationMenuRootPaneWindow
     'Has it been initialised yet?
     If Not This.NavigationMenuRootPaneWindow.Initialised Then
       .Initialise "NavigationMenuRootPaneWindow", This.NavigationViewRootCustomControl, Children, By.AutomationId, "PaneRoot"
     End If
  End With

  Dim CurrentNavigationMenuItemElementSearch As pLocator
  Set CurrentNavigationMenuItemElementSearch = Factory.GetNewLocator
  With CurrentNavigationMenuItemElementSearch
    .Initialise CalculatorType, This.NavigationMenuRootPaneWindow, Descendants, pConditions, "AND(Name, ControlType)"
    .Condition "Name", Name, IsTheString, CalculatorType
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.ListItem
    .Find (10) 'We need to find this before the next action on it!
    Set GetMenuElement = CurrentNavigationMenuItemElementSearch.Element
  End With
  Set CurrentNavigationMenuItemElementSearch = Nothing

End Function

Public Function GetCurrentCalculatorLandmarkGroupControl() As pLocator
  
  'We need to find the NavViewRoot Element again for the new current calculator
  This.NavigationViewRootCustomControl.Find (10)
  
  With This.CurrentCalculatorLandmarkGroupControl
    .Initialise "CurrentCalculatorLandmarkGroupControl", This.NavigationViewRootCustomControl, Children, pConditions, _
      "AND(ControlType, ClassName)"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Group
    .Condition "ClassName", ClassName, IsTheString, "LandmarkTarget"
  End With
  
  Set GetCurrentCalculatorLandmarkGroupControl = This.CurrentCalculatorLandmarkGroupControl
  
End Function
