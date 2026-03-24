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
  MasterWindowSearch As pSearch
  MainCalculatorSubWindowSearch As pSearch
  NavigationViewRootCustomControlSearch As pSearch
  OpenCloseNavigationMenuButtonSearch As pSearch
  NavigationMenuRootPaneWindowSearch As pSearch
  CurrentCalculatorLandmarkGroupControlSearch As pSearch
End Type

Private This As PageAttributes

Private Sub Class_Initialize()
  Factory.GetRootDesktopElement
  OpenCalculator
  InitializeAllElements
  FindMasterWindowElement
  FindNavigationElements
End Sub

Private Sub Class_Terminate()
  Window.CloseWindow This.MasterWindowSearch.ElementName, This.MasterWindowSearch.FoundUIAElement
  Set This.MasterWindowSearch = Nothing
  Set This.MainCalculatorSubWindowSearch = Nothing
  Set This.NavigationViewRootCustomControlSearch = Nothing
  Set This.OpenCloseNavigationMenuButtonSearch = Nothing
  Set This.NavigationMenuRootPaneWindowSearch = Nothing
  Set This.CurrentCalculatorLandmarkGroupControlSearch = Nothing
End Sub

Private Sub OpenCalculator()
  WindowsProcesses.RunShellExecuteToStartNewProcess "Microsoft Windows Calculator", "open", "shell:appsFolder\Microsoft.WindowsCalculator_8wekyb3d8bbwe!App", VBA.Constants.vbNullString, VBA.Constants.vbNullString, WindowStyle.Normal
End Sub

Private Sub InitializeAllElements()
  Set This.MasterWindowSearch = Factory.GetNewSearch
  Set This.MainCalculatorSubWindowSearch = Factory.GetNewSearch
  Set This.NavigationViewRootCustomControlSearch = Factory.GetNewSearch
  Set This.OpenCloseNavigationMenuButtonSearch = Factory.GetNewSearch
  Set This.NavigationMenuRootPaneWindowSearch = Factory.GetNewSearch
  Set This.CurrentCalculatorLandmarkGroupControlSearch = Factory.GetNewSearch
End Sub

Private Sub FindMasterWindowElement()
  
  With This.MasterWindowSearch
    .Initialise "MasterWindow", RootDesktopUIAElement, Children, pConditions, "AND(NameIs, ControlType, ClassName, WindowInteractionState)"
    'We can't use just 'Name' as this appears in 'ClassName'!
    .Condition "NameIs", Name, IsTheString, "Calculator"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Window
    .Condition "ClassName", ClassName, IsTheString, "ApplicationFrameWindow"
    .Condition "WindowInteractionState", WindowWindowInteractionState, EqualsNumber, UIAWindowInteractionStates.ReadyForUserInteraction
    .Find (10)
  End With
      
End Sub

Private Sub FindNavigationElements()
  
  With This.MainCalculatorSubWindowSearch
    .Initialise "MainCalculatorSubWindow", This.MasterWindowSearch.FoundUIAElement, Children, pConditions, "AND(NameIs, ControlType, ClassName)"
    .Condition "NameIs", Name, IsTheString, "Calculator"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Window
    .Condition "ClassName", ClassName, IsTheString, "Windows.UI.Core.CoreWindow"
    .Find (10)
  End With

  With This.NavigationViewRootCustomControlSearch
    .Initialise "NavigationView", This.MainCalculatorSubWindowSearch.FoundUIAElement, Children, By.AutomationId, "NavView"
    .Find (10)
  End With
    
  With This.OpenCloseNavigationMenuButtonSearch
    .Initialise "OpenCloseNavigationMenuButton", This.NavigationViewRootCustomControlSearch.FoundUIAElement, Children, pConditions, _
      "AND(OR(NameIsOpen,NameIsClose), ControlType, ClassName)"
    .Condition "NameIsOpen", Name, IsTheString, "Open Navigation"
    .Condition "NameIsClose", Name, IsTheString, "Close Navigation"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Button
    .Condition "ClassName", ClassName, IsTheString, "Button"
    .Find (10)
  End With
        
End Sub

Public Sub SelectCalculatorType(CalculatorType As String)
  OpenNavigationMenu
  Actions.Click CalculatorType, GetMenuElement(CalculatorType)
End Sub

Private Sub OpenNavigationMenu()
  Actions.Click This.OpenCloseNavigationMenuButtonSearch.ElementName, This.OpenCloseNavigationMenuButtonSearch.FoundUIAElement
End Sub

Private Function GetMenuElement(CalculatorType As String) As IUIAutomationElement
  
  'We need to find this element each time the menu is opened
  With This.NavigationMenuRootPaneWindowSearch
     'Has it been initialised yet?
     If This.NavigationMenuRootPaneWindowSearch.ElementName = "" Then
      .Initialise "NavigationMenuRootPaneWindow", This.NavigationViewRootCustomControlSearch.FoundUIAElement, Children, By.AutomationId, "PaneRoot"
     End If
     .Find
  End With

  Dim CurrentNavigationMenuItemElementSearch As pSearch
  Set CurrentNavigationMenuItemElementSearch = Factory.GetNewSearch
  With CurrentNavigationMenuItemElementSearch
    .Initialise CalculatorType, This.NavigationMenuRootPaneWindowSearch.FoundUIAElement, Descendants, pConditions, "AND(Name, ControlType)"
    .Condition "Name", Name, IsTheString, CalculatorType
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.ListItem
    .Find (10)
    Set GetMenuElement = CurrentNavigationMenuItemElementSearch.FoundUIAElement
  End With
  Set CurrentNavigationMenuItemElementSearch = Nothing

End Function

Public Function GetCurrentCalculatorLandmarkElement() As IUIAutomationElement
  
  'We need to find the NavViewRoot Element again for the new current calculator
  This.NavigationViewRootCustomControlSearch.Find (10)
  
  With This.CurrentCalculatorLandmarkGroupControlSearch
    .Initialise "CurrentCalculatorLandmarkGroupControl", This.NavigationViewRootCustomControlSearch.FoundUIAElement, Children, pConditions, _
      "AND(ControlType, ClassName)"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Group
    .Condition "ClassName", ClassName, IsTheString, "LandmarkTarget"
    .Find (10)
  End With
  Set GetCurrentCalculatorLandmarkElement = This.CurrentCalculatorLandmarkGroupControlSearch.FoundUIAElement
  
End Function
