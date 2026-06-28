VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LetCodeDotInPageRadio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder LetCodeDotIn
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Const WEB_APP_NAME = "Letcode"
Const TARGET_PAGE_URL = "https://letcode.in/radio"
Const TARGET_PAGE_TITLE = "Radio Buttons | LetCode with Koushik"

Private Type PageAttributes
  WebBrowser As Object
  RootWebArea As pLocator
  AllControls As pLocator
  RadioAndCheckboxHeader As pElement
  SelectAnyOneSubHeader As pElement
  SelectAnyOneYes As pElement
  SelectAnyOneNo As pElement
  ConfirmYouCanSelectOnlyOneRadioButton As pElement
  ConfirmYouCanSelectOnlyOneRadioButtonYes As pElement
  ConfirmYouCanSelectOnlyOneRadioButtonNo As pElement
  FindTheBug As pElement
  FindTheBugYes As pElement
  FindTheBugNo As pElement
  FindWhichOneIsSelected As pElement
  FindWhichOneIsSelectedFoo As pElement
  FindWhichOneIsSelectedBar As pElement
  ConfirmLastFieldIsDisabled As pElement
  ConfirmLastFieldIsDisabledGoing As pElement
  ConfirmLastFieldIsDisabledNotGoing As pElement
  ConfirmLastFieldIsDisabledMaybe As pElement
  FindIfTheCheckboxIsSelected As pElement
  FindIfTheCheckboxIsSelectedRememberMe As pElement
  AcceptTheTermsAndConditions As pElement
  IAgreeToTheFAKETermsAndConditions As pElement
End Type

Private This As PageAttributes

Private Sub Class_Initialize()
  Set This.WebBrowser = Factory.GetNewWebBrowser
  Set This.AllControls = Factory.GetNewLocator
End Sub

Private Sub Class_Terminate()
  Set This.WebBrowser = Nothing
  Set This.RootWebArea = Nothing
  Set This.AllControls = Nothing
  Set This.RadioAndCheckboxHeader = Nothing
  Set This.SelectAnyOneSubHeader = Nothing
  Set This.SelectAnyOneYes = Nothing
  Set This.SelectAnyOneNo = Nothing
  Set This.ConfirmYouCanSelectOnlyOneRadioButton = Nothing
  Set This.ConfirmYouCanSelectOnlyOneRadioButtonYes = Nothing
  Set This.ConfirmYouCanSelectOnlyOneRadioButtonNo = Nothing
  Set This.FindTheBug = Nothing
  Set This.FindTheBugYes = Nothing
  Set This.FindTheBugNo = Nothing
  Set This.FindWhichOneIsSelected = Nothing
  Set This.FindWhichOneIsSelectedFoo = Nothing
  Set This.FindWhichOneIsSelectedBar = Nothing
  Set This.ConfirmLastFieldIsDisabled = Nothing
  Set This.ConfirmLastFieldIsDisabledGoing = Nothing
  Set This.ConfirmLastFieldIsDisabledNotGoing = Nothing
  Set This.ConfirmLastFieldIsDisabledMaybe = Nothing
  Set This.FindIfTheCheckboxIsSelected = Nothing
  Set This.FindIfTheCheckboxIsSelectedRememberMe = Nothing
  Set This.AcceptTheTermsAndConditions = Nothing
  Set This.IAgreeToTheFAKETermsAndConditions = Nothing
End Sub

Public Sub Initialize()

  With This.WebBrowser
    .Start WEB_APP_NAME, TARGET_PAGE_URL, TARGET_PAGE_TITLE
    Set This.RootWebArea = .GetRootWebArea
  End With

    
  Dim AllControls() As pElement
  With This.AllControls
    .Initialise "AllControls", This.RootWebArea, Descendants, pConditions, _
      "AND" & _
        "(" & _
          "OR(AriaRoleHeading, AriaRoleDescription, AriaRoleRadio, AriaRoleCheckBox, ClassName, AND(ControlTypeGroup, AriaRoleNullString, GroupClassForFirefox))," & _
          "NOT(AND(AriaRoleDescription, OR(NameIsProducts, NameIsGrooming, NameIsRadioAndCheckbox, NameIsAndCheckbox, NameIsYes, NameIsNo!, Notice1, Notice2, NameIsFoo, NameIsBar, NameIsGoing, NameIsNotGoing, NameIsMaybe, NameIsRememberMe)))," & _
          "NOT(AND(AriaRoleHeading, OR(NameIsTopics, NameIsLearningPoints)))" & _
         ")"
'    .Initialise "AllControls", This.RootWebArea, Descendants, pConditions, _
'      "AND" & _
'        "(" & _
'          "OR(AriaRoleHeading, AriaRoleDescription, AriaRoleRadio, AriaRoleCheckBox, ClassName)," & _
'          "NOT(AND(AriaRoleDescription, OR(NameIsProducts, NameIsGrooming, NameIsRadioAndCheckbox, NameIsAndCheckbox, NameIsYes, NameIsNo!, Notice1, Notice2, NameIsFoo, NameIsBar, NameIsGoing, NameIsNotGoing, NameIsMaybe, NameIsRememberMe)))," & _
'          "NOT(AND(AriaRoleHeading, NameIsTopics))" & _
'         ")"
    .AriaRoleHeading: .AriaRoleDescription: .AriaRoleRadio: .AriaRoleCheckBox: .ClassName "label" 'Firefox uses class name 'labels'
    .Condition "ControlTypeGroup", UIAProperties.ControlType, UIAPropertyComparisons.EqualsNumber, UIAControlTypeIDs.Group
'    .AriaRoleGeneric: .AriaRoleMain: .AriaRoleNavigation
    .AriaRoleNullString: .Condition "GroupClassForFirefox", UIAProperties.ClassName, IsTheString, "block text-sm font-semibold mb-2 text-slate-800 dark:text-slate-200"
    .NameIs_ "Products"
    .NameIs_ "Grooming"
    .NameIs_ " & Checkbox", "AndCheckbox" 'Yandex!
    .NameIs_ "Radio & Checkbox", "RadioAndCheckbox" 'Yandex!
    .NameIs_ "Yes", , True: .NameIs_ "No", "No!", True
    .Condition "Notice1", Name, StartsWithTheString, "* Notice: both buttons can be active due to non-matching "
    .Condition "Notice2", Name, EndsWithTheString, " names (intentional bug)."
    .NameIs_ "Foo", , True: .NameIs_ "Bar", , True
    .NameIs_ "Going", , True: .NameIs_ "Not going", "NotGoing", True: .NameIs_ "Maybe", , True
    .NameIs_ "Remember me", "RememberMe", True
    .NameIs_ "These are topics related to the article that might interest you", "Topics"
    .NameIs_ "Learning Points", "LearningPoints"
    .FindAll False
    AllControls = .Elements
  End With

  If Not pEssence.Utils.IsArrayEmpty(AllControls) Then
    Dim i As Integer
    Dim CurrentUIAElement As IUIAutomationElement
    For i = 0 To UBound(AllControls)
      Set CurrentUIAElement = AllControls(i).UIAElement
      With CurrentUIAElement
        Select Case i + 1
          Case 1
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Heading
            Debug.Assert CurrentUIAElement.CurrentName = "Radio & Checkbox"
            Set This.RadioAndCheckboxHeader = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          Case 2
            Debug.Assert (CurrentUIAElement.CurrentAriaRole = AriaRoles.Description) Or (CurrentUIAElement.CurrentClassName = "label") Or (CurrentUIAElement.CurrentClassName = "block text-sm font-semibold mb-2 text-slate-800 dark:text-slate-200")
            Debug.Assert CurrentUIAElement.CurrentName = "Select any one"
            Set This.SelectAnyOneSubHeader = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          Case 3
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Radio
            Debug.Assert CurrentUIAElement.CurrentName = "Yes"
            Set This.SelectAnyOneYes = Factory.GetNewElement(This.SelectAnyOneSubHeader.GivenName & " - Yes", CurrentUIAElement)
          Case 4
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Radio
            Debug.Assert CurrentUIAElement.CurrentName = "No"
            Set This.SelectAnyOneNo = Factory.GetNewElement(This.SelectAnyOneSubHeader.GivenName & " - No", CurrentUIAElement)
          Case 5
            Debug.Assert (CurrentUIAElement.CurrentAriaRole = AriaRoles.Description) Or (CurrentUIAElement.CurrentClassName = "label") Or (CurrentUIAElement.CurrentClassName = "block text-sm font-semibold mb-2 text-slate-800 dark:text-slate-200")
            Debug.Assert CurrentUIAElement.CurrentName = "Confirm you can select only one radio button"
            Set This.ConfirmYouCanSelectOnlyOneRadioButton = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          Case 6
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Radio
            Debug.Assert CurrentUIAElement.CurrentName = "Yes"
            Set This.ConfirmYouCanSelectOnlyOneRadioButtonYes = Factory.GetNewElement(This.ConfirmYouCanSelectOnlyOneRadioButton.GivenName & " - Yes", CurrentUIAElement)
          Case 7
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Radio
            Debug.Assert CurrentUIAElement.CurrentName = "No"
            Set This.ConfirmYouCanSelectOnlyOneRadioButtonNo = Factory.GetNewElement(This.ConfirmYouCanSelectOnlyOneRadioButton.GivenName & " - No", CurrentUIAElement)
          Case 8
            Debug.Assert (CurrentUIAElement.CurrentAriaRole = AriaRoles.Description) Or (CurrentUIAElement.CurrentClassName = "label") Or (CurrentUIAElement.CurrentClassName = "block text-sm font-semibold mb-2 text-slate-800 dark:text-slate-200")
            Debug.Assert CurrentUIAElement.CurrentName = "Find the bug"
            Set This.FindTheBug = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          Case 9
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Radio
            Debug.Assert CurrentUIAElement.CurrentName = "Yes"
            Set This.FindTheBugYes = Factory.GetNewElement(This.FindTheBug.GivenName & " - Yes", CurrentUIAElement)
          Case 10
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Radio
            Debug.Assert CurrentUIAElement.CurrentName = "No"
            Set This.FindTheBugNo = Factory.GetNewElement(This.FindTheBug.GivenName & " - No", CurrentUIAElement)
          Case 11
            Debug.Assert (CurrentUIAElement.CurrentAriaRole = AriaRoles.Description) Or (CurrentUIAElement.CurrentClassName = "label") Or (CurrentUIAElement.CurrentClassName = "block text-sm font-semibold mb-2 text-slate-800 dark:text-slate-200")
            Debug.Assert CurrentUIAElement.CurrentName = "Find which one is selected"
            Set This.FindWhichOneIsSelected = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          Case 12
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Radio
            Debug.Assert CurrentUIAElement.CurrentName = "Foo"
            Set This.FindWhichOneIsSelectedFoo = Factory.GetNewElement(This.FindWhichOneIsSelected.GivenName & " - Foo", CurrentUIAElement)
          Case 13
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Radio
            Debug.Assert CurrentUIAElement.CurrentName = "Bar"
            Set This.FindWhichOneIsSelectedBar = Factory.GetNewElement(This.FindWhichOneIsSelected.GivenName & " - Bar", CurrentUIAElement)
          Case 14
            Debug.Assert (CurrentUIAElement.CurrentAriaRole = AriaRoles.Description) Or (CurrentUIAElement.CurrentClassName = "label") Or (CurrentUIAElement.CurrentClassName = "block text-sm font-semibold mb-2 text-slate-800 dark:text-slate-200")
            Debug.Assert CurrentUIAElement.CurrentName = "Confirm last field is disabled"
            Set This.ConfirmLastFieldIsDisabled = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          Case 15
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Radio
            Debug.Assert CurrentUIAElement.CurrentName = "Going"
            Set This.ConfirmLastFieldIsDisabledGoing = Factory.GetNewElement(This.ConfirmLastFieldIsDisabled.GivenName & " - Going", CurrentUIAElement)
          Case 16
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Radio
            Debug.Assert CurrentUIAElement.CurrentName = "Not going"
            Set This.ConfirmLastFieldIsDisabledNotGoing = Factory.GetNewElement(This.ConfirmLastFieldIsDisabled.GivenName & " - Not going", CurrentUIAElement)
          Case 17
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Radio
            Debug.Assert CurrentUIAElement.CurrentName = "Maybe"
            Set This.ConfirmLastFieldIsDisabledMaybe = Factory.GetNewElement(This.ConfirmLastFieldIsDisabled.GivenName & " - Maybe", CurrentUIAElement)
          Case 18
            Debug.Assert (CurrentUIAElement.CurrentAriaRole = AriaRoles.Description) Or (CurrentUIAElement.CurrentClassName = "label") Or (CurrentUIAElement.CurrentClassName = "block text-sm font-semibold mb-2 text-slate-800 dark:text-slate-200")
            Debug.Assert CurrentUIAElement.CurrentName = "Find if the checkbox is selected?"
            Set This.FindIfTheCheckboxIsSelected = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          Case 19
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.CheckBox
            Debug.Assert CurrentUIAElement.CurrentName = "Remember me"
            Set This.FindIfTheCheckboxIsSelectedRememberMe = Factory.GetNewElement(This.FindIfTheCheckboxIsSelected.GivenName & " - Remember me", CurrentUIAElement)
          Case 20
            Debug.Assert (CurrentUIAElement.CurrentAriaRole = AriaRoles.Description) Or (CurrentUIAElement.CurrentClassName = "label") Or (CurrentUIAElement.CurrentClassName = "block text-sm font-semibold mb-2 text-slate-800 dark:text-slate-200")
            Debug.Assert CurrentUIAElement.CurrentName = "Accept the T&C"
            Set This.AcceptTheTermsAndConditions = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          Case 21
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.CheckBox
            Debug.Assert CurrentUIAElement.CurrentName = "I agree to the FAKE terms and conditions"
            Set This.IAgreeToTheFAKETermsAndConditions = Factory.GetNewElement(CurrentUIAElement.CurrentName, CurrentUIAElement)
          Case 22
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Description
            Debug.Assert VBA.Strings.Trim(CurrentUIAElement.CurrentName) = "I agree to the"
          Case 23
            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Description
            Debug.Assert CurrentUIAElement.CurrentName = "FAKE terms and conditions"
'          Case 24
'            Debug.Assert CurrentUIAElement.CurrentAriaRole = AriaRoles.Description
'            Debug.Assert CurrentUIAElement.CurrentName = "On completion of this exercise, you can learn the following concepts."
          Case Else
            'Just ignore any other buttons!
'             MsgBox "Error - button not handed!? #" & i & ": '" & CurrentUIAElement.CurrentName & "'"
        End Select
      End With
    Next i
  End If

End Sub

Public Sub Automate()
  
  'SelectAnyOne
  Debug.Assert This.SelectAnyOneYes.IsEnabled
  Debug.Assert This.SelectAnyOneYes.IsSelected()
  Debug.Assert This.SelectAnyOneNo.IsEnabled
  Debug.Assert Not This.SelectAnyOneNo.IsSelected()
  
  This.SelectAnyOneYes.Click
  This.SelectAnyOneYes.WaitForPatternState UIAPatterns.SelectionItem, "CurrentIsSelected", 1
  This.SelectAnyOneNo.WaitForPatternState UIAPatterns.SelectionItem, "CurrentIsNotSelected", 0
  Debug.Assert This.SelectAnyOneYes.IsSelected()
  Debug.Assert Not This.SelectAnyOneNo.IsSelected()
  
  This.SelectAnyOneNo.Click
  This.SelectAnyOneNo.WaitForPatternState UIAPatterns.SelectionItem, "CurrentIsSelected", 1
  This.SelectAnyOneYes.WaitForPatternState UIAPatterns.SelectionItem, "CurrentIsNotSelected", 0
  Debug.Assert This.SelectAnyOneNo.IsSelected()
  Debug.Assert Not This.SelectAnyOneYes.IsSelected()
  
  'CofirmYouCanSelectOnlyOneRadioButton
  Debug.Assert This.ConfirmYouCanSelectOnlyOneRadioButtonYes.IsEnabled
  Debug.Assert Not This.ConfirmYouCanSelectOnlyOneRadioButtonYes.IsSelected()
  Debug.Assert This.ConfirmYouCanSelectOnlyOneRadioButtonNo.IsEnabled
  Debug.Assert Not This.ConfirmYouCanSelectOnlyOneRadioButtonNo.IsSelected()
  
  This.ConfirmYouCanSelectOnlyOneRadioButtonYes.Click
  This.ConfirmYouCanSelectOnlyOneRadioButtonYes.WaitForPatternState UIAPatterns.SelectionItem, "CurrentIsSelected", 1
  This.ConfirmYouCanSelectOnlyOneRadioButtonNo.WaitForPatternState UIAPatterns.SelectionItem, "CurrentIsNotSelected", 0
  Debug.Assert This.ConfirmYouCanSelectOnlyOneRadioButtonYes.IsSelected()
  Debug.Assert Not This.ConfirmYouCanSelectOnlyOneRadioButtonNo.IsSelected()
  
  This.ConfirmYouCanSelectOnlyOneRadioButtonNo.Click
  This.ConfirmYouCanSelectOnlyOneRadioButtonNo.WaitForPatternState UIAPatterns.SelectionItem, "CurrentIsSelected", 1
  This.ConfirmYouCanSelectOnlyOneRadioButtonYes.WaitForPatternState UIAPatterns.SelectionItem, "CurrentIsNotSelected", 0
  Debug.Assert This.ConfirmYouCanSelectOnlyOneRadioButtonNo.IsSelected()
  Debug.Assert Not This.ConfirmYouCanSelectOnlyOneRadioButtonYes.IsSelected()

  'FindTheBug
  Debug.Assert This.FindTheBugYes.IsEnabled
  Debug.Assert Not This.FindTheBugYes.IsSelected()
  Debug.Assert This.FindTheBugNo.IsEnabled
  Debug.Assert Not This.FindTheBugNo.IsSelected()
  
  This.FindTheBugYes.Click
  This.FindTheBugYes.WaitForPatternState UIAPatterns.SelectionItem, "CurrentIsSelected", 1
  This.FindTheBugNo.WaitForPatternState UIAPatterns.SelectionItem, "CurrentIsNotSelected", 0
  Debug.Assert This.FindTheBugYes.IsSelected()
  Debug.Assert Not This.FindTheBugNo.IsSelected()
  
  This.FindTheBugNo.Click
  This.FindTheBugYes.WaitForPatternState UIAPatterns.SelectionItem, "CurrentIsSelected", 0
  This.FindTheBugNo.WaitForPatternState UIAPatterns.SelectionItem, "CurrentIsSelected", 1
  Debug.Assert This.FindTheBugNo.IsSelected()
  Debug.Assert This.FindTheBugYes.IsSelected() 'BUG: The Yes radio button should NOT still be selected!

  'FindWhichOneIsSelected
  If This.FindWhichOneIsSelectedFoo.IsSelected() Then
    Debug.Assert This.FindWhichOneIsSelectedFoo.IsSelected()
    Debug.Assert Not This.FindWhichOneIsSelectedBar.IsSelected()
  ElseIf This.FindWhichOneIsSelectedBar.IsSelected() Then
    Debug.Assert This.FindWhichOneIsSelectedBar.IsSelected()
    Debug.Assert Not This.FindWhichOneIsSelectedFoo.IsSelected()
  End If
  
  'ConfirmLastFieldIsDisabled
  Debug.Assert This.ConfirmLastFieldIsDisabledGoing.IsEnabled
  Debug.Assert This.ConfirmLastFieldIsDisabledNotGoing.IsEnabled
  Debug.Assert Not This.ConfirmLastFieldIsDisabledMaybe.IsEnabled
  
  'FindIfTheCheckboxIsSelected
  Debug.Assert This.FindIfTheCheckboxIsSelectedRememberMe.IsSelected
  This.FindIfTheCheckboxIsSelectedRememberMe.Click
  This.FindIfTheCheckboxIsSelectedRememberMe.WaitForPatternState UIAPatterns.Toggle, "CurrentToggleStateOff", 1
  Debug.Assert Not This.FindIfTheCheckboxIsSelectedRememberMe.IsSelected
  
  'AcceptTheTermsAndConditions
  Debug.Assert Not This.IAgreeToTheFAKETermsAndConditions.IsSelected
  This.IAgreeToTheFAKETermsAndConditions.Click
  This.IAgreeToTheFAKETermsAndConditions.WaitForPatternState UIAPatterns.Toggle, "CurrentToggleStateOn", 1
  Debug.Assert This.IAgreeToTheFAKETermsAndConditions.IsSelected

End Sub
