Attribute VB_Name = "HTML_Selections"
'@Folder HTML
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2026 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public Function Selections() As Boolean
  
  LiveDebugging = False

  Dim Succeeded As Boolean
  Succeeded = False

  Dim WebBrowser As Object
  Dim SearchBox As pLocator
  Dim SearchButton As pLocator
  Dim SourceFileURL As String
  
  Dim DocumentRoot As pLocator
  
  Dim Header1 As String
  Dim SingleSelectComboBoxElement As pElement
  Dim SelectPlaywrightOption As pElement
  Dim SelectSeleniumOption As pElement
  Dim SelectCypressOption As pElement
  Dim SelectPhosphorusOption As pElement
  
  Dim Header2 As String
  Dim MultiSelectListBoxElement As pElement
  Dim MultiSelectPlaywrightOption As pElement
  Dim MultiSelectSeleniumOption As pElement
  Dim MultiSelectCypressOption As pElement
  Dim MultiSelectPhosphorusOption As pElement
  Dim SelectedItemsLabelElement As pElement
  Dim SelectedItemsEditElement As pElement

  On Error GoTo ErrorHandler
  
  If Not RunningAllExamples Then
    Factory.CurrentWebBrowserType = WebBrowserType.Brave
    Window.HighlightElements = False
  End If

  'UIA Dropdowns doesn't work in FireFox
  If Factory.CurrentWebBrowserType = WebBrowserType.Firefox Then
    GoTo ExitFunction
  End If

  SourceFileURL = HTMLCommon.GetULREncodedFilepath("Tests\WebPages\HTML\SingleAndMultipleSelections.html")
  
  Set WebBrowser = Factory.GetNewWebBrowser
  WebBrowser.Start "HTML Single And Multiple Selections", SourceFileURL, "Single/Multiple Selections"

  Set DocumentRoot = WebBrowser.GetRootWebArea

  'SINGLE-SELECT
  '=============
  
  'Check the header
  Dim Header1Text As pElement
  Header1 = "Single-Select"
  
  If Factory.CurrentWebBrowserType = WebBrowserType.Chromium Then
    'Chromium rolls all the text elements into one!
    Debug.Assert VBA.Strings.Replace(DocumentRoot.Element.Text, vbLf, "") = "Single-Select Choose Framework: Playwright Selenium Cypress Phosphorus   Multi-Select To Text Box Choose Frameworks (Hold Ctrl/Shift to select multiple): Playwright Selenium Cypress Phosphorus  Selected Items:   "
  Else
    Set Header1Text = DocumentRoot.FindElement("SingleSelectHeader", By.NameIs, Header1)
    Debug.Assert Header1Text.ControlType = UIAControlTypeIDs.Text
  End If

  'Check the Choose Framework text
  Dim FrameworkLabelSingle As pElement
  If Factory.CurrentWebBrowserType = WebBrowserType.Chromium Then
    Set FrameworkLabelSingle = DocumentRoot.FindElement("Choose Framework:", By.NameIs, "Choose Framework:", FindFirst:=True)
  Else
    Set FrameworkLabelSingle = Header1Text.FindSibling("Choose Framework:")
  End If
  Debug.Assert (FrameworkLabelSingle.ControlType = UIAControlTypeIDs.Text) Or (FrameworkLabelSingle.ControlType = UIAControlTypeIDs.Group) 'Firefox is Group!
  Debug.Assert FrameworkLabelSingle.Name = "Choose Framework:"
  
  'Check the combobox
  Set SingleSelectComboBoxElement = FrameworkLabelSingle.FindLabelledElement
  Debug.Assert (SingleSelectComboBoxElement.ControlType = UIAControlTypeIDs.ComboBox) Or (SingleSelectComboBoxElement.ControlType = UIAControlTypeIDs.Text)
  Debug.Assert (SingleSelectComboBoxElement.AriaRole = AriaRoles.ComboBox) Or (SingleSelectComboBoxElement.AriaRole = "")
 
  'Check the combobox expand state
  With SingleSelectComboBoxElement
    Debug.Print .ExpandedCollapsedState = ExpandCollapseState.ExpandCollapseState_Collapsed
    .Click
    Debug.Print .ExpandedCollapsedState = ExpandCollapseState.ExpandCollapseState_Expanded
    .Click
    Debug.Print .ExpandedCollapsedState = ExpandCollapseState.ExpandCollapseState_Collapsed
  End With

  'Now get each list item - Chromium needs the selection box to be expanded here
  SingleSelectComboBoxElement.Click
  Set SelectPlaywrightOption = SingleSelectComboBoxElement.FindDescendant("Playwright")
  Set SelectSeleniumOption = SingleSelectComboBoxElement.FindDescendant("Selenium")
  Set SelectCypressOption = SingleSelectComboBoxElement.FindDescendant("Cypress")
  Set SelectPhosphorusOption = SingleSelectComboBoxElement.FindDescendant("Phosphorus")
  
  'Check the names
  Debug.Assert SelectPlaywrightOption.Name = "Playwright"
  Debug.Assert SelectSeleniumOption.Name = "Selenium"
  Debug.Assert SelectCypressOption.Name = "Cypress"
  Debug.Assert SelectPhosphorusOption.Name = "Phosphorus"
  
  'Check the defaults!
  Debug.Assert SelectPlaywrightOption.IsSelected
  Debug.Assert SelectSeleniumOption.IsNotSelected
  Debug.Assert SelectCypressOption.IsNotSelected
  Debug.Assert SelectPhosphorusOption.IsNotSelected
  Debug.Assert SingleSelectComboBoxElement.Value = "Playwright"

  'Tick each box in turn & recheck
  Debug.Assert SingleSelectComboBoxElement.Click.IsExpanded
  Debug.Assert SelectSeleniumOption.Click.IsSelected
  Debug.Assert SelectPlaywrightOption.IsNotSelected
  Debug.Assert SelectCypressOption.IsNotSelected
  Debug.Assert SelectPhosphorusOption.IsNotSelected
  Debug.Assert SingleSelectComboBoxElement.Value = "Selenium"

  Debug.Assert SingleSelectComboBoxElement.Click.IsExpanded
  Debug.Assert SelectCypressOption.Click.IsSelected
  Debug.Assert SelectPlaywrightOption.IsNotSelected
  Debug.Assert SelectSeleniumOption.IsNotSelected
  Debug.Assert SelectPhosphorusOption.IsNotSelected
  Debug.Assert SingleSelectComboBoxElement.Value = "Cypress"

  SingleSelectComboBoxElement.Click.IsExpanded
  Debug.Assert SelectPhosphorusOption.Click.IsSelected
  Debug.Assert SelectPlaywrightOption.IsNotSelected
  Debug.Assert SelectSeleniumOption.IsNotSelected
  Debug.Assert SelectCypressOption.IsNotSelected
  Debug.Assert SingleSelectComboBoxElement.Value = "Phosphorus"

  SingleSelectComboBoxElement.Click.IsExpanded
  Debug.Assert SelectPlaywrightOption.Click.IsSelected
  Debug.Assert SelectSeleniumOption.IsNotSelected
  Debug.Assert SelectCypressOption.IsNotSelected
  Debug.Assert SelectPhosphorusOption.IsNotSelected
  Debug.Assert SingleSelectComboBoxElement.Value = "Playwright"

  'Debug.Assert VBA.Strings.InStr(1, SingleSelectComboBoxElement.Click.AriaProperties, "expanded=false")
  'Close the drop down
  SingleSelectComboBoxElement.Click
'  Debug.Assert SingleSelectComboBoxElement.ExpandedCollapsedState = ExpandCollapseState.ExpandCollapseState_Collapsed
  
  'MULTI-SELECT
  '=============

  'Check the header
  Dim Header2Text As pElement
  Header2 = "Multi-Select To Text Box"

  If Not Factory.CurrentWebBrowserType = WebBrowserType.Chromium Then
    Set Header2Text = DocumentRoot.FindElement("MultiSelectHeader", By.NameIs, Header2)
    Debug.Assert Header2Text.ControlType = UIAControlTypeIDs.Text
  End If
    
  'Check the Choose Framework text
  Dim ChooseFrameworkLabelMulti As pElement
  If Factory.CurrentWebBrowserType = WebBrowserType.Chromium Then
    Set ChooseFrameworkLabelMulti = DocumentRoot.FindElement("Choose Frameworks (Hold Ctrl/Shift to select multiple):", By.NameIs, "Choose Frameworks (Hold Ctrl/Shift to select multiple):", FindFirst:=True)
  Else
    Set ChooseFrameworkLabelMulti = Header2Text.FindSibling("Choose Frameworks (Hold Ctrl/Shift to select multiple):", UIAControlTypeIDs.Text)
  End If
  Debug.Assert ChooseFrameworkLabelMulti.ControlType = UIAControlTypeIDs.Text
  Debug.Assert ChooseFrameworkLabelMulti.Name = "Choose Frameworks (Hold Ctrl/Shift to select multiple):"

  'Check the combobox
  Set MultiSelectListBoxElement = ChooseFrameworkLabelMulti.FindLabelledElement
  Debug.Assert MultiSelectListBoxElement.ControlType = UIAControlTypeIDs.List
  Debug.Assert (MultiSelectListBoxElement.AriaRole = AriaRoles.ListBox) Or (MultiSelectListBoxElement.AriaRole = "")

  'Check the combobox expand state where possible - its not expandable!
  If MultiSelectListBoxElement.AriaProperties <> "" Then
    Debug.Assert VBA.Strings.InStr(1, MultiSelectListBoxElement.AriaProperties, "expanded=false") > 0
  End If
  
  'Now get each list item
  Set MultiSelectPlaywrightOption = MultiSelectListBoxElement.FindDescendant("Playwright")
  Set MultiSelectSeleniumOption = MultiSelectListBoxElement.FindDescendant("Selenium")
  Set MultiSelectCypressOption = MultiSelectListBoxElement.FindDescendant("Cypress")
  Set MultiSelectPhosphorusOption = MultiSelectListBoxElement.FindDescendant("Phosphorus")

  Set SelectedItemsLabelElement = MultiSelectListBoxElement.FindSibling("Selected Items:")
  Debug.Assert SelectedItemsLabelElement.Name = "Selected Items:"
  Debug.Assert SelectedItemsLabelElement.AriaRole = AriaRoles.Description Or SelectedItemsLabelElement.AriaRole = ""
  
  Set SelectedItemsEditElement = SelectedItemsLabelElement.FindLabelledElement
  Debug.Assert SelectedItemsEditElement.Name = "Selected Items:"
  Debug.Assert SelectedItemsEditElement.AriaRole = AriaRoles.TextBox Or SelectedItemsEditElement.AriaRole = ""

  'Check the names
  Debug.Assert MultiSelectPlaywrightOption.Name = "Playwright"
  Debug.Assert MultiSelectSeleniumOption.Name = "Selenium"
  Debug.Assert MultiSelectCypressOption.Name = "Cypress"
  Debug.Assert MultiSelectPhosphorusOption.Name = "Phosphorus"

  'Check the default!
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsNotSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = ""

  'Select & Deselect each item in turn

  'MultiSelectPlaywrightOption
  Debug.Assert MultiSelectPlaywrightOption.Click.IsSelected
  Debug.Assert MultiSelectSeleniumOption.IsNotSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Playwright"

  Debug.Assert MultiSelectPlaywrightOption.Click.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsNotSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = ""
 
  'MultiSelectSeleniumOption
  Debug.Assert MultiSelectSeleniumOption.Click.IsSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Selenium"
  
  Debug.Assert MultiSelectSeleniumOption.Click.IsNotSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = ""
 
  'MultiSelectCypressOption
  Debug.Assert MultiSelectCypressOption.Click.IsSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Cypress"
  
  Debug.Assert MultiSelectCypressOption.Click.IsNotSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = ""
 
  'MultiSelectPhosphorusOption
  Debug.Assert MultiSelectPhosphorusOption.Click.IsSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsNotSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Phosphorus"
  
  Debug.Assert MultiSelectPhosphorusOption.Click.IsNotSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsNotSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = ""
 
  'Select but don't Deselect each item in turn
  
  'MultiSelectPlaywrightOption
  Debug.Assert MultiSelectPlaywrightOption.Click.IsSelected
  Debug.Assert MultiSelectSeleniumOption.IsNotSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Playwright"

  'MultiSelectSeleniumOption
  Debug.Assert MultiSelectSeleniumOption.Click.IsSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Selenium"
   
  'MultiSelectCypressOption
  Debug.Assert MultiSelectCypressOption.Click.IsSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Cypress"
 
  'MultiSelectPhosphorusOption
  Debug.Assert MultiSelectPhosphorusOption.Click.IsSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsNotSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Phosphorus"
  
  'Select multiple individual items with Ctrl+Click
  
  'MultiSelectPlaywrightOption+MultiSelectSeleniumOption
  Debug.Assert MultiSelectPlaywrightOption.Click.IsSelected
  Debug.Assert MultiSelectSeleniumOption.CtrlClick.IsSelected
  Debug.Assert MultiSelectPlaywrightOption.IsSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Playwright, Selenium"
  
  '-MultiSelectPlaywrightOption
  Debug.Assert MultiSelectPlaywrightOption.CtrlClick.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Selenium"
  
  '+MultiSelectCypressOption
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsSelected
  Debug.Assert MultiSelectCypressOption.CtrlClick.IsSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Selenium, Cypress"
  
  '+MultiSelectPhosphorusOption
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectCypressOption.IsSelected
  Debug.Assert MultiSelectSeleniumOption.IsSelected
  Debug.Assert MultiSelectPhosphorusOption.CtrlClick.IsSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectCypressOption.IsSelected
  Debug.Assert MultiSelectSeleniumOption.IsSelected
  Debug.Assert SelectedItemsEditElement.Text = "Selenium, Cypress, Phosphorus"
  
  '-MultiSelectCypressOption
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectCypressOption.IsSelected
  Debug.Assert MultiSelectSeleniumOption.IsSelected
  Debug.Assert MultiSelectPhosphorusOption.IsSelected
  
  Debug.Assert MultiSelectCypressOption.CtrlClick.IsNotSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsSelected
  Debug.Assert MultiSelectPhosphorusOption.IsSelected
  Debug.Assert SelectedItemsEditElement.Text = "Selenium, Phosphorus"
  
  'MultiSelectPlaywrightOption click only
  Debug.Assert MultiSelectPlaywrightOption.Click.IsSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Playwright"
  Debug.Assert MultiSelectPlaywrightOption.Click.IsNotSelected

  'MultiSelectPlaywrightOption+MultiSelectCypressOption
  Debug.Assert MultiSelectPlaywrightOption.Click.IsSelected
  Debug.Assert MultiSelectCypressOption.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Playwright"

  Debug.Assert MultiSelectCypressOption.ShiftClick(MultiSelectPlaywrightOption).IsSelected
  Debug.Assert MultiSelectPlaywrightOption.IsSelected
  Debug.Assert MultiSelectSeleniumOption.IsSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Playwright, Selenium, Cypress"
  
  '+MultiSelectPhosphorusOption
  Debug.Assert MultiSelectPhosphorusOption.ShiftClick(MultiSelectPlaywrightOption).IsSelected
  Debug.Assert MultiSelectPlaywrightOption.IsSelected
  Debug.Assert MultiSelectSeleniumOption.IsSelected
  Debug.Assert MultiSelectCypressOption.IsSelected
  Debug.Assert SelectedItemsEditElement.Text = "Playwright, Selenium, Cypress, Phosphorus"

  '+MultiSelectSeleniumOption
  Debug.Assert MultiSelectSeleniumOption.ShiftClick(MultiSelectPhosphorusOption).IsSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsSelected
  Debug.Assert MultiSelectCypressOption.IsSelected
  Debug.Assert SelectedItemsEditElement.Text = "Selenium, Cypress, Phosphorus"
  
  '+MultiSelectCypressOption
  Debug.Assert MultiSelectCypressOption.ShiftClick(MultiSelectSeleniumOption).IsSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.IsSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Selenium, Cypress"
  
  '+MultiSelectPlaywrightOption
  Debug.Assert MultiSelectPlaywrightOption.ShiftClick(MultiSelectCypressOption).IsSelected
  Debug.Assert MultiSelectSeleniumOption.IsSelected
  Debug.Assert MultiSelectCypressOption.IsSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Playwright, Selenium, Cypress"
  
  'Clear all
  Debug.Assert MultiSelectPlaywrightOption.Click.IsNotSelected
  Debug.Assert MultiSelectSeleniumOption.Click.IsNotSelected
  Debug.Assert MultiSelectCypressOption.Click.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = ""
  
  'MultiSelectCypressOption+MultiSelectSeleniumOption
  Debug.Assert MultiSelectCypressOption.Click.IsSelected
  Debug.Assert MultiSelectSeleniumOption.ShiftClick(MultiSelectCypressOption).IsSelected
  Debug.Assert MultiSelectPlaywrightOption.IsNotSelected
  Debug.Assert MultiSelectPhosphorusOption.IsNotSelected
  Debug.Assert SelectedItemsEditElement.Text = "Selenium, Cypress"
  
  'Select & Click by name
  MultiSelectPlaywrightOption.Click.IsSelected: Debug.Assert SelectedItemsEditElement.Text = "Playwright"
  MultiSelectSeleniumOption.Click.IsSelected: Debug.Assert SelectedItemsEditElement.Text = "Selenium"
  MultiSelectCypressOption.Click.IsSelected: Debug.Assert SelectedItemsEditElement.Text = "Cypress"
  MultiSelectPhosphorusOption.Click.IsSelected: Debug.Assert SelectedItemsEditElement.Text = "Phosphorus"
  
  If RunningAllExamples Then
    Debug.Print "Selections Selected!"
  Else
    Window.HighlightElements = False
    MsgBox "Selections Selected!"
    Phosphorus.Log4PStatic.CloseLogger
  End If
  
  Succeeded = True
  GoTo ExitFunction
   
ErrorHandler:
  If RunningAllExamples Then
    Debug.Print Err.Description & " (Error Number #" & Err.Number & ") for " & Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser in 'Selections' Example!"
  Else
    MsgBox Err.Description & " (Error Number #" & Err.Number & ") for " & Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser in 'Selections' Example!", vbCritical
  End If
  GoTo ExitFunction
  
ExitFunction:
  Set WebBrowser = Nothing
  Set SearchBox = Nothing
  Set SearchButton = Nothing

  Set Header1Text = Nothing
  Set FrameworkLabelSingle = Nothing
  Set SingleSelectComboBoxElement = Nothing
  Set SelectPlaywrightOption = Nothing
  Set SelectSeleniumOption = Nothing
  Set SelectCypressOption = Nothing
  Set SelectPhosphorusOption = Nothing
  
  Set Header2Text = Nothing
  Set ChooseFrameworkLabelMulti = Nothing
  Set MultiSelectListBoxElement = Nothing
  Set MultiSelectPlaywrightOption = Nothing
  Set MultiSelectSeleniumOption = Nothing
  Set MultiSelectCypressOption = Nothing
  Set MultiSelectPhosphorusOption = Nothing
  Set SelectedItemsEditElement = Nothing
  
  If Not RunningAllExamples Then
    Window.HighlightElements = False
  End If
  
  Selections = Succeeded

End Function
