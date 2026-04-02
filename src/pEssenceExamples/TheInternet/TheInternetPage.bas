VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TheInternetPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder TheInternet
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Const WEB_APP_NAME = "The-Internet"
'Const TARGET_PAGE_URL = "https://the-internet.herokuapp.com/"
Const TARGET_PAGE_URL = "https://the-internet.herokuapp.com"
Const TARGET_PAGE_TITLE = "The Internet"
'Const TARGET_PAGE_LINK = "https://iana.org/domains/example"

Private WebBrowser As EdgeWebBrowser

Private Type PageAttributes
  RootWebArea As pLocator
  ForkMe As pLocator
  HomePageHeading1 As pLocator
  HomePageHeading2 As pLocator
  ListOfExamples As pLocator
End Type

Private This As PageAttributes

Private PrivateElements As New Scripting.Dictionary

Private Sub Class_Initialize()
  Set WebBrowser = Factory.GetNewEdgeWebBrowser
  Set This.ForkMe = Factory.GetNewLocator
  Set This.HomePageHeading1 = Factory.GetNewLocator
  Set This.HomePageHeading2 = Factory.GetNewLocator
  Set This.ListOfExamples = Factory.GetNewLocator
End Sub

Private Sub Class_Terminate()
  Set WebBrowser = Nothing
  Set This.RootWebArea = Nothing
  Set This.ForkMe = Nothing
  Set This.HomePageHeading1 = Nothing
  Set This.HomePageHeading2 = Nothing
  Set This.ListOfExamples = Nothing
End Sub

Public Sub Initialize()
  
  With WebBrowser
    .StartNormal WEB_APP_NAME, TARGET_PAGE_URL, TARGET_PAGE_TITLE
    Set This.RootWebArea = .GetRootWebArea
  End With
  
  With This.ForkMe
    .Initialise "ForkMe", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, "link"
    .Condition "NameIs", Name, IsTheString, "Fork me on GitHub"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 1
  End With
  
  With This.HomePageHeading1
    .Initialise "Heading1", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, "heading"
    .Condition "NameIs", Name, IsTheString, "Welcome to the-internet"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 2
  End With

  With This.HomePageHeading2
    .Initialise "Heading1", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, "heading"
    .Condition "NameIs", Name, IsTheString, "Available Examples"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 3
  End With

  With This.ListOfExamples
    .Initialise "ListOfExamples", This.RootWebArea, Children, pConditions, "AND(AriaRole, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, "list"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 4
  End With

End Sub

Public Sub RunHomePageChecks()
  Debug.Assert WebBrowser.GetCurrentURL = TARGET_PAGE_URL
  Debug.Assert UIAProps.GetProperty(This.RootWebArea.FoundUIAElement, Name) = TARGET_PAGE_TITLE
  Debug.Assert This.HomePageHeading1.ElementExists(10)
  Debug.Assert UIAProps.GetProperty(This.HomePageHeading1.FoundUIAElement, Level) = 1
  Debug.Assert This.HomePageHeading2.ElementExists(10)
  Debug.Assert UIAProps.GetProperty(This.HomePageHeading2.FoundUIAElement, Level) = 2
  Debug.Assert This.ListOfExamples.ElementExists(10) = True
  Debug.Assert UIAProps.GetProperty(This.ListOfExamples.FoundUIAElement, SizeOfSet) = 44
End Sub

Private Sub SelectListItem(ItemName As String, Optional SubPageHeadingText)
    
  Dim ListItem As pLocator
  Set ListItem = Factory.GetNewLocator
  With ListItem
    .Initialise "ListItem", This.ListOfExamples, Children, pConditions, "AND(ControlType, NameIs)"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.ListItem
    .Condition "NameIs", Name, IsTheString, ItemName
    .Find 10
  End With
  
  Dim ListItemHyperlink As pLocator
  Set ListItemHyperlink = Factory.GetNewLocator
  With ListItemHyperlink
    .Initialise "ListItem", ListItem, Children, pConditions, "AND(ControlType, NameIs)"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Hyperlink
    .Condition "NameIs", Name, IsTheString, ItemName
    .Find 10
  End With
  
  Actions.Click ItemName, ListItemHyperlink.FoundUIAElement
  
  Set ListItem = Nothing
  Set ListItemHyperlink = Nothing
  
  Debug.Assert This.RootWebArea.ElementExists(10)
  Debug.Assert This.ForkMe.ElementExists(2)
  Debug.Assert This.HomePageHeading1.ElementDoesntExist(0)
  Debug.Assert This.HomePageHeading2.ElementDoesntExist(0)
  Debug.Assert This.ListOfExamples.ElementDoesntExist(0)
      
  If IsMissing(SubPageHeadingText) Then
    SubPageHeadingText = ItemName
  End If
  
  Dim SubPageHeading As pLocator
  Set SubPageHeading = Factory.GetNewLocator
  With SubPageHeading
    .Initialise "SubPageHeading", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, "heading"
    .Condition "NameIs", Name, IsTheString, SubPageHeadingText
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 2
    Debug.Assert .ElementExists(2)
  End With

End Sub

Public Sub Checkboxes()
  
  SelectListItem "Checkboxes"
  Debug.Assert WebBrowser.GetCurrentURL = TARGET_PAGE_URL & "/checkboxes"

  Dim FirstCheckbox As pLocator
  Set FirstCheckbox = Factory.GetNewLocator
  With FirstCheckbox
    .Initialise "FirstCheckbox", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.CheckBox
    .Condition "AriaRole", AriaRole, IsTheString, "checkbox"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 3
    .Find 10
  End With

  Dim FirstCheckboxDescription As pLocator
  Set FirstCheckboxDescription = Factory.GetNewLocator
  With FirstCheckboxDescription
    .Initialise "FirstCheckboxDescription", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ", NameIs)"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Text
    .Condition "AriaRole", AriaRole, IsTheString, "description"
    .Condition "NameIs", Name, IsTheString, " checkbox 1"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 4
    .Find 10
  End With

  Dim SecondCheckbox As pLocator
  Set SecondCheckbox = Factory.GetNewLocator
  With SecondCheckbox
    .Initialise "SecondCheckbox", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.CheckBox
    .Condition "AriaRole", AriaRole, IsTheString, "checkbox"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 6
    .Find 10
  End With

  Dim SecondCheckboxDescription As pLocator
  Set SecondCheckboxDescription = Factory.GetNewLocator
  With SecondCheckboxDescription
    .Initialise "FirstCheckboxDescription", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ", NameIs)"
    .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Text
    .Condition "AriaRole", AriaRole, IsTheString, "description"
    .Condition "NameIs", Name, IsTheString, " checkbox 2"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 7
    .Find 10
  End With
  
  'The second checkbox is already checked!
  Debug.Assert Actions.GetToggleState(SecondCheckbox.ElementName, SecondCheckbox.FoundUIAElement)
  
  Actions.Click FirstCheckbox.ElementName, FirstCheckbox.FoundUIAElement
  Debug.Assert Actions.GetToggleState(FirstCheckbox.ElementName, FirstCheckbox.FoundUIAElement)
  
  Actions.Click SecondCheckbox.ElementName, SecondCheckbox.FoundUIAElement
  Debug.Assert Not Actions.GetToggleState(SecondCheckbox.ElementName, SecondCheckbox.FoundUIAElement)
  
  Actions.Click FirstCheckbox.ElementName, FirstCheckbox.FoundUIAElement
  Debug.Assert Not Actions.GetToggleState(FirstCheckbox.ElementName, FirstCheckbox.FoundUIAElement)
  
  WebBrowser.NavigateBack
  Debug.Assert WebBrowser.GetCurrentURL = TARGET_PAGE_URL

End Sub

Public Sub DragAndDrop()
  
  SelectListItem "Drag and Drop"
  Debug.Assert WebBrowser.GetCurrentURL = TARGET_PAGE_URL & "/drag_and_drop"
  
  Dim i As Integer
  For i = 1 To 4
  
    Dim FirstItem As pLocator
    Set FirstItem = Factory.GetNewLocator
    With FirstItem
      .Initialise "FirstItem", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
      .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Group
      .Condition "AriaRole", AriaRole, IsTheString, "banner"
      .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 3
      .Find 10
    End With

    Dim SecondItem As pLocator
    Set SecondItem = Factory.GetNewLocator
    With SecondItem
      .Initialise "SecondItem", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
      .Condition "ControlType", ControlType, EqualsNumber, UIAControlTypeIDs.Group
      .Condition "AriaRole", AriaRole, IsTheString, "banner"
      .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 4
      .Find 10
    End With
    Select Case i
      Case 1 'A->B
        Actions.DragAndDrop "FirstItem", FirstItem.FoundUIAElement, "SecondItem", SecondItem.FoundUIAElement
      Case 2 'B<-A
        Actions.DragAndDrop "SecondItem", SecondItem.FoundUIAElement, "FirstItem", FirstItem.FoundUIAElement
      Case 3 'B<-A
        Actions.DragAndDrop "SecondItem", SecondItem.FoundUIAElement, "FirstItem", FirstItem.FoundUIAElement
      Case 4 'B->A
        Actions.DragAndDrop "FirstItem", FirstItem.FoundUIAElement, "SecondItem", SecondItem.FoundUIAElement
    End Select
    
    Set FirstItem = Nothing
    Set SecondItem = Nothing
  
  Next i
  
  WebBrowser.NavigateBack
  Debug.Assert WebBrowser.GetCurrentURL = TARGET_PAGE_URL
   
End Sub

Public Sub FormAuthentication()
  SelectListItem "Form Authentication", "Login Page"
  FormAuthentication_Login
  FormAuthentication_Secure
  WebBrowser.NavigateBack 'Takes us back to the secure page
  WebBrowser.NavigateBack 'Takes us back to the login page
  WebBrowser.NavigateBack 'Finally takes us back to the homepage!
  Debug.Assert WebBrowser.GetCurrentURL = TARGET_PAGE_URL
End Sub
  
Private Sub FormAuthentication_Login()

  Debug.Assert WebBrowser.GetCurrentURL = TARGET_PAGE_URL & "/login"

  Dim SubHeader As pLocator
  Set SubHeader = Factory.GetNewLocator
  With SubHeader
    .Initialise "SubHeader", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, ClassName, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "ControlType", ControlType, EqualsNumber, Text
    .Condition "AriaRole", AriaRole, IsTheString, "heading"
    .Condition "ClassName", ClassName, IsTheString, "subheader"
    .Condition "NameIs", Name, IsTheString, "This is where you can log into the secure area. Enter tomsmith for the username and SuperSecretPassword! for the password. If the information is wrong you should see error messages."
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 3
    Debug.Assert .ElementExists(5)
  End With

  Dim UsernameLabel As pLocator
  Set UsernameLabel = Factory.GetNewLocator
  With UsernameLabel
    .Initialise "UsernameLabel", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "ControlType", ControlType, EqualsNumber, Text
    .Condition "AriaRole", AriaRole, IsTheString, "description"
    .Condition "NameIs", Name, IsTheString, "Username"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 4
    Debug.Assert .ElementExists(0)
  End With

  Dim UsernameTextBox As pLocator
  Set UsernameTextBox = Factory.GetNewLocator
  With UsernameTextBox
    .Initialise "UsernameTextBox", This.RootWebArea, Children, pConditions, "AND(ControlType, AutomationId, AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "ControlType", ControlType, EqualsNumber, Edit
    .Condition "AutomationId", UIAProperties.AutomationId, IsTheString, "username"
    .Condition "AriaRole", AriaRole, IsTheString, "textbox"
    .Condition "NameIs", Name, IsTheString, "Username"
    '.Condition "LabeledBy", LabeledBy, IsTheString, "Username" ' This returns an element!
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 5
    Debug.Assert .ElementExists(0)
  End With

  Dim PasswordLabel As pLocator
  Set PasswordLabel = Factory.GetNewLocator
  With PasswordLabel
    .Initialise "PasswordLabel", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "ControlType", ControlType, EqualsNumber, Text
    .Condition "AriaRole", AriaRole, IsTheString, "description"
    .Condition "NameIs", Name, IsTheString, "Password"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 6
    Debug.Assert .ElementExists(0)
  End With

  Dim PasswordTextBox As pLocator
  Set PasswordTextBox = Factory.GetNewLocator
  With PasswordTextBox
    .Initialise "FirstItem", This.RootWebArea, Children, pConditions, "AND(ControlType, AutomationId, AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "ControlType", ControlType, EqualsNumber, Edit
    .Condition "AutomationId", UIAProperties.AutomationId, IsTheString, "password"
    .Condition "AriaRole", AriaRole, IsTheString, "textbox"
    .Condition "NameIs", Name, IsTheString, "Password"
    '.Condition "LabeledBy", LabeledBy, IsTheString, "Password" ' This returns an element!
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 7
    Debug.Assert .ElementExists(0)
  End With

  Dim LoginButton As pLocator
  Set LoginButton = Factory.GetNewLocator
  With LoginButton
    .Initialise "LoginButton", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, NameEndsWith, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "ControlType", ControlType, EqualsNumber, Button
    .Condition "AriaRole", AriaRole, IsTheString, "button"
    'Name='? Login' - starts with a unicode character?
    .Condition "NameEndsWith", Name, EndsWithTheString, " Login"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 8
    Debug.Assert .ElementExists(0)
  End With

  Actions.SetValue UsernameTextBox.ElementName, UsernameTextBox.FoundUIAElement, "tomsmith"
  Actions.SetValue PasswordTextBox.ElementName, PasswordTextBox.FoundUIAElement, "SuperSecretPassword!"
  Actions.Click LoginButton.ElementName, LoginButton.FoundUIAElement
  
  Set SubHeader = Nothing
  Set UsernameLabel = Nothing
  Set UsernameTextBox = Nothing
  Set PasswordLabel = Nothing
  Set PasswordTextBox = Nothing
  Set LoginButton = Nothing

End Sub

Private Sub FormAuthentication_Secure()

'We need to a FindElement here!
  Snooze 1000
Debug.Assert This.RootWebArea.ElementExists
  Debug.Assert WebBrowser.GetCurrentURL = TARGET_PAGE_URL & "/secure"

  Dim SecureMesssage As pLocator
  Set SecureMesssage = Factory.GetNewLocator
  With SecureMesssage
    .Initialise "SecureMesssage", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "ControlType", ControlType, EqualsNumber, Text
    .Condition "AriaRole", AriaRole, IsTheString, "description"
    .Condition "NameIs", Name, IsTheString, " You logged into a secure area!"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 2
    Debug.Assert .ElementExists(5)
  End With

  Dim SecureMesssageHyperlink As pLocator
  Set SecureMesssageHyperlink = Factory.GetNewLocator
  With SecureMesssageHyperlink
    .Initialise "SecureMesssageHyperlink", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "ControlType", ControlType, EqualsNumber, Hyperlink
    .Condition "AriaRole", AriaRole, IsTheString, "link"
    .Condition "NameIs", Name, IsTheString, "×"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 3
    Debug.Assert .ElementExists(0)
  End With

  'Click on it then verify the element doesn't exist! Wait?
  Actions.Click SecureMesssageHyperlink.ElementName, SecureMesssageHyperlink.FoundUIAElement
  Snooze 500
  Debug.Assert SecureMesssage.ElementDoesntExist(0)
  Debug.Assert SecureMesssageHyperlink.ElementDoesntExist(0)
  
  Dim Heading As pLocator
  Set Heading = Factory.GetNewLocator
  With Heading
    .Initialise "Heading", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "ControlType", ControlType, EqualsNumber, Text
    .Condition "AriaRole", AriaRole, IsTheString, "heading"
    .Condition "NameIs", Name, IsTheString, "Secure Area"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 2
    Debug.Assert .ElementExists(5)
  End With

  Dim SubHeader As pLocator
  Set SubHeader = Factory.GetNewLocator
  With SubHeader
    .Initialise "SubHeader", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, ClassName, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "ControlType", ControlType, EqualsNumber, Text
    .Condition "AriaRole", AriaRole, IsTheString, "heading"
    .Condition "ClassName", ClassName, IsTheString, "subheader"
    .Condition "NameIs", Name, IsTheString, "Welcome to the Secure Area. When you are done click logout below."
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 3
    Debug.Assert .ElementExists(0)
  End With

  Dim Logout As pLocator
  Set Logout = Factory.GetNewLocator
  With Logout
    .Initialise "SubHeader", This.RootWebArea, Children, pConditions, "AND(ControlType, AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "ControlType", ControlType, EqualsNumber, Hyperlink
    .Condition "AriaRole", AriaRole, IsTheString, "link"
    .Condition "NameIs", Name, IsTheString, "Logout"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 4
    Debug.Assert .ElementExists(0)
  End With

'Click logout takes us back!
  Actions.Click Logout.ElementName, Logout.FoundUIAElement
  Snooze 500
  Debug.Assert Heading.ElementDoesntExist(2)

  Set SecureMesssage = Nothing
  Set SecureMesssageHyperlink = Nothing
  Set SecureMesssage = Nothing
  Set SubHeader = Nothing
  Set Logout = Nothing
  
End Sub

