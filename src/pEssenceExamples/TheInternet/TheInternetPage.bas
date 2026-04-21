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
Const TARGET_PAGE_URL = "https://the-internet.herokuapp.com"
Const TARGET_PAGE_TITLE = "The Internet"

Private Type PageAttributes
  WebBrowser As Object
  RootWebArea As pLocator
  ForkMe As pLocator
  HomePageHeading1 As pLocator
  HomePageHeading2 As pLocator
  ListOfExamples As pLocator
End Type

Private This As PageAttributes

Private Sub Class_Initialize()
  Set This.WebBrowser = Factory.GetNewWebBrowser
  Set This.ForkMe = Factory.GetNewLocator
  Set This.HomePageHeading1 = Factory.GetNewLocator
  Set This.HomePageHeading2 = Factory.GetNewLocator
  Set This.ListOfExamples = Factory.GetNewLocator
End Sub

Private Sub Class_Terminate()
  Set This.WebBrowser = Nothing
  Set This.RootWebArea = Nothing
  Set This.ForkMe = Nothing
  Set This.HomePageHeading1 = Nothing
  Set This.HomePageHeading2 = Nothing
  Set This.ListOfExamples = Nothing
End Sub

Public Sub Initialize()
  
  With This.WebBrowser
    .Start WEB_APP_NAME, TARGET_PAGE_URL, TARGET_PAGE_TITLE
    Set This.RootWebArea = .GetRootWebArea
  End With
  
  With This.ForkMe
    .Initialise "ForkMe", This.RootWebArea, Descendants, pConditions, "AND(AriaRoleLink, NameIs)"
    .AriaRoleLink: .NameIs "Fork me on GitHub"
  End With
  
  With This.HomePageHeading1
    .Initialise "Heading1", This.RootWebArea, Descendants, pConditions, "AND(AriaRoleHeading, NameIs)"
    .AriaRoleHeading: .NameIs "Welcome to the-internet"
  End With

  With This.HomePageHeading2
    .Initialise "Heading1", This.RootWebArea, Descendants, pConditions, "AND(AriaRoleHeading, NameIs)"
    .AriaRoleHeading: .NameIs "Available Examples"
  End With

  With This.ListOfExamples
    .Initialise "ListOfExamples", This.RootWebArea, Descendants, By.AriaRole, AriaRoles.List
  End With

End Sub

Public Sub RunHomePageChecks()
  Debug.Assert (This.WebBrowser.GetCurrentURL = TARGET_PAGE_URL) Or (This.WebBrowser.GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", ""))
  Debug.Assert This.RootWebArea.Element.GetProperty(Name) = TARGET_PAGE_TITLE
  Debug.Assert This.HomePageHeading1.ElementExists(10)
  Debug.Assert This.HomePageHeading1.Element.GetProperty(Level) = 1
  Debug.Assert This.HomePageHeading2.ElementExists(10)
  Debug.Assert This.HomePageHeading2.Element.GetProperty(Level) = 2
  Debug.Assert This.ListOfExamples.ElementExists(10) = True
  Debug.Assert This.ListOfExamples.Element.GetProperty(SizeOfSet) = 44 Or This.ListOfExamples.Element.GetProperty(SizeOfSet) = 0 '44 Fails in Firefox!?
End Sub

Private Sub SelectListItem(ItemName As String, Optional SubPageHeadingText As String)
    
  Dim ListItem As pLocator
  Set ListItem = Factory.GetNewLocator
  With ListItem
    .Initialise "ListItem", This.ListOfExamples, Children, pConditions, "AND(AriaRoleListItem, NameIs)"
    .AriaRoleListItem: .NameIs ItemName
    .Find 10
  End With
  
  Dim ListItemHyperlink As pLocator
  Set ListItemHyperlink = Factory.GetNewLocator
  With ListItemHyperlink
    .Initialise "ListItemHyperlink", ListItem, Children, pConditions, "AND(AriaRoleLink, NameIs)"
    .AriaRoleLink: .NameIs ItemName
    .Find 10
  End With
  
  Actions.Click ListItemHyperlink.Element
  
  Set ListItem = Nothing
  Set ListItemHyperlink = Nothing
  
  Debug.Assert This.RootWebArea.ElementExists(10)
  Debug.Assert This.ForkMe.ElementExists(2)
  Debug.Assert This.HomePageHeading1.ElementDoesntExist(0)
  Debug.Assert This.HomePageHeading2.ElementDoesntExist(0)
  Debug.Assert This.ListOfExamples.ElementDoesntExist(0)
      
  If SubPageHeadingText = "" Then
    SubPageHeadingText = ItemName
  End If
  
  Dim SubPageHeading As pLocator
  Set SubPageHeading = Factory.GetNewLocator
  With SubPageHeading
    .Initialise "SubPageHeading", This.RootWebArea, Descendants, pConditions, "AND(AriaRoleHeading, NameIs)"
    .AriaRoleHeading: .NameIs SubPageHeadingText:
    Debug.Assert .ElementExists(2)
  End With

End Sub

Public Sub Checkboxes()
  
  SelectListItem "Checkboxes"
  Debug.Assert (This.WebBrowser.GetCurrentURL = TARGET_PAGE_URL & "/checkboxes") Or (This.WebBrowser.GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", "") & "/checkboxes")

  Dim FirstCheckbox As pLocator
  Set FirstCheckbox = Factory.GetNewLocator
  With FirstCheckbox
    .Initialise "FirstCheckbox", This.RootWebArea, Descendants, By.AriaRole, AriaRoles.CheckBox: .PositionInMatchingSet 1
    .Find 10
  End With

  Dim FirstCheckboxDescription As pLocator
  Set FirstCheckboxDescription = Factory.GetNewLocator
  With FirstCheckboxDescription
    .Initialise "FirstCheckboxDescription", This.RootWebArea, Descendants, pConditions, "AND(OR(AriaRoleDescription, ControlType), NameIs)"
    .AriaRoleDescription: .NameIs " checkbox 1": .ControlType Text
    .Find 10
  End With

  Dim SecondCheckbox As pLocator
  Set SecondCheckbox = Factory.GetNewLocator
  With SecondCheckbox
    .Initialise "SecondCheckbox", This.RootWebArea, Descendants, By.AriaRole, AriaRoles.CheckBox
    .PositionInMatchingSet 2
    .Find 10
  End With

  Dim SecondCheckboxDescription As pLocator
  Set SecondCheckboxDescription = Factory.GetNewLocator
  With SecondCheckboxDescription
    .Initialise "SecondCheckboxDescription", This.RootWebArea, Descendants, pConditions, "AND(OR(AriaRoleDescription, ControlType), OR(NameIsCB21,NameIsCB22))"
    .AriaRoleDescription: .Condition "NameIsCB21", Name, IsTheString, " checkbox 2": .Condition "NameIsCB22", Name, IsTheString, " checkbox 2 "
    .ControlType Text
    .Find 10
  End With
  
  'The second checkbox is already checked!
  Debug.Assert SecondCheckbox.Element.GetToggleState()
  
  Actions.Click FirstCheckbox.Element
  Debug.Assert FirstCheckbox.Element.GetToggleState()
  
  Actions.Click SecondCheckbox.Element
  Debug.Assert Not SecondCheckbox.Element.GetToggleState()
  
  Actions.Click FirstCheckbox.Element
  Debug.Assert Not FirstCheckbox.Element.GetToggleState()

  This.WebBrowser.NavigateBack
  Debug.Assert (This.WebBrowser.GetCurrentURL = TARGET_PAGE_URL) Or (This.WebBrowser.GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", ""))

End Sub

Public Sub DragAndDrop()
  
  SelectListItem "Drag and Drop"
  Debug.Assert (This.WebBrowser.GetCurrentURL = TARGET_PAGE_URL & "/drag_and_drop") Or (This.WebBrowser.GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", "") & "/drag_and_drop")
  
  Dim i As Integer
  For i = 1 To 4
  
    Dim FirstItem As pLocator
    Set FirstItem = Factory.GetNewLocator
    With FirstItem
      .Initialise "FirstItem", This.RootWebArea, Descendants, By.AriaRole, AriaRoles.Banner
      '.PositionInTreescope 3
      .PositionInMatchingSet 1
      .Find 10
    End With

    Dim SecondItem As pLocator
    Set SecondItem = Factory.GetNewLocator
    With SecondItem
      .Initialise "SecondItem", This.RootWebArea, Descendants, By.AriaRole, AriaRoles.Banner:
      '.PositionInTreescope 4
      .PositionInMatchingSet 2
      .Find 10
    End With
        
    Select Case i
      Case 1 'A->B
        Actions.DragAndDrop FirstItem.Element, SecondItem.Element
      Case 2 'B<-A
        Actions.DragAndDrop SecondItem.Element, FirstItem.Element
      Case 3 'B<-A
        Actions.DragAndDrop SecondItem.Element, FirstItem.Element
      Case 4 'B->A
        Actions.DragAndDrop FirstItem.Element, SecondItem.Element
    End Select
    
    Set FirstItem = Nothing
    Set SecondItem = Nothing
  
  Next i
  
  This.WebBrowser.NavigateBack
  Debug.Assert (This.WebBrowser.GetCurrentURL = TARGET_PAGE_URL) Or (This.WebBrowser.GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", ""))

End Sub

Public Sub FormAuthentication()
  SelectListItem "Form Authentication", "Login Page"
  FormAuthentication_Login
  FormAuthentication_Secure
  This.WebBrowser.NavigateBack 'Takes us back to the secure page
  This.WebBrowser.NavigateBack 'Takes us back to the login page
  This.WebBrowser.NavigateBack 'Finally takes us back to the homepage!
  Debug.Assert (This.WebBrowser.GetCurrentURL = TARGET_PAGE_URL) Or (This.WebBrowser.GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", ""))
End Sub
  
Private Sub FormAuthentication_Login()

  Debug.Assert (This.WebBrowser.GetCurrentURL = TARGET_PAGE_URL & "/login") Or (This.WebBrowser.GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", "") & "/login")

  Dim SubHeader As pLocator
  Set SubHeader = Factory.GetNewLocator
  With SubHeader
    .Initialise "SubHeader", This.RootWebArea, Descendants, pConditions, "AND(AriaRoleHeading, ClassName, NameIs)"
    .AriaRoleHeading: .ClassName "subheader"
    .NameIs "This is where you can log into the secure area. Enter tomsmith for the username and SuperSecretPassword! for the password. If the information is wrong you should see error messages."
    Debug.Assert .ElementExists(5)
  End With

  Dim UsernameLabel As pLocator
  Set UsernameLabel = Factory.GetNewLocator
  With UsernameLabel
    .Initialise "UsernameLabel", This.RootWebArea, Descendants, pConditions, "AND(OR(AriaRoleDescription, ControlType), NameIs)"
    .AriaRoleDescription: .NameIs "Username"
    .ControlType UIAControlTypeIDs.Group
    Debug.Assert .ElementExists(0)
  End With

  Dim UsernameTextBox As pLocator
  Set UsernameTextBox = Factory.GetNewLocator
  With UsernameTextBox
    .Initialise "UsernameTextBox", This.RootWebArea, Descendants, pConditions, "AND(AriaRoleTextBox, NameIs)"
    .AriaRoleTextBox: .NameIs "Username"
    '.Condition "LabeledBy", LabeledBy, IsTheString, "Username" ' This returns an element!
    Debug.Assert .ElementExists(0)
  End With

  Dim PasswordLabel As pLocator
  Set PasswordLabel = Factory.GetNewLocator
  With PasswordLabel
    .Initialise "PasswordLabel", This.RootWebArea, Descendants, pConditions, "AND(OR(AriaRoleDescription, ControlType), NameIs)"
    .AriaRoleDescription: .NameIs "Password"
    .ControlType UIAControlTypeIDs.Group
    Debug.Assert .ElementExists(0)
  End With

  Dim PasswordTextBox As pLocator
  Set PasswordTextBox = Factory.GetNewLocator
  With PasswordTextBox
    .Initialise "PasswordTextBox", This.RootWebArea, Descendants, pConditions, "AND(OR(AriaRoleTextBox, ControlType), NameIs)"
    .AriaRoleTextBox: .NameIs "Password"
    .ControlType UIAControlTypeIDs.Edit
    'TODO: .Condition "LabeledBy", LabeledBy, IsTheString, "Password" ' This returns an element!
    Debug.Assert .ElementExists(0)
  End With

  Dim LoginButton As pLocator
  Set LoginButton = Factory.GetNewLocator
  With LoginButton
    .Initialise "LoginButton", This.RootWebArea, Descendants, pConditions, "AND(AriaRoleButton, NameEndsWith)"
    .AriaRoleButton: .Condition "NameEndsWith", Name, EndsWithTheString, " Login"
    'Name='? Login' - starts with a unicode character?
    Debug.Assert .ElementExists(0)
  End With

  UsernameTextBox.Element.SetValue "tomsmith"
  PasswordTextBox.Element.SetValue "SuperSecretPassword!"
  Actions.Click LoginButton.Element

  If Factory.CurrentWebBrowserType = Chrome Then
    This.WebBrowser.AcknowledgeChangeYourPasswordAlert
  End If

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
  Debug.Assert (This.WebBrowser.GetCurrentURL = TARGET_PAGE_URL & "/secure") Or (This.WebBrowser.GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", "") & "/secure")

  Dim SecureMesssage As pLocator
  Set SecureMesssage = Factory.GetNewLocator
  With SecureMesssage
    .Initialise "SecureMesssage", This.RootWebArea, Descendants, pConditions, "AND(OR(AriaRoleDescription, ControlType), OR(NameIs1, NameIs2))"
    .AriaRoleDescription
    .Condition "NameIs1", Name, IsTheString, " You logged into a secure area!"
    .Condition "NameIs2", Name, IsTheString, " You logged into a secure area! "
    .ControlType UIAControlTypeIDs.Text
    Debug.Assert .ElementExists(5)
  End With

  Dim SecureMesssageHyperlink As pLocator
  Set SecureMesssageHyperlink = Factory.GetNewLocator
  With SecureMesssageHyperlink
    .Initialise "SecureMesssageHyperlink", This.RootWebArea, Descendants, pConditions, "AND(AriaRoleLink, NameIs)"
    .AriaRoleLink: .NameIs "×"
    Debug.Assert .ElementExists(0)
  End With

  'Click on it then verify the element doesn't exist! Wait?
  Actions.Click SecureMesssageHyperlink.Element
  Snooze 500
  Debug.Assert SecureMesssage.ElementDoesntExist(0)
  Debug.Assert SecureMesssageHyperlink.ElementDoesntExist(0)
  
  Dim Heading As pLocator
  Set Heading = Factory.GetNewLocator
  With Heading
    .Initialise "Heading", This.RootWebArea, Descendants, pConditions, "AND(AriaRoleHeading, NameIs)"
    .AriaRoleHeading: .NameIs "Secure Area"
    Debug.Assert .ElementExists(5)
  End With

  Dim SubHeader As pLocator
  Set SubHeader = Factory.GetNewLocator
  With SubHeader
    .Initialise "SubHeader", This.RootWebArea, Descendants, pConditions, "AND(AriaRoleHeading, ClassName, NameIs)"
    .AriaRoleHeading: .ClassName "subheader": .NameIs "Welcome to the Secure Area. When you are done click logout below."
    Debug.Assert .ElementExists(0)
  End With

  Dim Logout As pLocator
  Set Logout = Factory.GetNewLocator
  With Logout
    .Initialise "SubHeader", This.RootWebArea, Descendants, pConditions, "AND(AriaRoleLink, NameIs)"
    .AriaRoleLink: .NameIs "Logout" '
    Debug.Assert .ElementExists(0)
  End With

  'Click logout takes us back!
  Actions.Click Logout.Element
  Snooze 500
  Debug.Assert Heading.ElementDoesntExist(2)

  Set SecureMesssage = Nothing
  Set SecureMesssageHyperlink = Nothing
  Set SecureMesssage = Nothing
  Set SubHeader = Nothing
  Set Logout = Nothing
  
End Sub

