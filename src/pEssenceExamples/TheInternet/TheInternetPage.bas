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
    .pWB_Start WEB_APP_NAME, TARGET_PAGE_URL, TARGET_PAGE_TITLE
    Set This.RootWebArea = .pWB_GetRootWebArea
  End With
  
  With This.ForkMe
    .Initialise "ForkMe", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Link
    .NameIs "Fork me on GitHub"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 1
  End With
  
  With This.HomePageHeading1
    .Initialise "Heading1", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Heading
    .NameIs "Welcome to the-internet"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 2
  End With

  With This.HomePageHeading2
    .Initialise "Heading1", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Heading
    .NameIs "Available Examples"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 3
  End With

  With This.ListOfExamples
    .Initialise "ListOfExamples", This.RootWebArea, Children, pConditions, "AND(AriaRole, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.List
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 4
  End With

End Sub

Public Sub RunHomePageChecks()
  Debug.Assert (This.WebBrowser.pWB_GetCurrentURL = TARGET_PAGE_URL) Or (This.WebBrowser.pWB_GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", ""))
  Debug.Assert This.RootWebArea.Element.GetProperty(Name) = TARGET_PAGE_TITLE
  Debug.Assert This.HomePageHeading1.ElementExists(10)
  Debug.Assert This.HomePageHeading1.Element.GetProperty(Level) = 1
  Debug.Assert This.HomePageHeading2.ElementExists(10)
  Debug.Assert This.HomePageHeading2.Element.GetProperty(Level) = 2
  Debug.Assert This.ListOfExamples.ElementExists(10) = True
  Debug.Assert This.ListOfExamples.Element.GetProperty(SizeOfSet) = 44
End Sub

Private Sub SelectListItem(ItemName As String, Optional SubPageHeadingText As String)
    
  Dim ListItem As pLocator
  Set ListItem = Factory.GetNewLocator
  With ListItem
    .Initialise "ListItem", This.ListOfExamples, Children, pConditions, "AND(AriaRole, NameIs)"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.ListItem
    .NameIs ItemName
    .Find 10
  End With
  
  Dim ListItemHyperlink As pLocator
  Set ListItemHyperlink = Factory.GetNewLocator
  With ListItemHyperlink
    .Initialise "ListItemHyperlink", ListItem, Children, pConditions, "AND(AriaRole, NameIs)"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Link
    .NameIs ItemName
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
    .Initialise "SubPageHeading", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Heading
    .NameIs SubPageHeadingText
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 2
    Debug.Assert .ElementExists(2)
  End With

End Sub

Public Sub Checkboxes()
  
  SelectListItem "Checkboxes"
  Debug.Assert (This.WebBrowser.pWB_GetCurrentURL = TARGET_PAGE_URL & "/checkboxes") Or (This.WebBrowser.pWB_GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", "") & "/checkboxes")

  Dim FirstCheckbox As pLocator
  Set FirstCheckbox = Factory.GetNewLocator
  With FirstCheckbox
    .Initialise "FirstCheckbox", This.RootWebArea, Children, pConditions, "AND(AriaRole, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.CheckBox
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 3
    .Find 10
  End With

  Dim FirstCheckboxDescription As pLocator
  Set FirstCheckboxDescription = Factory.GetNewLocator
  With FirstCheckboxDescription
    .Initialise "FirstCheckboxDescription", This.RootWebArea, Children, pConditions, "AND(AriaRole, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ", NameIs)"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Description
    .NameIs " checkbox 1"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 4
    .Find 10
  End With

  Dim SecondCheckbox As pLocator
  Set SecondCheckbox = Factory.GetNewLocator
  With SecondCheckbox
    .Initialise "SecondCheckbox", This.RootWebArea, Children, pConditions, "AND(AriaRole, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.CheckBox
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 6
    .Find 10
  End With

  Dim SecondCheckboxDescription As pLocator
  Set SecondCheckboxDescription = Factory.GetNewLocator
  With SecondCheckboxDescription
    .Initialise "FirstCheckboxDescription", This.RootWebArea, Children, pConditions, "AND(AriaRole, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ", NameIs)"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Description
    .NameIs " checkbox 2"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 7
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
  
  This.WebBrowser.pWB_NavigateBack
  Debug.Assert (This.WebBrowser.pWB_GetCurrentURL = TARGET_PAGE_URL) Or (This.WebBrowser.pWB_GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", ""))

End Sub

Public Sub DragAndDrop()
  
  SelectListItem "Drag and Drop"
  Debug.Assert (This.WebBrowser.pWB_GetCurrentURL = TARGET_PAGE_URL & "/drag_and_drop") Or (This.WebBrowser.pWB_GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", "") & "/drag_and_drop")
  
  Dim i As Integer
  For i = 1 To 4
  
    Dim FirstItem As pLocator
    Set FirstItem = Factory.GetNewLocator
    With FirstItem
      .Initialise "FirstItem", This.RootWebArea, Children, pConditions, "AND(AriaRole, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
      .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Banner
      .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 3
      .Find 10
    End With

    Dim SecondItem As pLocator
    Set SecondItem = Factory.GetNewLocator
    With SecondItem
      .Initialise "SecondItem", This.RootWebArea, Children, pConditions, "AND(AriaRole, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
      .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Banner
      .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 4
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
  
  This.WebBrowser.pWB_NavigateBack
  Debug.Assert (This.WebBrowser.pWB_GetCurrentURL = TARGET_PAGE_URL) Or (This.WebBrowser.pWB_GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", ""))

End Sub

Public Sub FormAuthentication()
  SelectListItem "Form Authentication", "Login Page"
  FormAuthentication_Login
  FormAuthentication_Secure
  This.WebBrowser.pWB_NavigateBack 'Takes us back to the secure page
  This.WebBrowser.pWB_NavigateBack 'Takes us back to the login page
  This.WebBrowser.pWB_NavigateBack 'Finally takes us back to the homepage!
  Debug.Assert (This.WebBrowser.pWB_GetCurrentURL = TARGET_PAGE_URL) Or (This.WebBrowser.pWB_GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", ""))
End Sub
  
Private Sub FormAuthentication_Login()

  Debug.Assert (This.WebBrowser.pWB_GetCurrentURL = TARGET_PAGE_URL & "/login") Or (This.WebBrowser.pWB_GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", "") & "/login")

  Dim SubHeader As pLocator
  Set SubHeader = Factory.GetNewLocator
  With SubHeader
    .Initialise "SubHeader", This.RootWebArea, Children, pConditions, "AND(AriaRole, ClassName, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Heading
    .ClassName "subheader"
    .NameIs "This is where you can log into the secure area. Enter tomsmith for the username and SuperSecretPassword! for the password. If the information is wrong you should see error messages."
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 3
    Debug.Assert .ElementExists(5)
  End With

  Dim UsernameLabel As pLocator
  Set UsernameLabel = Factory.GetNewLocator
  With UsernameLabel
    .Initialise "UsernameLabel", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Description
    .NameIs "Username"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 4
    Debug.Assert .ElementExists(0)
  End With

  Dim UsernameTextBox As pLocator
  Set UsernameTextBox = Factory.GetNewLocator
  With UsernameTextBox
    .Initialise "UsernameTextBox", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.TextBox
    .NameIs "Username"
    '.Condition "LabeledBy", LabeledBy, IsTheString, "Username" ' This returns an element!
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 5
    Debug.Assert .ElementExists(0)
  End With

  Dim PasswordLabel As pLocator
  Set PasswordLabel = Factory.GetNewLocator
  With PasswordLabel
    .Initialise "PasswordLabel", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Description
    .NameIs "Password"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 6
    Debug.Assert .ElementExists(0)
  End With

  Dim PasswordTextBox As pLocator
  Set PasswordTextBox = Factory.GetNewLocator
  With PasswordTextBox
    .Initialise "FirstItem", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.TextBox
    .NameIs "Password"
    '.Condition "LabeledBy", LabeledBy, IsTheString, "Password" ' This returns an element!
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 7
    Debug.Assert .ElementExists(0)
  End With

  Dim LoginButton As pLocator
  Set LoginButton = Factory.GetNewLocator
  With LoginButton
    .Initialise "LoginButton", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameEndsWith, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Button
    'Name='? Login' - starts with a unicode character?
    .Condition "NameEndsWith", Name, EndsWithTheString, " Login"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 8
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
  Debug.Assert (This.WebBrowser.pWB_GetCurrentURL = TARGET_PAGE_URL & "/secure") Or (This.WebBrowser.pWB_GetCurrentURL = Replace(Replace(TARGET_PAGE_URL, "http://", ""), "https://", "") & "/secure")

  Dim SecureMesssage As pLocator
  Set SecureMesssage = Factory.GetNewLocator
  With SecureMesssage
    .Initialise "SecureMesssage", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Description
    .NameIs " You logged into a secure area!"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 2
    Debug.Assert .ElementExists(5)
  End With

  Dim SecureMesssageHyperlink As pLocator
  Set SecureMesssageHyperlink = Factory.GetNewLocator
  With SecureMesssageHyperlink
    .Initialise "SecureMesssageHyperlink", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Link
    .NameIs "×"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 3
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
    .Initialise "Heading", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Heading
    .NameIs "Secure Area"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 2
    Debug.Assert .ElementExists(5)
  End With

  Dim SubHeader As pLocator
  Set SubHeader = Factory.GetNewLocator
  With SubHeader
    .Initialise "SubHeader", This.RootWebArea, Children, pConditions, "AND(AriaRole, ClassName, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Heading
    .ClassName "subheader"
    .NameIs "Welcome to the Secure Area. When you are done click logout below."
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 3
    Debug.Assert .ElementExists(0)
  End With

  Dim Logout As pLocator
  Set Logout = Factory.GetNewLocator
  With Logout
    .Initialise "SubHeader", This.RootWebArea, Children, pConditions, "AND(AriaRole, NameIs, " & UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER & ")"
    .Condition "AriaRole", AriaRole, IsTheString, AriaRoles.Link
    .NameIs "Logout"
    .Condition UIAProps.POSITION_OF_ELEMENT_IN_TREESCOPE_COUNTER, 0, EqualsNumber, 4
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

