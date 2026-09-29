Attribute VB_Name = "HelloWorld"
'@Folder HelloWorld
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Function HelloWorldWideWeb() As Boolean

  Dim Succeeded As Boolean
  Succeeded = False

  Dim WebBrowser As Object
  Dim SearchBox As pLocator
  Dim SearchButton As pLocator

  On Error GoTo ErrorHandler
  
  If Not RunningAllExamples Then
    'pPath always requires the Logger class
    Phosphorus.Log4PStatic.GetLogger
'    WebBrowserCommon.GetInternetSpeeds
WebBrowserCommon.SetDummyMobileDataInternetSpeeds
    Factory.CurrentWebBrowserType = 1
  End If
  
  Set WebBrowser = Factory.GetNewWebBrowser
  WebBrowser.Start "Hello World!", "https://www.google.com", "Google"

  Set SearchBox = Factory.GetNewLocator
  With SearchBox
    .Initialise _
      "SearchBox", WebBrowser.GetRootWebArea, TreeScope_Descendants, By.pConditions, "OR(AND(AriaRoleComboBox, NameIsSearch),AND(ControlTypeEdit, NameIsGoogleSearch))"
      .Condition "AriaRoleComboBox", UIAProperties.AriaRole, IsTheString, AriaRoles.ComboBox
      .Condition "NameIsSearch", Name, IsTheString, "Search"
      .Condition "ControlTypeEdit", UIAProperties.ControlType, EqualsNumber, UIAControlTypeIDs.Edit
      .Condition "NameIsGoogleSearch", Name, IsTheString, "Google Search"
  End With
        
  Set SearchButton = Factory.GetNewLocator
  With SearchButton
    .Initialise _
      "SearchButton", WebBrowser.GetRootWebArea, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleButton, NameIs)": .AriaRoleButton: .NameIs "Google Search"
      ': .PositionInMatchingSet 2
  End With

  SearchBox.Element.SetValue "Hello World"
  SearchButton.Element.Click

  If RunningAllExamples Then
    Debug.Print "Behold The World Wide Web!"
  Else
    MsgBox "Behold The World Wide Web!"
    Phosphorus.Log4PStatic.CloseLogger
  End If
  Snooze 1000
  Succeeded = True
  GoTo ExitSub
   
ErrorHandler:
  If RunningAllExamples Then
    Debug.Print Err.Description & " (Error Number #" & Err.Number & ") for " & Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser in 'ExampleDomainDotCom' Example!"
  Else
    MsgBox Err.Description & " (Error Number #" & Err.Number & ") for " & Factory.GetWebBrowserName(Factory.CurrentWebBrowserType) & " Web Browser in 'ExampleDomainDotCom' Example!", vbCritical
  End If
  GoTo ExitSub
  
ExitSub:
  Set WebBrowser = Nothing
  Set SearchBox = Nothing
  Set SearchButton = Nothing
  If Not RunningAllExamples Then
    Window.HighlightElements = False
  End If
  
  HelloWorldWideWeb = Succeeded

End Function
