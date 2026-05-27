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

Sub HelloWorldWideWeb()

  Dim WebBrowser As Object
  Dim SearchBox As pLocator
  Dim SearchButton As pLocator
  
  If Not RunningAllExamples Then
    Factory.CurrentWebBrowserType = 1
  End If
  
  Set WebBrowser = Factory.GetNewWebBrowser
  WebBrowser.Start "Hello World!", "https://www.google.com", "Google"
  
  Set SearchBox = Factory.GetNewLocator
  With SearchBox
    .Initialise _
      "SearchBox", WebBrowser.GetRootWebArea, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleComboBox, NameIs)": .AriaRoleComboBox: .NameIs "Search"
  End With
  
  Set SearchButton = Factory.GetNewLocator
  With SearchButton
    .Initialise _
      "SearchButton", WebBrowser.GetRootWebArea, TreeScope_Descendants, By.pConditions, _
      "AND(AriaRoleButton, NameIs)": .AriaRoleButton: .NameIs "Google Search": .PositionInMatchingSet 2
  End With
  
  SearchBox.Element.SetValue "Hello World"
  SearchButton.Element.Click

  If RunningAllExamples Then: Debug.Print "Behold The World Wide Web!": Else: MsgBox "Behold The World Wide Web!"
  Snooze 1000
  
  Set WebBrowser = Nothing
  Set SearchBox = Nothing
  Set SearchButton = Nothing
  
End Sub
