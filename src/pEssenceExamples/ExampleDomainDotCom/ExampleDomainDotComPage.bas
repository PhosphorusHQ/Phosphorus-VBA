VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExampleDomainDotComPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder ExampleDomainDotCom
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Const WEB_APP_NAME = "Example.com"
Const TARGET_PAGE_URL = "https://www.example.com/"
Const TARGET_PAGE_TITLE = "Example Domain"
Const TARGET_PAGE_LINK = "https://iana.org/domains/example"


Private Type PageAttributes
  WebBrowser As Object
  RootWebArea As pLocator
  Heading As pLocator
  Description As pLocator
  Link As pLocator
End Type

Private This As PageAttributes

Private Sub Class_Initialize()
  Set This.WebBrowser = Factory.GetNewWebBrowser
End Sub

Private Sub Class_Terminate()
  Set This.WebBrowser = Nothing
  Set This.Heading = Nothing
  Set This.Description = Nothing
  Set This.Link = Nothing
End Sub

Public Sub Initialize()
  
  With This.WebBrowser
    .Start WEB_APP_NAME, TARGET_PAGE_URL, TARGET_PAGE_TITLE
    Set This.RootWebArea = .GetRootWebArea
  End With
  
  Set This.Heading = Factory.GetNewLocator
  With This.Heading
    .Initialise "Heading", This.RootWebArea, Children, pConditions, "AND(AriaRoleHeading, NameIs)"
    .AriaRoleHeading
    .NameIs "Example Domain"
  End With
  
  Set This.Description = Factory.GetNewLocator
  With This.Description
    .Initialise "Heading", This.RootWebArea, Children, pConditions, "AND(AriaRoleDescription, NameIs)"
    .AriaRoleDescription
    .NameIs "This domain is for use in documentation examples without needing permission. Avoid use in operations."
  End With
    
  Set This.Link = Factory.GetNewLocator
  With This.Link
    .Initialise "Link", This.RootWebArea, Children, pConditions, "AND(AriaRoleLink, NameIs)"
    .AriaRoleLink
    .NameIs "Learn more"
  End With
  
End Sub

Public Sub RunChecks()
  Debug.Assert This.RootWebArea.Element.GetProperty(Name) = TARGET_PAGE_TITLE
  Debug.Assert This.Heading.ElementExists(10) = True
  Debug.Assert This.Description.ElementExists(10) = True
  Debug.Assert This.Link.ElementExists(10) = True
  Debug.Assert This.Link.Element.GetValue() = TARGET_PAGE_LINK
End Sub
