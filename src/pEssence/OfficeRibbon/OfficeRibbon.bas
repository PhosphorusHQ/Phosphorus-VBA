Attribute VB_Name = "OfficeRibbon"
'@Folder OfficeRibbon
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public Sub SelectTab(AppRootWindow As pLocator, TabName As String)
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator

  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//Pane[@Name=""Ribbon""]/Tab[@Name=""Ribbon Tabs""]//TabItem[@Name=""" & TabName & """]"
  With RootTestLocator
    .Initialise "RootTestLocator", AppRootWindow, None, By.pPath, pPathString
    .Find 10
    .Element.Click
  End With

End Sub

