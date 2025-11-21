VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder pPath
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public AutomationDictionaries As UIAutomationDictionaries

Public UIAutomation As CUIAutomation
Public ApplicationRootUIElement As UIAutomationClient.IUIAutomationElement
Public TreeWalker As UIAutomationClient.IUIAutomationTreeWalker

Public UnitTestingMode As Boolean ' Needed as some parts of PPath are dynamic, e.g. PID
Public CurrentLocationPathExpressionCounter As Integer
Public Steps As pPath.Steps
Public Axes As pPath.Axes
Public NodeTests As pPath.NodeTests
Public Predicates As pPath.Predicates

Public pPathReturnClass As pPath.ReturnClass

