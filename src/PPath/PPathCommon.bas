VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPathCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder PPath
Option Explicit

Public DebugMode As Boolean

Public AutomationDictionaries As UIAutomationDictionaries

Public UIAutomation As CUIAutomation
Public ApplicationRootUIElement As UIAutomationClient.IUIAutomationElement
Public TreeWalker As UIAutomationClient.IUIAutomationTreeWalker

Public UnitTestingMode As Boolean ' Needed as some parts of PPath are dynamic, e.g. PID
Public CurrentLocationPathExpressionCounter As Integer
Public Steps As PPathSteps
Public Axes As PPathAxes
Public NodeTests As PPathNodeTests
Public Predicates As PPathPredicates

Public PPathReturnClass As PPathReturnClass
