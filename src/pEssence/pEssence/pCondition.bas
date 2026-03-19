VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder pEssence
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Public ConditionName As String
Public UIAProperty As UIAProperties
Public UIAPropertyComparison As UIAPropertyComparisons
Public UIAPropertyValue As Variant

Public Function Evaluate(Element As IUIAutomationElement) As Boolean
  Dim ReturnValue As Boolean
  Select Case UIAPropertyComparison
    Case UIAPropertyComparisons.Equals
      'Cast from IUnknown type
       ReturnValue = ((VBA.Conversion.CVar(UIAProps.GetProperty(Element, UIAProperty)) = UIAPropertyValue))
    Case Else
      MsgBox "PJG?"
      Debug.Print "Unhandled UIAPropertyComparison " & UIAPropertyComparison
  End Select
  Evaluate = ReturnValue
End Function
