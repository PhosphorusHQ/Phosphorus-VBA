VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NodeTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder PPath
' =======================================================================
'  Phosphorus Test & Automation Suite
'  Copyright (c) 2025 Peter Jeffrey Gale
'
'  Licensed under the GNU GENERAL PUBLIC License
'  Full licence: see LICENCE in the distribution folder & main module
'  https://www.gnu.org/licenses/gpl-3.0.html#license-text
' =======================================================================
Option Explicit

Private This As pPath.Common

Public Sub Initialise(ByRef sharedthis As pPath.Common)
  Set This = sharedthis
End Sub

Public Sub ProcessNextNodeTest(myNextStep As pPath.Step)

  Dim intElementCounter As Integer
  Dim intNumberOfElements As Integer
  Dim eleCurrentUIElement As UIAutomationClient.IUIAutomationElement
  Dim strCurrentNavigationalPPath As String
  Dim strCurrentAttributeName As String
  Dim strCurrentNodeControlType As String
  
  intNumberOfElements = This.pPathReturnClass.GetNumberOfWorkingCopyOfCandidateElements(This.CurrentLocationPathExpressionCounter)
  For intElementCounter = 1 To intNumberOfElements
    Set eleCurrentUIElement = This.pPathReturnClass.GetWorkingCopyElement(This.CurrentLocationPathExpressionCounter, intElementCounter)
    strCurrentNavigationalPPath = This.pPathReturnClass.GetWorkingCopyNavigationalPPath(This.CurrentLocationPathExpressionCounter, intElementCounter)
    strCurrentAttributeName = This.pPathReturnClass.GetWorkingCopyAttributeName(This.CurrentLocationPathExpressionCounter, intElementCounter)
    strCurrentNodeControlType = This.AutomationDictionaries.ControlTypeIDs(eleCurrentUIElement.CurrentControlType)

    If (myNextStep.NodeTest = "*") Or _
       ((myNextStep.PrincipalNodeKind = PrincipalNodeKindType.Elements) And (strCurrentNodeControlType = myNextStep.NodeTest)) Or _
       ((myNextStep.PrincipalNodeKind = PrincipalNodeKindType.Attributes) And (strCurrentAttributeName = myNextStep.NodeTest)) Or _
       (myNextStep.NodeTest = "node()") Or _
       (myNextStep.NodeTest = "element()") _
    Then
      If myNextStep.NodeTestKind = "" Then
        This.pPathReturnClass.AddMatchingElement This.pPathReturnClass.GetCandidateElements(This.CurrentLocationPathExpressionCounter), eleCurrentUIElement, strCurrentAttributeName, strCurrentNavigationalPPath
      Else
        GetValuesOfDataType eleCurrentUIElement, myNextStep.NodeTestKind, strCurrentNavigationalPPath
      End If
    End If

    'Get attributes?
    If (myNextStep.NodeTest = "node()") Then
      This.Axes.GetAttributes eleCurrentUIElement, strCurrentNavigationalPPath & "/"
    End If
  
    'Get text?
    If (myNextStep.NodeTest = "text()") Then
      GetTextNode eleCurrentUIElement, strCurrentNavigationalPPath
    End If
  
  Next intElementCounter

End Sub

'https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controlpattern-ids
Public Sub GetTextNode(UIElement As UIAutomationClient.IUIAutomationElement, InitialPPath As String)

  If GetTextValue(UIElement) <> "" Then
    Dim strCurrentNavigationalPPath As String
    strCurrentNavigationalPPath = InitialPPath & "/text()"
    This.pPathReturnClass.AddMatchingElement This.pPathReturnClass.GetCandidateElements(This.CurrentLocationPathExpressionCounter), UIElement, "text()", strCurrentNavigationalPPath
  End If

End Sub

Public Function GetTextValue(UIElement As UIAutomationClient.IUIAutomationElement) As String

  Dim TextPattern As IUIAutomationTextPattern
  Dim TextContent As String
  TextContent = ""
  On Error Resume Next
  Set TextPattern = UIElement.GetCurrentPattern(UIA_TextPatternId)
  On Error GoTo 0

  If This.UnitTestingMode Then
    TextContent = "Text Text"
  Else
    If Not TextPattern Is Nothing Then
      TextContent = TextPattern.DocumentRange.GetText(-1) 'Get all text
    End If
  End If
  
  GetTextValue = TextContent
  
End Function

Public Sub GetValuesOfDataType(UIElement As UIAutomationClient.IUIAutomationElement, DataType As String, InitialPPath As String)
  
  Dim varValue As Variant
  Dim strElementDataType As String
  
  varValue = pPath.Utils.GetValue(UIElement)
  If varValue = "" Then
    strElementDataType = "null"
  Else
    If IsNumeric(varValue) Then
      If (varValue - VBA.Conversion.CLng(varValue)) = 0 Then
        strElementDataType = "integer"
      Else
        strElementDataType = "decimal"
      End If
    ElseIf varValue = True Or varValue = False Then
      strElementDataType = "boolean"
    ElseIf VBA.Strings.InStr(1, varValue, "https://") = 1 Or VBA.Strings.InStr(1, varValue, "http://") = 1 Then
      strElementDataType = "hyperlink"
    Else
      Dim dtDate As Date
      On Error Resume Next
      dtDate = CDate(varValue)
      If Err.Number = 13 Then
        strElementDataType = "string"
      Else
        strElementDataType = "date"
      End If
      On Error GoTo 0
    End If
    
    If (strElementDataType = DataType) Or _
       (strElementDataType = "integer" And DataType = "decimal") Then
      Dim strCurrentNavigationalPPath As String
      strCurrentNavigationalPPath = InitialPPath & "/@value"
      This.pPathReturnClass.AddMatchingElement This.pPathReturnClass.GetCandidateElements(This.CurrentLocationPathExpressionCounter), UIElement, "text()", strCurrentNavigationalPPath
    End If
  End If
  
End Sub


