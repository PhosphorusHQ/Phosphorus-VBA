VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPathMatchingElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder PPath
Option Explicit
Option Base 1

Private MatchingElements() As UIAutomationClient.IUIAutomationElement
Private AttributeNames() As String
Private NavigationalPPaths() As String

Public Sub AddMatchingElement( _
  longMatchingElementNumber As Long, _
  ByRef MatchingElement As UIAutomationClient.IUIAutomationElement, _
  AttributeName As String, _
  NavigationalXPath As String)
  
  ReDim Preserve MatchingElements(longMatchingElementNumber)
  Set MatchingElements(longMatchingElementNumber) = MatchingElement
  
  ReDim Preserve AttributeNames(longMatchingElementNumber)
  AttributeNames(longMatchingElementNumber) = AttributeName
  
  ReDim Preserve NavigationalPPaths(longMatchingElementNumber)
  NavigationalPPaths(longMatchingElementNumber) = NavigationalXPath

End Sub

Public Sub EraseAllElements()
  Erase MatchingElements
  Erase AttributeNames
  Erase NavigationalPPaths
End Sub

Public Function GetMatchingElement(intElementNumber As Integer) As UIAutomationClient.IUIAutomationElement
  Set GetMatchingElement = MatchingElements(intElementNumber)
End Function

Public Function GetAttributeNames() As String()
  GetAttributeNames = AttributeNames
End Function

Public Function GetAttributeName(intElementNumber As Integer) As String
  GetAttributeName = AttributeNames(intElementNumber)
End Function

Public Function GetNavigationalPPaths() As String()
  GetNavigationalPPaths = NavigationalPPaths
End Function

Public Function GetNavigationalPPath(intElementNumber) As String
  GetNavigationalPPath = NavigationalPPaths(intElementNumber)
End Function

Public Function GetNumberOfMatchingElements() As Long
  GetNumberOfMatchingElements = Utils.GetSizeOfArray(MatchingElements)
End Function

Public Function GetMatchingElementsArray() As UIAutomationClient.IUIAutomationElement()
  GetMatchingElementsArray = MatchingElements
End Function

Public Function GetNavigationalXPathsArray() As String()
  GetNavigationalXPathsArray = NavigationalPPaths
End Function

Public Sub ClearDownAllArrays()
  Erase MatchingElements
  Erase AttributeNames
  Erase NavigationalPPaths
End Sub

Public Sub SetMatchingElements(NewMatchingElements() As UIAutomationClient.IUIAutomationElement)
  MatchingElements = NewMatchingElements
End Sub

Public Sub SetAttributeNames(NewAttributeNames() As String)
  AttributeNames = NewAttributeNames
End Sub

Public Sub SetNavigationalPPaths(NewNavigationalPPaths() As String)
  NavigationalPPaths = NewNavigationalPPaths
End Sub


