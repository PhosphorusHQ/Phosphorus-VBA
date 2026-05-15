VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmpHilby 
   Caption         =   "pHilby - Phosphorus UIAutomation Spy Tool"
   ClientHeight    =   7380
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12588
   OleObjectBlob   =   "frmpHilby.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmpHilby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder pHilby
Option Explicit

Public AppName As String
Private WithEvents mcTree As pExternals.clsTreeView
Attribute mcTree.VB_VarHelpID = -1
Private AllTreeViewNodes As Scripting.Dictionary
Private iMaxNumberOfLevels As Integer

Private UnhandledPatterns As Scripting.Dictionary
Public HighlightSelectedNodeElement As Boolean

Private Sub UserForm_Initialize()
'See the Compile constant DebugMode in tools, VBAProject properties
'DebugMode=1 will enable the #If to Stop in Error handlers
'Normally this should be removed before distributing
    
  If Me.fraTreeControl.Font.Size < 4 Then
    Me.fraTreeControl.Font.Size = 4
  End If
    
  #If DEBUGMODE = 1 Then
    gFormInit = gFormInit + 1
  #End If

  Set AllTreeViewNodes = New Scripting.Dictionary
    
End Sub

Private Sub UserForm_Terminate()
  Window.ReleaseHighlighting
  #If DEBUGMODE = 1 Then
    gFormTerm = gFormTerm + 1
    ClassCounts
  #End If
  Set AllTreeViewNodes = Nothing
End Sub

Private Sub UserForm_Activate()
  Window.ActivateWindowByCaptionAndClassName Me.Caption, "ThunderDFrame"
  MsgBox "pHilby Ready!", vbExclamation, "Phosphorus"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  'Make sure all objects are destroyed
  If Not mcTree Is Nothing Then
    mcTree.TerminateTree
  End If
  Set mcTree = Nothing
End Sub

Public Sub LoadTreeView(RootUIAElement As IUIAutomationElement, Optional MaxNumberOfLevels As Integer)
  UnloadTreeView
  iMaxNumberOfLevels = MaxNumberOfLevels
  AllTreeViewNodes.RemoveAll
  Set UnhandledPatterns = New Scripting.Dictionary
  LoadUIATree RootUIAElement
End Sub

Private Sub UnloadTreeView()
  If Not mcTree Is Nothing Then
    mcTree.NodesClear
    Set mcTree = Nothing
  End If
End Sub

Private Sub LoadUIATree(RootUIAElement As IUIAutomationElement)

  Dim cRoot As clsNode
  Dim cNode As clsNode
  Dim cExtraNode As clsNode
  Dim i As Long
  Dim k As Long

  Set mcTree = pExternals.Factory.GetTreeView
    
  With mcTree
    'Pass the frame to the TreeControl of the treeview class
    Set .TreeControl = Me.fraTreeControl
    Call .NodesClear
    .AppName = Me.AppName
    .EnableLabelEdit() = False
    .ShowLines = True
    .ShowExpanders = True
    .NarratorReaderControl = True
    LoadRootElementAndAllDescendants RootUIAElement
    'Create the node controls and display the tree
    .Refresh
  End With

End Sub

Private Sub LoadRootElementAndAllDescendants(UIAElement As IUIAutomationElement)

  Dim key As String
  key = "1"
  AddItemToAllTreeViewNodes key, UIAElement
  
  Dim RuntimeId As String
  RuntimeId = UIACommon.GetElementRuntimeId(UIAElement)

  ' Add the root node and make it bold
  Dim Root As clsNode
  Set Root = mcTree.AddRoot(sKey:=key, vCaption:=GetCaption(UIAElement))
  Root.Bold = True
  Root.ControlTipText = "Key: 1; " & "RuntimeID: " & RuntimeId
  Root.Tag = UIAElement.CurrentName
  Root.Expanded = True

  LoadAllRootElementDescendants key, UIAElement, 1, Root

End Sub

Private Sub LoadAllRootElementDescendants(RootKey As String, UIAElement As IUIAutomationElement, LevelCounter As Integer, ParentNode As clsNode)

  'If iMaxNumberOfLevels=0 show all levels
  If iMaxNumberOfLevels = 0 Or (iMaxNumberOfLevels >= LevelCounter) Then
        
    Dim AllElements As IUIAutomationElementArray
    Set AllElements = UIAElement.FindAll(TreeScope.Children, UIA.CreateTrueCondition)
    
    If AllElements.Length = 0 Then
      ' No children, just exit
      Exit Sub
    End If
    
    Dim i As Long
    Dim SubKey As String
    Dim RuntimeId As String
    Dim CurrentUIAElement As IUIAutomationElement
    Dim ChildNode As clsNode
    For i = 0 To AllElements.Length - 1
      SubKey = RootKey & "." & (i + 1)
      Set CurrentUIAElement = AllElements.GetElement(i)
      RuntimeId = UIACommon.GetElementRuntimeId(CurrentUIAElement)
      AddItemToAllTreeViewNodes SubKey, CurrentUIAElement
      Set ChildNode = ParentNode.AddChild(sKey:=SubKey, vCaption:=GetCaption(CurrentUIAElement))
      ChildNode.ControlTipText = "Key: " & SubKey & "; " & "RuntimeID: " & RuntimeId
      ChildNode.Tag = CurrentUIAElement.CurrentName
      ChildNode.Expanded = False
      LoadAllRootElementDescendants SubKey, CurrentUIAElement, LevelCounter + 1, ChildNode
    Next i
    
  End If

End Sub

Private Function GetCaption(UIAElement As IUIAutomationElement) As String
  GetCaption = UIAElement.CurrentName & " (" & UIAProps.GetControlTypeName(UIAElement.CurrentControlType) & ")"
End Function

Private Sub AddItemToAllTreeViewNodes(TreeNodeKey As String, UIAElement As IUIAutomationElement)
  'Note that we have to store the properties here as the UIAElement may not be live after pHulby is loaded
  Dim coll As New Collection
  coll.Add Item:=UIAElement, key:="UIAElement"
  coll.Add Item:=GetUIAElementProperties(UIAElement), key:="AllProperties"
  coll.Add Item:=GetUIAElementPatterns(UIAElement), key:="AllPatterns"
  AllTreeViewNodes.Add TreeNodeKey, coll
End Sub

Private Function GetUIAElementProperties(UIAElement As IUIAutomationElement) As String
  
  'https://excelmacromastery.com/vba-dictionary/ Sorting by keys
  Dim arrList As Object
  Set arrList = CreateObject("System.Collections.ArrayList")
  
  'Populate the unsorted dictionary
  Dim PropertyId As UIAProperties
  Dim PropertyName As String
  Dim PropertyStringValue As String
  Dim UnsortedDictionary As New Scripting.Dictionary
  For PropertyId = 30000 To 30500
    PropertyName = UIAProps.GetPropertyName(PropertyId)
    If PropertyName <> "" Then
      PropertyStringValue = UIAProps.GetPropertyValueAsString(UIAElement, PropertyId)
      If PropertyStringValue = "" Then
        ' Leave as is!
      ElseIf PropertyStringValue = "True" Or PropertyStringValue = "False" Then
        ' Leave as is!
      ElseIf IsNumeric(PropertyStringValue) Then
        ' Leave as is!
      Else
        PropertyStringValue = Chr(34) & PropertyStringValue & Chr(34)
      End If
      'TODO: Display all properties?
      UnsortedDictionary.Add PropertyName, PropertyStringValue
      arrList.Add PropertyName
    End If
  Next PropertyId
  
  'Sort array
  arrList.Sort
  
  'Populate the unsorted text
  Dim AllPropertiesText As String
  Dim key As Variant
  For Each key In arrList
    If AllPropertiesText <> "" Then
       AllPropertiesText = AllPropertiesText & vbCrLf
    End If
    AllPropertiesText = AllPropertiesText & key & ": " & UnsortedDictionary(key)
  Next key

  ' Clean up
  Set arrList = Nothing
  Set UnsortedDictionary = Nothing
    
  'Set return value
  GetUIAElementProperties = AllPropertiesText
  
End Function

Private Function GetUIAElementPatterns(UIAElement As IUIAutomationElement) As String

  'https://excelmacromastery.com/vba-dictionary/ Sorting by keys
  Dim arrList As Object
  Set arrList = CreateObject("System.Collections.ArrayList")
  
  'Populate the unsorted array list
  Dim PatternId As UIAPatterns
  Dim PatternText As String
  For PatternId = 10000 To 10050
    Dim PatternName As String
    PatternName = UIAPatts.GetPatternName(PatternId)
    If PatternName <> "" Then
      PatternText = GetPatternText(UIAElement, PatternName, PatternId)
      If PatternText <> "" Then
        arrList.Add PatternText
      End If
    End If
  Next PatternId
  
  'Sort array
  arrList.Sort
  
  'Populate the unsorted text
  Dim AllPatternsText As String
  Dim key As Variant
  For Each key In arrList
    If AllPatternsText <> "" Then
       AllPatternsText = AllPatternsText & vbCrLf
    End If
    AllPatternsText = AllPatternsText & key
  Next key

  ' Clean up
  Set arrList = Nothing
    
  'Set return value
  GetUIAElementPatterns = AllPatternsText
  
End Function

Private Function GetPatternText(UIAElement As IUIAutomationElement, PatternName As String, PatternId As UIAPatterns) As String
  Dim R As String
  If UIAPatts.HasPattern(UIAElement, PatternId) Then
    Select Case PatternId
      Case UIAPatterns.Invoke
        R = "" 'No current settings
      Case UIAPatterns.LegacyIAccessible
        R = GetPatternText_LegacyIAccessible(UIAElement)
      Case UIAPatterns.Selection
        R = GetPatternText_Selection(UIAElement)
      Case UIAPatterns.Text
        R = "" 'No current settings
      Case UIAPatterns.Toggle
        R = GetPatternText_Toggle(UIAElement)
      Case UIAPatterns.Value
        R = GetPatternText_Value(UIAElement)
      Case UIAPatterns.Window
        R = GetPatternText_Window(UIAElement)
      Case Else
        If Not UnhandledPatterns.Exists(PatternId) Then
          'MsgBox PatternName & " Pattern not handled!", vbInformation, "Phosphorus - pHilby"
          Debug.Print PatternName & " Pattern not handled!"
          UnhandledPatterns.Add PatternId, PatternName
        End If
    End Select
    If R = "" Then
      R = PatternName & vbCrLf
    Else
      R = PatternName & vbCrLf & String(Len(PatternName), "=") & R
    End If
  End If
  GetPatternText = R
End Function

Private Function GetPatternText_LegacyIAccessible(UIAElement As IUIAutomationElement) As String
  Dim Pattern As IUIAutomationLegacyIAccessiblePattern
  Set Pattern = UIAElement.GetCurrentPattern(UIAPatterns.LegacyIAccessible)
  GetPatternText_LegacyIAccessible = vbCrLf & _
     "  Child: " & Pattern.CurrentChildId & vbCrLf & _
     "  DefaultAction: " & CStr(Pattern.CurrentDefaultAction) & vbCrLf & _
     "  Description: " & Pattern.CurrentDescription & vbCrLf & _
     "  Help: " & Pattern.CurrentHelp & vbCrLf & _
     "  KeyboardShortcut: " & Pattern.CurrentKeyboardShortcut & vbCrLf & _
     "  Name: " & Pattern.CurrentName & vbCrLf & _
     "  Value: " & Pattern.CurrentValue & vbCrLf
End Function

Private Function GetPatternText_Selection(UIAElement As IUIAutomationElement) As String
  Dim Pattern As IUIAutomationSelectionPattern
  Set Pattern = UIAElement.GetCurrentPattern(UIAPatterns.Selection)
  GetPatternText_Selection = vbCrLf & _
     "  CanSelectMultiple: " & (Pattern.CurrentCanSelectMultiple = 1) & vbCrLf & _
     "  SelectionRequired: " & (Pattern.CurrentIsSelectionRequired = 1) & vbCrLf
End Function

Private Function GetPatternText_Toggle(UIAElement As IUIAutomationElement) As String
  Dim Pattern As IUIAutomationTogglePattern
  Set Pattern = UIAElement.GetCurrentPattern(UIAPatterns.Toggle)
  GetPatternText_Toggle = vbCrLf & _
     "  ToggleState: " & (Pattern.CurrentToggleState = 1) & vbCrLf
End Function

Private Function GetPatternText_Value(UIAElement As IUIAutomationElement) As String
  Dim Pattern As IUIAutomationValuePattern
  Set Pattern = UIAElement.GetCurrentPattern(UIAPatterns.Value)
  GetPatternText_Value = vbCrLf & _
     "  ReadOnly: " & (Pattern.CurrentIsReadOnly = 1) & vbCrLf & _
     "  CurrentValue: " & CStr(Pattern.CurrentValue) & vbCrLf
End Function

Private Function GetPatternText_Window(UIAElement As IUIAutomationElement) As String
  Dim Pattern As IUIAutomationWindowPattern
  Set Pattern = UIAElement.GetCurrentPattern(UIAPatterns.Window)
  GetPatternText_Window = vbCrLf & _
     "  CanMaximize: " & (Pattern.CurrentCanMaximize = 1) & vbCrLf & _
     "  CanMinimize: " & (Pattern.CurrentCanMinimize = 1) & vbCrLf & _
     "  Modal: " & (Pattern.CurrentIsModal = 1) & vbCrLf & _
     "  Topmost: " & (Pattern.CurrentIsTopmost = 1) & vbCrLf & _
     "  WindowInteractionState: " & UIAProps.GetWindowInteractionStateName(Pattern.CurrentWindowInteractionState) & vbCrLf & _
     "  WindowVisualState: " & UIAProps.GetWindowVisualStateName(Pattern.CurrentWindowVisualState) & vbCrLf
End Function

Private Sub cmdSearch_Click()
  Dim SearchText As String
  SearchText = VBA.Interaction.InputBox("Input the text to search for...", "pHilby")
  If SearchText <> "" Then
    Dim Node As clsNode
    Dim MatchingNodes() As clsNode
    Dim Count As Integer
    For Each Node In mcTree.Nodes
    If Node.Tag = SearchText Then
      Node.Bold = True
      Count = Count + 1
      ReDim Preserve MatchingNodes(Count)
      Set MatchingNodes(Count) = Node
    Else
      Node.Bold = False
    End If
    Node.Expanded = False
    Next Node
    If Count = 0 Then
       MsgBox "No matching nodes found.", vbCritical, "pHilby"
    Else
      Dim i As Integer
      For i = 1 To Count
        Dim ExpandableNode As clsNode
        Set ExpandableNode = MatchingNodes(i)
        While Not ExpandableNode Is Nothing
          ExpandableNode.Expanded = True
          Set ExpandableNode = ExpandableNode.ParentNode
        Wend
      Next i
    End If
    mcTree.Refresh
  End If
End Sub

Private Sub mcTree_Click(cNode As pExternals.clsNode)
  'This gets fired when a node is clicked
  txtProperties = AllTreeViewNodes(cNode.key)("AllProperties")
  txtPatterns = AllTreeViewNodes(cNode.key)("AllPatterns")
  Window.ReleaseHighlighting
  If HighlightSelectedNodeElement Then
    Dim UIAElement As IUIAutomationElement
    Set UIAElement = AllTreeViewNodes(cNode.key)("UIAElement")
    If IsAlive(UIAElement) Then
      Window.HighlightElement UIAElement
      Application.OnTime Now + TimeValue("00:00:05"), "pHilby.ReleaseHighlighting"
    Else
      MsgBox "This element is not longer available!", vbCritical, "pHilby"
    End If
  End If
End Sub

Private Function IsAlive(UIAElement As IUIAutomationElement) As Boolean
  On Error Resume Next
  Dim pid As Long
  pid = UIAElement.CurrentProcessId  'any property access will fail if stale
  IsAlive = (Err.Number = 0) And (pid > 0)
  On Error GoTo 0
End Function

Private Sub chkHighlightSelectedNodeElement_Click()
  HighlightSelectedNodeElement = chkHighlightSelectedNodeElement.Value
  Window.HighlightElements = chkHighlightSelectedNodeElement.Value
End Sub

Private Sub cbReleaseHighlighting_Click()
  Window.ReleaseHighlighting
End Sub


