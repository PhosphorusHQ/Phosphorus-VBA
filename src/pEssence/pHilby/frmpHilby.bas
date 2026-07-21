VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmpHilby 
   Caption         =   "pHilby - Phosphorus UIAutomation Spy Tool"
   ClientHeight    =   7668
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15096
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

Const FormName As String = "pHilby"
Public AppName As String
Private WithEvents mcTree As pExternals.clsTreeView
Attribute mcTree.VB_VarHelpID = -1
Private AllTreeViewNodes As Scripting.Dictionary

Private mRootUIAElement As IUIAutomationElement
Private mMaxNumberOfLevels As Integer
Private mDelayLoadingInSeconds As Integer

Private ActiveNode As pExternals.clsNode
Private UnhandledPatterns As Scripting.Dictionary
Public HighlightSelectedNodeElement As Boolean

Private RootImageFolder As String
Private SearchExecuted As Boolean
Private pPathContextNodes() As pExternals.clsNode
Private pPathContextNodesBackColour As Long

'Private pPathContextNodesInitialpPaths() As String

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

  RootImageFolder = ThisWorkbook.Path & "\images\pHilby\"
  cbExpand.Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "expansion.png")
  cbCollapse.Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "collapse.png")
  cbHighlightSelectedNodeElement.Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "marker.png")
  cbUIAPolling.Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "hand.png")
  cbExport.Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "xls.png")
  cbExportToNode.Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "xls2.png")
  cbSetCurrentAsRootNode.Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "root-directory.png")
  cbSetParentAsRootNode.Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "root-directory2.png")
  cbReload.Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "hacker.png")
  cbSearch.Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "loupe.png")
  cbAddpPathContextNode.Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "addcontextnode.png")
  cbClearAllpPathContextNodes.Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "deletenodes.png")
  cbExecutepPath.Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "control-system.png")
  cbResetSearch.Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "reset.png")
    
  SearchExecuted = False
  cbResetSearch.Enabled = False
  pPathContextNodesBackColour = VBA.ColorConstants.vbCyan
End Sub

Private Sub UserForm_Terminate()
  Window.ReleaseHighlighting
  #If DEBUGMODE = 1 Then
    gFormTerm = gFormTerm + 1
    ClassCounts
  #End If
  Set AllTreeViewNodes = Nothing
  pHilby.StoppHilbyUIAUIAPolling
End Sub

Private Sub UserForm_Activate()
  Window.ActivateWindowByCaptionAndClassName Me.Caption, "ThunderDFrame"
  AppActivate Me.Caption
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  'Make sure all objects are destroyed
  If Not mcTree Is Nothing Then
    mcTree.TerminateTree
  End If
  Set mcTree = Nothing
End Sub

Public Sub LoadTreeView(RootUIAElement As IUIAutomationElement, Optional MaxNumberOfLevels As Integer, Optional DelayLoadingInSeconds As Integer = 0)
  mDelayLoadingInSeconds = DelayLoadingInSeconds
  txtDelayLoading.Value = mDelayLoadingInSeconds
  If DelayLoadingInSeconds > 0 Then
    Application.Wait VBA.DateTime.DateAdd("s", mDelayLoadingInSeconds, Now)
  End If
  UnloadTreeView
  Set mRootUIAElement = RootUIAElement
  mMaxNumberOfLevels = MaxNumberOfLevels
  txtMaxNumberOfLevels.Value = mMaxNumberOfLevels
  AllTreeViewNodes.RemoveAll
  Set UnhandledPatterns = New Scripting.Dictionary
  LoadUIATree RootUIAElement
  SetRootAsActiveNode
End Sub

Private Sub SetRootAsActiveNode()
  Set ActiveNode = mcTree.RootNodes(1)
  ActiveNode.Bold = True
  mcTree_Click ActiveNode
End Sub

Private Sub UnloadTreeView()
  If Not mcTree Is Nothing Then
    mcTree.NodesClear
    Set mcTree = Nothing
  End If
End Sub

Private Sub LoadUIATree(RootUIAElement As IUIAutomationElement)
  
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
    .ImageAdd WindowsImageAcquisition.LoadImage(RootImageFolder & "root-directory24px.png"), "RootNodeOpen"
    .ImageAdd WindowsImageAcquisition.LoadImage(RootImageFolder & "closedfolder24px.png"), "RootNodeClosed"
    .ImageAdd WindowsImageAcquisition.LoadImage(RootImageFolder & "closedfolder.png"), "ClosedFolder"
    .ImageAdd WindowsImageAcquisition.LoadImage(RootImageFolder & "root-directory16px.png"), "OpenFolder"
    .ImageAdd WindowsImageAcquisition.LoadImage(RootImageFolder & "point.png"), "Attribute"
    .ImageAdd WindowsImageAcquisition.LoadImage(RootImageFolder & "txt-file.png"), "Text"
    LoadRootElementAndAllDescendants RootUIAElement
    'Create the node controls and display the tree
    .Refresh
  End With

End Sub

Private Sub LoadRootElementAndAllDescendants(UIAElement As IUIAutomationElement)

  Dim Key As String
  Dim Path As String
  Dim RuntimeId As String
  Key = "1"
  Path = "/" & UIAProps.GetControlTypeName(UIAElement.CurrentControlType)
  RuntimeId = pPath.RuntimeIDs.GetElementRuntimeId(UIAElement)
  'Don't add any elements without a RuntimeId
  If RuntimeId <> "" Then
  
    AddItemToAllTreeViewNodes Key, UIAElement, Path, RuntimeId

    ' Add the root node and make it bold
    Dim Root As clsNode
    Set Root = mcTree.AddRoot(sKey:=Key, vCaption:=GetCaption(UIAElement), vImageMain:="RootNodeClosed", vImageExpanded:="RootNodeOpen")

    Root.Bold = True
    Root.ControlTipText = "Key: 1; " & "RuntimeId: " & RuntimeId
    Root.Tag = UIAElement.CurrentName
    Root.Expanded = True

    LoadAllRootElementDescendants Key, Path, UIAElement, 1, Root

  End If
  
End Sub

Private Sub LoadAllRootElementDescendants(RootKey As String, RootPath As String, UIAElement As IUIAutomationElement, LevelCounter As Integer, ParentNode As clsNode)

  'If iMaxNumberOfLevels=0 show all levels
  If mMaxNumberOfLevels = 0 Or (mMaxNumberOfLevels >= LevelCounter) Then
        
    Dim AllElements As IUIAutomationElementArray
    Set AllElements = UIAElement.FindAll(TreeScope.Children, UIA.CreateTrueCondition)
    
    If AllElements.Length = 0 Then
      ' No children, just exit
      Exit Sub
    End If
 
    'Keep a tally the count of all control types at the current level
    Dim ControlTypeCounts As Scripting.Dictionary
    Dim CurrentControlTypeName As String
    Dim CurrentControlTypeCount As Integer
    Set ControlTypeCounts = New Scripting.Dictionary
    
    Dim i As Long
    Dim SubKey As String
    Dim SubPath As String
    Dim RuntimeId As String
    Dim CurrentUIAElement As IUIAutomationElement
    Dim ChildNode As clsNode
    For i = 0 To AllElements.Length - 1
      SubKey = RootKey & "." & (i + 1)
      Set CurrentUIAElement = AllElements.GetElement(i)
      CurrentControlTypeName = UIAProps.GetControlTypeName(CurrentUIAElement.CurrentControlType)
      If ControlTypeCounts.Exists(CurrentControlTypeName) Then
        CurrentControlTypeCount = ControlTypeCounts(CurrentControlTypeName)
        ControlTypeCounts.Remove CurrentControlTypeName
        CurrentControlTypeCount = CurrentControlTypeCount + 1
      Else
        CurrentControlTypeCount = 1
      End If
      ControlTypeCounts.Add CurrentControlTypeName, CurrentControlTypeCount
      SubPath = RootPath & "/" & CurrentControlTypeName & "[" & CStr(CurrentControlTypeCount) & "]"
      RuntimeId = pPath.RuntimeIDs.GetElementRuntimeId(CurrentUIAElement)
      'Don't add any elements without a RuntimeId
      If RuntimeId <> "" Then
        AddItemToAllTreeViewNodes SubKey, CurrentUIAElement, SubPath, RuntimeId
        Set ChildNode = ParentNode.AddChild(sKey:=SubKey, vCaption:=GetCaption(CurrentUIAElement), vImageMain:="ClosedFolder", vImageExpanded:="OpenFolder")
        ChildNode.ControlTipText = "Key: " & SubKey & "; " & "RuntimeId: " & RuntimeId
        ChildNode.Tag = CurrentUIAElement.CurrentName
        ChildNode.Expanded = False
        LoadAllRootElementDescendants SubKey, SubPath, CurrentUIAElement, LevelCounter + 1, ChildNode
      End If
    Next i
    
  End If

End Sub

Private Function GetCaption(UIAElement As IUIAutomationElement) As String
  GetCaption = UIAElement.CurrentName & " (" & UIAProps.GetControlTypeName(UIAElement.CurrentControlType) & ")"
End Function

Private Sub AddItemToAllTreeViewNodes(TreeNodeKey As String, UIAElement As IUIAutomationElement, Path As String, RuntimeId As String)
  'Note that we have to store the properties here as the UIAElement may not be live after pHulby is loaded
  Dim Coll As New Collection
  Coll.Add Item:=UIAElement, Key:="UIAElement"
  Coll.Add Item:=Path, Key:="Path"
  Coll.Add Item:=RuntimeId, Key:="RuntimeId"
  Coll.Add Item:=UIAProps.GetPropertyValueAsString(UIAElement, UIAProperties.Name), Key:="Name"
  Coll.Add Item:=GetUIAElementProperties(TreeNodeKey, UIAElement, Path), Key:="AllProperties"
  Coll.Add Item:=GetUIAElementPatterns(UIAElement), Key:="AllPatterns"
  AllTreeViewNodes.Add TreeNodeKey, Coll
End Sub

Private Function GetUIAElementProperties(TreeNodeKey As String, UIAElement As IUIAutomationElement, Path As String) As String

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
      ElseIf PropertyId = UIAProperties.BoundingRectangle Then
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
  AllPropertiesText = "Absolute Path: " & Path & vbCrLf
    
  Dim Key As Variant
  For Each Key In arrList
     AllPropertiesText = AllPropertiesText & vbCrLf
    AllPropertiesText = AllPropertiesText & Key & ":" & vbTab & UnsortedDictionary(Key)
  Next Key

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
  Dim Key As Variant
  For Each Key In arrList
    If AllPatternsText <> "" Then
       AllPatternsText = AllPatternsText & vbCrLf
    End If
    AllPatternsText = AllPatternsText & Key
  Next Key

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
        R = GetPatternText_Text(UIAElement)
      Case UIAPatterns.Toggle
        R = GetPatternText_Toggle(UIAElement)
      Case UIAPatterns.Value
        R = GetPatternText_Value(UIAElement)
      Case UIAPatterns.Window
        R = GetPatternText_Window(UIAElement)
      Case Else
        If Not UnhandledPatterns.Exists(PatternId) Then
          'MsgBox PatternName & " Pattern not handled!", vbInformation, "Phosphorus - pHilby"
          Debug.Print PatternName & " Pattern not handled! (" & UIAElement.CurrentName & ")"
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
     "  Child:" & vbTab & Pattern.CurrentChildId & vbCrLf & _
     "  DefaultAction:" & vbTab & CStr(Pattern.CurrentDefaultAction) & vbCrLf & _
     "  Description:" & vbTab & Pattern.CurrentDescription & vbCrLf & _
     "  Help:" & vbTab & Pattern.CurrentHelp & vbCrLf & _
     "  KeyboardShortcut:" & vbTab & Pattern.CurrentKeyboardShortcut & vbCrLf & _
     "  Name:" & vbTab & Pattern.CurrentName & vbCrLf & _
     "  Value:" & vbTab & Pattern.CurrentValue & vbCrLf
End Function

Private Function GetPatternText_Selection(UIAElement As IUIAutomationElement) As String
  Dim Pattern As IUIAutomationSelectionPattern
  Set Pattern = UIAElement.GetCurrentPattern(UIAPatterns.Selection)
  GetPatternText_Selection = vbCrLf & _
     "  CanSelectMultiple:" & vbTab & (Pattern.CurrentCanSelectMultiple = 1) & vbCrLf & _
     "  SelectionRequired:" & vbTab & (Pattern.CurrentIsSelectionRequired = 1) & vbCrLf
End Function

Private Function GetPatternText_Text(UIAElement As IUIAutomationElement) As String
'  Dim Pattern As IUIAutomationTextPattern
'  Set Pattern = UIAElement.GetCurrentPattern(UIAPatterns.UIA_TextPatternId)
'  GetPatternText_Text = vbCrLf & _
'     "  Text:" & vbTab & Pattern.DocumentRange.GetText(-1) & vbCrLf
  GetPatternText_Text = vbCrLf & _
     "  Text:" & vbTab & UIAPatts.GetTextValue(UIAElement) & vbCrLf
End Function

Private Function GetPatternText_Toggle(UIAElement As IUIAutomationElement) As String
  Dim Pattern As IUIAutomationTogglePattern
  Set Pattern = UIAElement.GetCurrentPattern(UIAPatterns.Toggle)
  GetPatternText_Toggle = vbCrLf & _
     "  ToggleState:" & vbTab & (Pattern.CurrentToggleState = 1) & vbCrLf
End Function

Private Function GetPatternText_Value(UIAElement As IUIAutomationElement) As String
  Dim Pattern As IUIAutomationValuePattern
  Set Pattern = UIAElement.GetCurrentPattern(UIAPatterns.Value)
  GetPatternText_Value = vbCrLf & _
     "  ReadOnly:" & vbTab & (Pattern.CurrentIsReadOnly = 1) & vbCrLf & _
     "  CurrentValue:" & vbTab & CStr(Pattern.CurrentValue) & vbCrLf
End Function

Private Function GetPatternText_Window(UIAElement As IUIAutomationElement) As String
  Dim Pattern As IUIAutomationWindowPattern
  Set Pattern = UIAElement.GetCurrentPattern(UIAPatterns.Window)
  GetPatternText_Window = vbCrLf & _
     "  CanMaximize: " & vbTab & (Pattern.CurrentCanMaximize = 1) & vbCrLf & _
     "  CanMinimize:" & vbTab & (Pattern.CurrentCanMinimize = 1) & vbCrLf & _
     "  Modal:" & vbTab & (Pattern.CurrentIsModal = 1) & vbCrLf & _
     "  Topmost:" & vbTab & (Pattern.CurrentIsTopmost = 1) & vbCrLf & _
     "  WindowInteractionState:" & vbTab & UIAProps.GetWindowInteractionStateName(Pattern.CurrentWindowInteractionState) & vbCrLf & _
     "  WindowVisualState:" & vbTab & UIAProps.GetWindowVisualStateName(Pattern.CurrentWindowVisualState) & vbCrLf
End Function

Private Sub cbExpand_Click()
  If Not ActiveNode Is Nothing Then
    ExpandOrContractAllChildNodes ActiveNode, True
  End If
  mcTree.Refresh
  SetRootAsActiveNode
End Sub

Private Sub cbCollapse_Click()
  If Not ActiveNode Is Nothing Then
    ExpandOrContractAllChildNodes ActiveNode, False
  End If
  mcTree.Refresh
  SetRootAsActiveNode
End Sub

Private Sub ExpandOrContractAllChildNodes(Node As clsNode, Expand As Boolean)
  Node.Expanded = Expand
  Node.Bold = False
  Dim ChildNode As clsNode
  Dim NumberOfChildren, Counter As Integer
  NumberOfChildren = 0
  On Error Resume Next
  NumberOfChildren = Node.ChildNodes.Count
  On Error GoTo 0
  If NumberOfChildren > 0 Then
    For Counter = 1 To Node.ChildNodes.Count
      Set ChildNode = Node.ChildNodes.Item(Counter)
      ExpandOrContractAllChildNodes ChildNode, Expand
    Next
  End If
End Sub

Private Sub txtSearchText_Change()
  If txtSearchText.Value = "" Then
    cbSearch.Enabled = False
    cbExecutepPath.Enabled = False
  Else
    cbSearch.Enabled = True
    cbExecutepPath.Enabled = True
  End If
End Sub

Private Sub cbSearch_Click()
  Dim SearchText As String
'  SearchText = VBA.Interaction.InputBox("Input the text to search for...", FormName)
  SearchText = txtSearchText.Value
  If SearchText <> "" Then
    Dim Node As clsNode
    Dim MatchingNodes() As clsNode
    Dim Count As Integer
    For Each Node In mcTree.Nodes
    'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/wildcard-characters-used-in-string-comparisons
    'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/like-operator
    If Node.Tag Like SearchText Then
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
       'Ignore
       'MsgBox "No matching nodes found.", vbCritical, FormName
    Else
      Dim i As Integer
      If HighlightSelectedNodeElement Then
        Window.HighlightElements = True
        Window.ReleaseHighlighting
      End If
      For i = 1 To Count
        Dim ExpandableNode As clsNode
        Set ExpandableNode = MatchingNodes(i)
        If HighlightSelectedNodeElement Then
          Dim CurrentFoundUIAElement As IUIAutomationElement
          Set CurrentFoundUIAElement = AllTreeViewNodes(ExpandableNode.Key)("UIAElement")
          Window.HighlightElement CurrentFoundUIAElement, MultiHighlight:=True
        End If
        While Not ExpandableNode Is Nothing
          ExpandableNode.Expanded = True
          Set ExpandableNode = ExpandableNode.ParentNode
        Wend
      Next i
    End If
    mcTree.Refresh
    If Count > 0 Then
      mcTree.ScrollToView MatchingNodes(Count)
    Else
      MsgBox "No matching name found!", vbInformation, FormName
    End If
    'TODO: Activate/Click on found node!?
  End If
  
  SearchExecuted = True
  cbResetSearch.Enabled = True

End Sub

Private Sub cbAddpPathContextNode_Click()
  Dim i As Integer
  If Phosphorus.Utils.GetSizeOfArray(pPathContextNodes()) = -1 Then
    i = 1
    ReDim pPathContextNodes(i)
  Else
    Debug.Print Phosphorus.Utils.GetSizeOfArray(pPathContextNodes())
    i = Phosphorus.Utils.GetSizeOfArray(pPathContextNodes()) + 1
    ReDim Preserve pPathContextNodes(i)
  End If
  Set pPathContextNodes(i) = ActiveNode
  ActiveNode.BackColor = pPathContextNodesBackColour
  cbClearAllpPathContextNodes.Enabled = True
End Sub

Private Sub cbClearAllpPathContextNodes_Click()
  Dim Node As clsNode
  Dim Counter As Integer
  For Counter = 1 To Phosphorus.Utils.GetSizeOfArray(pPathContextNodes())
    Set Node = pPathContextNodes(Counter)
'    If Node.BackColor = pPathContextNodesBackColour Then
      Node.BackColor = VBA.ColorConstants.vbWhite
      Node.ForeColor = VBA.ColorConstants.vbBlack 'The active node text will have been colured white
'    End If
  Next Counter
  Erase pPathContextNodes()
  cbClearAllpPathContextNodes.Enabled = False
End Sub

Private Sub cbExecutepPath_Click()

  On Error GoTo CleanUp
  
  RemoveAllAttributeNodes

  If txtSearchText.Value = "" Then
    MsgBox "No pPath Set!", vbExclamation, FormName
    Exit Sub
  End If
      
  Application.Cursor = xlWait

  If HighlightSelectedNodeElement Then
    Window.HighlightElements = True
    Window.ReleaseHighlighting
  End If

  'Execute the pPath
  Dim pPathString As String
  pPathString = txtSearchText.Value

  Dim pPathLocator As pPath.Core
  Dim pPathResponse As pPath.ReturnClass
  Set pPathLocator = pPath.ConstantsAndStatic.GetNewPhosphorusPPath
  pPathLocator.Initialise
  Dim RootNode As pExternals.clsNode
  Dim RootNodeElement As UIAutomationClient.IUIAutomationElement
  Set RootNode = mcTree.RootNodes(1)
  Set RootNodeElement = AllTreeViewNodes(RootNode.Key)("UIAElement")
  pPathLocator.SetApplicationRootElement RootNodeElement
  
  'Load array of context node elements
  Dim pPathContextNodeElements() As UIAutomationClient.IUIAutomationElement
  Dim Node As clsNode
  Dim i As Integer
  For i = 1 To Phosphorus.Utils.GetSizeOfArray(pPathContextNodes())
    Set Node = pPathContextNodes(i)
    If i = 1 Then
      ReDim pPathContextNodeElements(i)
    Else
      ReDim Preserve pPathContextNodeElements(i)
    End If
    Set pPathContextNodeElements(i) = AllTreeViewNodes(Node.Key)("UIAElement")
  Next i
  i = i - 1
  
  Dim ContextNodeElement As UIAutomationClient.IUIAutomationElement
  Dim InitialPath As String
  If i = 0 Then
    Set pPathResponse = pPathLocator.Evaluate(pPathString)
  Else
    If i = 1 Then
      Dim ContextNode As pExternals.clsNode
      Set ContextNode = pPathContextNodes(i)
      Set ContextNodeElement = pPathContextNodeElements(i)
      InitialPath = AllTreeViewNodes(ContextNode.Key)("Path")
      Set pPathResponse = pPathLocator.Evaluate(pPathString, ContextNodeElement, InitialPath)
    Else
      Dim k As Integer
      For k = 1 To Phosphorus.Utils.GetSizeOfArray(pPathContextNodeElements())
        Set ContextNodeElement = pPathContextNodeElements(k)
        InitialPath = AllTreeViewNodes(Node.Key)("Path")
        pPathLocator.AddContextNode ContextNodeElement, InitialPath
      Next k
      Set pPathResponse = pPathLocator.Evaluate(pPathString)
    End If
  End If
  If pPathResponse.GetErrorMessage <> "" Then
    MsgBox "Error: " & pPathResponse.GetErrorMessage, vbCritical, FormName
    GoTo CleanUp
  End If
  txtSearchReturnValue.Value = pPathResponse.ReturnedValue
  
  'Get all matching elements
  Dim NumberOfMatchingElements As Long
  Dim CurrentMatchingRuntimeId As String

  Dim AttributeName As String
  Dim Caption As String
  Dim AttributeValue As String

  Dim MatchingElementsAndAttributes As Scripting.Dictionary
  Set MatchingElementsAndAttributes = New Scripting.Dictionary
  MatchingElementsAndAttributes.RemoveAll

  NumberOfMatchingElements = pPathResponse.GetFinalNumberOfMatchingElements
  Dim j As Integer
  Dim CurrentMatchingUIAElement As IUIAutomationElement
  For j = 1 To NumberOfMatchingElements
  
    Set CurrentMatchingUIAElement = pPathResponse.GetMatchingElement(j)
    CurrentMatchingRuntimeId = pPath.RuntimeIDs.GetElementRuntimeId(CurrentMatchingUIAElement)
  
    Dim Coll As New Collection
    Coll.Add Item:=CurrentMatchingRuntimeId, Key:="RuntimeId"

    Dim ElementpPath As String
    ElementpPath = pPathResponse.GetMatchingNavigationalPPath(j)
    AttributeName = ""
    AttributeValue = ""
    Dim StartOfAttribute As Integer
    StartOfAttribute = VBA.Conversion.CInt(VBA.Strings.InStr(1, ElementpPath, "/@"))
    If StartOfAttribute > 0 Then
      Caption = VBA.Strings.Mid(ElementpPath, StartOfAttribute + 2)
      Coll.Add Item:=Caption, Key:="Caption"
      Dim Split() As String
      Split = VBA.Strings.Split(Caption, "=")
      AttributeName = Split(0)
      If UBound(Split) > 0 Then
        AttributeValue = VBA.Strings.Replace(Split(1), "'", "")
      End If
    End If

    Coll.Add Item:=ElementpPath, Key:="ElementpPath"
    Coll.Add Item:=AttributeName, Key:="AttributeName" 'Only set for attributes!
    Dim CollKey As String
    CollKey = CurrentMatchingRuntimeId
    If AttributeName <> "" Then
      CollKey = CollKey & " " & AttributeName
    End If
    MatchingElementsAndAttributes.Add CollKey, Coll
    'Prepare for a new collection
    Set Coll = Nothing
    
  Next j
  
  Set CurrentMatchingUIAElement = Nothing

  'Loop through all nodes
  Dim CurrentTestNodeRuntimeId As String
  Dim TestNode As clsNode
  Dim MatchingTestNodes  As New Collection
  For Each TestNode In mcTree.Nodes
  
    'Dont process any new Attribute nodes or Text nodes!
    If (VBA.Strings.InStr(1, TestNode.Key, " Attribute: ") = 0) And (VBA.Strings.InStr(1, TestNode.Key, " Text: ") = 0) Then

      CurrentTestNodeRuntimeId = AllTreeViewNodes(TestNode.Key)("RuntimeId")

      TestNode.Expanded = False
      TestNode.ForeColor = VBA.ColorConstants.vbBlack
      TestNode.Bold = False
      If TestNode Is ActiveNode Then
        TestNode.ForeColor = VBA.ColorConstants.vbBlack
        TestNode.Bold = True
      End If
    
      Dim MatchingRuntimeId As Variant
      Dim MatchFound As Boolean
      MatchFound = False
      Dim Key As Variant
      Dim RuntimeId As String
      For Each Key In MatchingElementsAndAttributes
        
        AttributeName = MatchingElementsAndAttributes(Key)("AttributeName")
        RuntimeId = MatchingElementsAndAttributes(Key)("RuntimeId")
        
        Dim CurrentFoundUIAElement As IUIAutomationElement
        Set CurrentFoundUIAElement = AllTreeViewNodes(TestNode.Key)("UIAElement")
        
        If AttributeName = "" Then
        
          If CurrentTestNodeRuntimeId = RuntimeId Then
            MatchFound = True
            TestNode.ForeColor = VBA.ColorConstants.vbBlue
            
            If VBA.Strings.Right(pPathString, 6) = "text()" Then

              'Add a new text subnode!
              Dim NewTextNode As pExternals.clsNode
              Caption = "Text:=" & UIAPatts.GetTextValue(CurrentFoundUIAElement)
              Set NewTextNode = TestNode.AddChild(AttributeName & " Attribute: " & AttributeName & " RuntimeId: " & CurrentTestNodeRuntimeId, Caption, vImageMain:="Text")
              With NewTextNode
                .ForeColor = VBA.ColorConstants.vbBlue
                .ControlTipText = AttributeName & " Attribute: Value"
                .ControlTipText = "Key: " & TestNode.Key & "@" & Value & "; " & "RuntimeId: " & RuntimeId
                .Tag = "Value"
              End With
              MatchingTestNodes.Add NewTextNode
              Set NewTextNode = Nothing
            
            Else
            
              MatchingTestNodes.Add TestNode
            
            End If
            
            'Never highlight the desktop element
            If HighlightSelectedNodeElement And Not (CurrentFoundUIAElement Is Factory.GetRootDesktopElement) Then
              Window.HighlightElement CurrentFoundUIAElement, MultiHighlight:=True, DelayMs:=0
            End If
          End If
          
        Else
          
          Dim UniqueRuntimeIds As New Scripting.Dictionary
          
          If CurrentTestNodeRuntimeId = RuntimeId Then
          
            MatchFound = True
            'Only highlight the element for the first attribute of each element
            If Not UniqueRuntimeIds.Exists(CurrentTestNodeRuntimeId) Then
              UniqueRuntimeIds.Add CurrentTestNodeRuntimeId, CurrentTestNodeRuntimeId
              'Never highlight the desktop element
              If HighlightSelectedNodeElement And Not (CurrentFoundUIAElement Is Factory.GetRootDesktopElement) Then
                Window.HighlightElement CurrentFoundUIAElement, MultiHighlight:=True, DelayMs:=0
              End If
            End If
            'Add a new subnode!
            Dim NewNode As pExternals.clsNode
            Caption = MatchingElementsAndAttributes(Key)("Caption")
            Set NewNode = TestNode.AddChild(AttributeName & " Attribute: " & AttributeName & " RuntimeId: " & CurrentTestNodeRuntimeId, Caption, vImageMain:="Attribute")
            With NewNode
              .ForeColor = VBA.ColorConstants.vbBlue
              .ControlTipText = AttributeName & " Attribute: " & AttributeName
              .ControlTipText = "Key: " & TestNode.Key & "@" & AttributeName & "; " & "RuntimeId: " & RuntimeId
              .Tag = AttributeName
            End With
            MatchingTestNodes.Add NewNode
            Set NewNode = Nothing
          End If
          Set UniqueRuntimeIds = Nothing
        
        End If

      Next Key
  
    End If
  
  Next TestNode
  Set TestNode = Nothing

  If NumberOfMatchingElements = 0 Then
    MsgBox "No matching elements!", vbExclamation, FormName
    GoTo CleanUp
  End If
    
  Dim Continue As Boolean
  Dim CurrentNode As clsNode
  
  'Expand ancestors of all matching nodes
  Dim MatchingTestNode As Variant
  For Each MatchingTestNode In MatchingTestNodes
    Set CurrentNode = MatchingTestNode
    Continue = True
    While Not CurrentNode.ParentNode Is Nothing And Continue
      If CurrentNode.ParentNode.Expanded Then
        Continue = False
      Else
        CurrentNode.ParentNode.Expanded = True
        Set CurrentNode = CurrentNode.ParentNode
      End If
    Wend
    Set CurrentNode = Nothing
  Next MatchingTestNode
                                
  'Expand ancestors of all context nodes
  Dim c As Integer
  For c = 1 To Phosphorus.Utils.GetSizeOfArray(pPathContextNodes())
    Set CurrentNode = pPathContextNodes(c)
    Continue = True
    While Not CurrentNode.ParentNode Is Nothing And Continue
      If CurrentNode.ParentNode.Expanded Then
        Continue = False
      Else
        CurrentNode.ParentNode.Expanded = True
        Set CurrentNode = CurrentNode.ParentNode
      End If
    Wend
    Set CurrentNode = Nothing
  Next c
  
  mcTree.Refresh
  SearchExecuted = True
  cbResetSearch.Enabled = True

  If MatchingTestNodes.Count <> NumberOfMatchingElements Then
    MsgBox NumberOfMatchingElements & " matching elements were found in the UI but only " & MatchingTestNodes.Count & " were matched to nodes in the pHilby Tree!", vbInformation, FormName
  End If
  
CleanUp:
  
  If Err.Number <> 0 Then
    MsgBox "Error: " & Err.Description, vbInformation, FormName
  End If
  Set pPathLocator = Nothing
  Set pPathResponse = Nothing
  Set MatchingElementsAndAttributes = Nothing
  Application.Cursor = xlDefault
  If HighlightSelectedNodeElement Then
    Application.OnTime Now + TimeValue("00:00:10"), "pHilby.ReleaseHighlighting"
 End If

End Sub

Private Sub cbResetSearch_Click()
  Window.ReleaseHighlighting
  RemoveAllAttributeNodes
  Dim Node As clsNode
  For Each Node In mcTree.Nodes
    With Node
      .ForeColor = VBA.ColorConstants.vbBlack
      If Node Is ActiveNode Then
        .Bold = True
      Else
        .Bold = False
      End If
    End With
  Next Node
  'mcTree.Refresh
  'We need to reload the whole tree reset the underlying TreeView control
  cbReload_Click
  txtSearchText.Value = ""
  txtSearchReturnValue.Value = ""
  SearchExecuted = False
  cbResetSearch.Enabled = False
End Sub

Private Sub RemoveAllAttributeNodes()
  Dim Node As clsNode
  For Each Node In mcTree.Nodes
    If VBA.Strings.InStr(1, Node.Key, " Attribute: ") > 0 Then
      mcTree.NodeRemove Node
    End If
  Next Node
End Sub

Private Sub mcTree_Click(cNode As pExternals.clsNode)
'This gets fired when a node is clicked

  Set ActiveNode = cNode

'Dont process any new Attribute nodes!
If VBA.Strings.InStr(1, ActiveNode.Key, " Attribute: ") = 0 Then

  txtProperties = AllTreeViewNodes(cNode.Key)("AllProperties")
  txtPatterns = AllTreeViewNodes(cNode.Key)("AllPatterns")
  Window.ReleaseHighlighting
  If HighlightSelectedNodeElement Then
    Dim UIAElement As IUIAutomationElement
    Set UIAElement = AllTreeViewNodes(cNode.Key)("UIAElement")
    If IsAlive(UIAElement) Then
      Window.HighlightElement UIAElement
      Application.OnTime Now + TimeValue("00:00:05"), "pHilby.ReleaseHighlighting"
    Else
      MsgBox "This element is not longer available!", vbCritical, FormName
    End If
  End If

  cbExpand.Enabled = True
  cbCollapse.Enabled = True

End If

End Sub

Private Function IsAlive(UIAElement As IUIAutomationElement) As Boolean
  On Error Resume Next
  Dim PID As Long
  PID = UIAElement.CurrentProcessId  'any property access will fail if stale
  IsAlive = (Err.Number = 0) And (PID > 0)
  On Error GoTo 0
End Function

Private Sub cbHighlightSelectedNodeElement_Click()
  With cbHighlightSelectedNodeElement
    If .Tag = "Pushed" Then
      .Picture = WindowsImageAcquisition.LoadImage(ThisWorkbook.Path & "\images\pHilby\marker.png")
      .ControlTipText = "Highlight Element for Selected Node"
      .Tag = "Not Pushed"
      HighlightSelectedNodeElement = False
      Window.ReleaseHighlighting
    Else
      .Picture = WindowsImageAcquisition.LoadImage(ThisWorkbook.Path & "\images\pHilby\highlighter.png")
      .ControlTipText = "Don't Highlight Element for Selected Node"
      .Tag = "Pushed"
      HighlightSelectedNodeElement = True
    End If
    Window.HighlightElements = HighlightSelectedNodeElement
  End With
End Sub

Private Sub cbUIAPolling_Click()
  With cbUIAPolling
    If .Tag = "Pushed" Then
      UnpushcbUIAPolling
      pHilby.StoppHilbyUIAUIAPolling
    Else
      .Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "pointer.png")
      .ControlTipText = "Don't Highlight Node of Element Under Cursor"
      .Tag = "Pushed"
      Application.Wait (Now + TimeValue("0:00:02")) 'Allow time for user to move to the first element
      pHilby.StartpHilbyUIAPolling
    End If
  End With
End Sub

Public Sub UnpushcbUIAPolling()
  With cbUIAPolling
    .Picture = WindowsImageAcquisition.LoadImage(RootImageFolder & "hand.png")
    .ControlTipText = "Highlight Node of Element Under Cursor"
    .Tag = "Not Pushed"
  End With
End Sub

Public Sub SearchByCursorPoint(tPt As tagPOINT)
  Dim CurrentNode As clsNode
  Dim SmallestMatchingNode As clsNode
  Dim SmallestMatchingRect As BoundingRectangle
  Dim CurrentRect As BoundingRectangle
  Dim Inside As Boolean
  
  Set SmallestMatchingNode = Nothing
  
  Dim Count As Integer
  For Each CurrentNode In mcTree.Nodes

    CurrentNode.Bold = False
    CurrentNode.Expanded = False
    Dim CurrentElement As IUIAutomationElement
    Set CurrentElement = AllTreeViewNodes(CurrentNode.Key)("UIAElement")
    Set CurrentRect = Factory.GetNewBoundingRectangle(CurrentElement)
    Inside = tPt.X >= CurrentRect.Left And tPt.X <= CurrentRect.Right And tPt.Y >= CurrentRect.Top And tPt.Y <= CurrentRect.Bottom
    If Inside Then
      If SmallestMatchingNode Is Nothing Then
        Set SmallestMatchingNode = CurrentNode
        Set SmallestMatchingRect = CurrentRect
      Else
        If CurrentRect.Left >= SmallestMatchingRect.Left And _
           CurrentRect.Left <= SmallestMatchingRect.Right And _
           CurrentRect.Top >= SmallestMatchingRect.Top And _
           CurrentRect.Bottom <= SmallestMatchingRect.Bottom Then
          Set SmallestMatchingNode = CurrentNode
          Set SmallestMatchingRect = CurrentRect
        End If
      End If
    End If
  Next CurrentNode
   
  If SmallestMatchingNode Is Nothing Then
      MsgBox "No matching nodes found.", vbCritical, FormName
  Else
    Dim i As Integer
    Dim ExpandableNode As clsNode
    Set ExpandableNode = SmallestMatchingNode
    While Not ExpandableNode Is Nothing
      ExpandableNode.Expanded = True
      Set ExpandableNode = ExpandableNode.ParentNode
    Wend
    mcTree.Refresh
    mcTree.ScrollToView SmallestMatchingNode
    SmallestMatchingNode.Bold = True
  End If

End Sub

Private Sub cbExportToNode_Click()
  cbExportToExcel ActiveNode
End Sub

Private Sub cbExport_Click()
  cbExportToExcel
End Sub

Private Sub cbExportToExcel(Optional TargetNode As pExternals.clsNode)

  Application.ScreenUpdating = False
  Application.Cursor = xlWait
    
  Dim wb As New Excel.Workbook
  Set wb = Workbooks.Add
  
  Dim ws As Worksheet
  Dim wsIndex As Worksheet
  For Each ws In wb.Sheets
    If ws.Name = "Sheet1" Then
      ws.Name = "Index"
      Set wsIndex = ws
    Else
      ws.Delete
    End If
  Next ws

'Stop
  wsIndex.Range("a1").Value = "Link"
  wsIndex.Range("B1").Value = "Path"
  wsIndex.Range("C1").Value = "Name"
  With wsIndex.Range("A1:C1").Font
    .Bold = True
    .Underline = xlUnderlineStyleSingle
  End With
  wsIndex.Range("B2").Select
  ActiveWindow.FreezePanes = True
  
  Dim SheetCounter As Integer
  Range("A1").Select
  Dim Key As Variant
  For Each Key In AllTreeViewNodes
    
    Dim ExportCurrentNode As Boolean
    ExportCurrentNode = True
    
    If Not TargetNode Is Nothing Then
      If Len(Key) > Len(TargetNode.Key) Then
        ExportCurrentNode = False
      Else
        If Len(Key) = Len(TargetNode.Key) Then
          If Key = TargetNode.Key Then
            ExportCurrentNode = True
          Else
            ExportCurrentNode = False
          End If
        Else
          If VBA.Strings.Left(TargetNode.Key, Len(Key) + 1) = (Key & ".") Then
            ExportCurrentNode = True
          Else
            ExportCurrentNode = False
          End If
        End If
      End If
    End If
    
    If ExportCurrentNode Then

      SheetCounter = SheetCounter + 1
      wsIndex.Cells(SheetCounter + 1, 2).Value = "'" & Key
      wsIndex.Hyperlinks.Add _
        Anchor:=wsIndex.Cells(SheetCounter + 1, 1), _
        Address:="", _
        SubAddress:="Node" & SheetCounter & "!A1", _
        TextToDisplay:="Node" & SheetCounter & "!A1"

      wsIndex.Cells(SheetCounter + 1, 2).Value = AllTreeViewNodes(Key)("Path")
      wsIndex.Cells(SheetCounter + 1, 3).Value = AllTreeViewNodes(Key)("Name")

      Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
      ws.Name = "Node" & SheetCounter
      ws.Hyperlinks.Add _
        Anchor:=ws.Range("A1"), _
        Address:="", _
        SubAddress:=wsIndex.Name & "!A" & (SheetCounter + 1), _
        TextToDisplay:="Back"

      Dim i As Integer
      Dim Properties() As String
      Properties = Split(AllTreeViewNodes(Key)("AllProperties"), vbCrLf)

      Dim PropertyValues() As String
      ReDim PropertyValues(UBound(Properties))
      For i = 0 To UBound(Properties)
        Dim Property() As String
        If InStr(Properties(i), vbTab) > 0 Then
          Property = Split(Properties(i), vbTab)
          Properties(i) = Property(0)
          PropertyValues(i) = Property(1)
        End If
      Next

      ws.Range("A2").Resize(UBound(Properties) + 1) = Application.Transpose(Properties)
      ws.Range("B2").Resize(UBound(Properties) + 1) = Application.Transpose(PropertyValues)

      With ws.Range("A3")
        .Value = "Properties"
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
      End With

      Dim AllPatterns As String
      AllPatterns = AllTreeViewNodes(Key)("AllPatterns")
      AllPatterns = Replace(AllPatterns, "=", "")
      AllPatterns = Replace(AllPatterns, vbCrLf & vbCrLf, vbCrLf)
      Dim Patterns() As String
      Patterns = Split(AllPatterns, vbCrLf)
      Dim PatternValues() As String
      ReDim PatternValues(UBound(Patterns))
      For i = 0 To UBound(Patterns)
        Dim Pattern() As String
        If InStr(Patterns(i), vbTab) > 0 Then
          Pattern = Split(Patterns(i), vbTab)
          Patterns(i) = Pattern(0)
          PatternValues(i) = Pattern(1)
        End If
      Next

      Dim StartOfPatterns As Range
      Set StartOfPatterns = ws.Range("A" & UBound(Properties) + 4)
      With StartOfPatterns
        .Value = "Patterns"
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
      End With
      StartOfPatterns.Offset(1, 0).Resize(UBound(Patterns) + 1) = Application.Transpose(Patterns)
      StartOfPatterns.Offset(1, 1).Resize(UBound(Patterns) + 1) = Application.Transpose(PatternValues)

      ws.Range("A1").EntireColumn.AutoFit
      ws.Range("4:4").Select
      ActiveWindow.FreezePanes = True
      ws.Range("B4").Activate
    
    End If
    
  Next
  
  With wsIndex
    .Activate
    .Cells.EntireColumn.AutoFit
  End With

  Application.ScreenUpdating = True
  Application.Cursor = xlDefault
  
  MsgBox "Export finished!", vbInformation, FormName
  
End Sub

Private Sub cbReload_Click()
  cbClearAllpPathContextNodes_Click
  If Not mRootUIAElement Is Nothing Then
    Application.Cursor = xlWait
    LoadTreeView mRootUIAElement, MaxNumberOfLevels:=mMaxNumberOfLevels, DelayLoadingInSeconds:=mDelayLoadingInSeconds
    txtProperties = ""
    txtPatterns = ""
    Application.Cursor = xlDefault
  End If
End Sub

Private Sub cbSetCurrentAsRootNode_Click()
  If Not ActiveNode Is Nothing Then
    Set mRootUIAElement = AllTreeViewNodes(ActiveNode.Key)("UIAElement")
    cbReload_Click
  End If
End Sub

Private Sub cbSetParentAsRootNode_Click()
  If Not ActiveNode Is Nothing Then
    Dim CurrentRootUIAElement As IUIAutomationElement
    Set CurrentRootUIAElement = AllTreeViewNodes(ActiveNode.Key)("UIAElement")
    If Not CurrentRootUIAElement Is Factory.GetRootDesktopElement Then
      Set TreeWalker = UIA.ControlViewWalker
      Set mRootUIAElement = TreeWalker.GetParentElement(CurrentRootUIAElement)
      cbReload_Click
    End If
  End If
End Sub

Private Sub LevelsSpinButton_SpinUp()
  mMaxNumberOfLevels = mMaxNumberOfLevels + 1
  txtMaxNumberOfLevels.Value = mMaxNumberOfLevels
End Sub

Private Sub LevelsSpinButton_SpinDown()
  If mMaxNumberOfLevels > 0 Then
    mMaxNumberOfLevels = mMaxNumberOfLevels - 1
  End If
  txtMaxNumberOfLevels.Value = mMaxNumberOfLevels
End Sub

Private Sub DelaySpinButton_SpinUp()
  mDelayLoadingInSeconds = mDelayLoadingInSeconds + 1
  txtDelayLoading.Value = mDelayLoadingInSeconds
End Sub

Private Sub DelaySpinButton_SpinDown()
  If mDelayLoadingInSeconds > 0 Then
    mDelayLoadingInSeconds = mDelayLoadingInSeconds - 1
  End If
  txtDelayLoading = mDelayLoadingInSeconds
End Sub

Private Sub lblFlatIconLink_Click()
  ThisWorkbook.FollowHyperlink "https://www.flaticon.com/"
End Sub

Private Sub lblTreeViewLink_Click()
  ThisWorkbook.FollowHyperlink "https://jkp-ads.com/articles/treeview.aspx"
End Sub

