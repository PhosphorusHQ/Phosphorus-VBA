VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmpHilby 
   Caption         =   "pHilby - Phosphorus UIAutomation Spy Tool"
   ClientHeight    =   6552
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11616
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

Private WithEvents mcTree As pExternals.clsTreeView
Attribute mcTree.VB_VarHelpID = -1
Public AppName As String

'Private mbExit As Boolean    ' to exit a SpinButton event


'''' for stress testing this demo
Private mlCntChildren As Long
'Private mlDemoNo As Long

Const mcPtPixel As Single = 0.75

Private iMaxNumberOfLevels As Integer

Private Sub UserForm_Initialize()
'See the Compile constant DebugMode in tools, VBAProject properties
'DebugMode=1 will enable the #If to Stop in Error handlers
'Normally this should be removed before distributing
    
    ' Hide the Image container
'    Me.frmImageBox.Visible = False
'    Me.frmImageBox.Enabled = False
'    Me.Width = 440
'  mbExit = True
    
'    Me.labIndent.Tag = Me.labIndent.Caption
'    Me.labNodeHeight.Tag = Me.labNodeHeight.Caption
'    Me.labFont.Tag = Me.labFont.Caption
'    Me.txtChildren.Value = 15: mlCntChildren = 15
'    Me.SpinButton1.Value = 20
'    Me.SpinButton2.Value = 16
  If Me.fraTreeControl.Font.Size < 4 Then
    Me.fraTreeControl.Font.Size = 4
  End If
'    Me.SpinButton3.Value = Me.fraTreeControl.Font.Size
'    mbExit = False
'    lblBuild.Caption = "Build Number: " & GCSBUILD & " "
    
  #If DEBUGMODE = 1 Then
    gFormInit = gFormInit + 1
  #End If

'  #If Mac Then
'        Dim objCtl As MSForms.Control
'
'        With Me
'            .Font.Size = 10
'            .Width = .Width * 4 / 3
'            .Height = .Height * 4 / 3
'            .BackColor = .labInfo.BackColor
'        End With
'
'        For Each objCtl In Me.Controls
'            With objCtl
'                .Left = Int(.Left * 4 / 3)
'                .Top = Int(.Top * 4 / 3)
'                .Width = Int(.Width * 4 / 3)
'                .Height = Int(.Height * 4 / 3)
'                Select Case TypeName(objCtl)
'                Case "Image", "SpinButton"
'                Case "TextBox", "Frame"
'                    .Font.Size = 10
'                Case Else
'                    .Font.Size = 10
'                End Select
'            End With
'        Next
'    #End If
    
End Sub

Private Sub UserForm_Terminate()
  #If DEBUGMODE = 1 Then
    gFormTerm = gFormTerm + 1
    ClassCounts
  #End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Make sure all objects are destroyed
  If Not mcTree Is Nothing Then
    mcTree.TerminateTree
  End If
  Set mcTree = Nothing
End Sub

'########### Treeview container frame events ################

' Enter/Exit events are not trapped with 'WithEvents' in the treeclass
' so if needed they can be trapped in the form. (These will toggle the active node's highlight states)

Private Sub frTreeControl_Enter()
Stop
    If Not mcTree Is Nothing Then
        mcTree.EnterExit False
    End If
End Sub

Private Sub frTreeControl_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Stop
    If Not mcTree Is Nothing Then
        mcTree.EnterExit True
    End If
End Sub

'/########### Treeview container frame events ################


'===========================
Public Sub LoadTreeView(RootUIAElement As IUIAutomationElement, Optional MaxNumberOfLevels As Integer)

'    cmdStop_Click
  UnloadTreeView
'    cmdTestSort.Visible = Me.cbxTestSort.Value
'    Me.cbxCheckBoxes.Caption = "ChkBox TriState"
  iMaxNumberOfLevels = MaxNumberOfLevels
  LoadUIATree RootUIAElement
'    Me.SpinButton3.Enabled = False
'    mlDemoNo = 1
'    Me.SpinButton3.Enabled = False
'  If Not mcTree Is Nothing Then
'    mbExit = True
'        SpinButton1 = mcTree.Indentation / mcPtPixel
'        SpinButton2 = mcTree.NodeHeight / mcPtPixel
'        UpdateInfoLabel
'    mbExit = False
'????
'       ' Me.fraTreeControl.SetFocus
'  End If
End Sub

Private Sub UnloadTreeView()
  If Not mcTree Is Nothing Then
    mcTree.NodesClear
    Set mcTree = Nothing
  End If
'  mlDemoNo = 0
'    Me.labInfo = ""
'    Me.Caption = AppName
'    cmdTestSort.Visible = False
'    Me.SpinButton3.Enabled = True
'    Me.cbxCheckBoxes.Caption = "Checkboxes"
End Sub

Private Sub LoadUIATree(RootUIAElement As IUIAutomationElement)

  Dim cRoot As clsNode
  Dim cNode As clsNode
  Dim cExtraNode As clsNode
  Dim i As Long
  Dim k As Long

  Set mcTree = pExternals.Factory.GetTreeView
    
  On Error GoTo ErrorHandler

  With mcTree
        
    'The VBA treeview needs a container Frame control on the userform.
    'Just draw a frame on your form where the treeview should appear
    'Make sure the frame has no border and no caption
    'Set it to show both scroll bars. (Keepscrollbarsvisible to both)
    'Set it to the special effect "Sunken"
    'Note that node label formats such as back/fore colors and font adopt
    'the frame container's properties, so change if/as required.

    'Pass the frame to the TreeControl of the treeview class
    Set .TreeControl = Me.fraTreeControl

    Call .NodesClear

    .AppName = Me.AppName

        'Set some characteristics of the root of the tree,
        'which for this demo we pick up from checkbox and spinner controls on the form
'        .CheckBoxes(bTriState:=cbxCheckBoxes) = cbxCheckBoxes
'        .RootButton = cbxRootButton
        '.LabelEdit(bAutoSort:=True) = IIf(cbxAllowEditing.Value, 0, 1)  'default 0, can be edited (like LabelEditConstants tvwAutomatic/tvwManual)
        ' new in v022, EnableLabelEdit added as alternative to LabelEdit
        
'        .EnableLabelEdit(bAutoSort:=True, bMultiLine:=True) = cbxAllowEditing.Value
'??? Not needed?
'.EnableLabelEdit(bAutoSort:=False, bMultiLine:=False) = False
.EnableLabelEdit() = False

'        .FullWidth = cbxFullWidth.Value
'        .Indentation = SpinButton1.Value * mcPtPixel
'        .NodeHeight = SpinButton2.Value * mcPtPixel
'        .ShowLines = Me.cbxShowlines.Value
'????
.ShowLines = True
'        .ShowExpanders = Me.cbxShowExpanders
'????
.ShowExpanders = True

        '.LineColor = RGB(100, 100, 200)
'        .NarratorReaderControl = Me.cbxNarrator     '(default=true, ensures Narrator will read out the active node's caption)
.NarratorReaderControl = True

'        If cbxShowExpanders And cbxExpanderIcons Then
'            ' Win7 style arrow icons, try "Win7Plus1" & "Win7Plus2" for preference
'            Call .ExpanderImage(Me.frmImageBox.Controls("Win7Minus").Picture, _
'                                Me.frmImageBox.Controls("Win7Plus2").Picture)
'        End If
'        If cbxCheckBoxes And cbxCheckboxIcons Then
'            Call .CheckboxImage(Me.frmImageBox.Controls("CheckboxFalse").Picture, _
'                                Me.frmImageBox.Controls("CheckboxTrue").Picture, _
'                                Me.frmImageBox.Controls("CheckboxNull").Picture)
'        End If

'iLevelCounter = 1
LoadRootElementAndAllDescendants RootUIAElement
'create the node controls and display the tree
.Refresh
'Me.fraTreeControl.SetFocus
Exit Sub

''''''''''    For k = 1 To 1    ' # roots
''''''''''
''''''''''      ' add a Root node and make it bold
''''''''''      Set cRoot = .AddRoot(sKey:="Root" & k, vCaption:="Root Node" & k)
''''''''''      cRoot.Bold = True
''''''''''
''''''''''      cRoot.ControlTipText = "Tip for Root Node" & k & ". Context tips can also be added to all nodes"
''''''''''
''''''''''      'Add branches with child nodes to the root:
''''''''''      'Keys are optional but if using them they must be unique,
''''''''''      'attempting to add a node with a duplicate key will cause a runtime error.
''''''''''      '(below we will include unique keys with all the nodes)
''''''''''
''''''''''      Set cNode = cRoot.AddChild(cRoot.Key & "_1", "1.A")
''''''''''
''''''''''      'Add a 2nd branch to the root:
''''''''''      Set cNode = cRoot.AddChild(cRoot.Key & "_2", "2.B")
''''''''''
''''''''''      'If you want to add more child branches to a branch later on, use a variable to store the branch.
''''''''''      Set cExtraNode = cNode.AddChild(cRoot.Key & "_2.1", "2.1  level 2")
''''''''''      cExtraNode.Expanded = False    ' this node will initially be collapsed, it's child node controls will not initially be created
''''''''''
''''''''''      'To add a branches to a branch, make sure you set a variable to its 'main' or parent branch when created
''''''''''      'Then use the Branch's AddChild method:
''''''''''
''''''''''      Set cNode = cNode.AddChild(cRoot.Key & "_2.2", "2.2  level 2")
''''''''''
''''''''''      ' with TriState this node's parents and automatically checked, each triggering a NodeCheck event
''''''''''      ' the following three child nodes will also be automatically checked but as not yet added will not trigger NodeCheck events
''''''''''      If .CheckBoxes Then cNode.Checked = True
''''''''''
''''''''''        Set cNode = cNode.AddChild(cRoot.Key & "_2.2.1", "2.2.1  level 3")
''''''''''        Set cNode = cNode.AddChild(cRoot.Key & "_2.2.1.1", "2.2.1.1  level 4")
''''''''''        Set cNode = cNode.AddChild(cRoot.Key & "_A", "2.2.1.1.1   level 5 with extra text to test scrollwidth")
''''''''''
''''''''''        'Now add another branch to the branch we stored earlier
''''''''''        cExtraNode.AddChild cRoot.Key & "_2.1.1", "2.1.1  level 3"
''''''''''
''''''''''        'Add a 3rd branch to the root, with a child node
''''''''''        Set cNode = cRoot.AddChild(cRoot.Key & "_3", "3.C")
''''''''''        cNode.AddChild cRoot.Key & "_3.1", "3.1  level 2"
''''''''''
''''''''''        'Add a 4th branch to the root
''''''''''        Set cNode = cRoot.AddChild(cRoot.Key & "_4", "4.D")
''''''''''        cNode.Caption = "4.D  +" & mlCntChildren & ""
''''''''''
''''''''''        ' add a bunch of child nodes to the root's 4th node
''''''''''        For i = 1 To mlCntChildren  ' 15
''''''''''           cNode.AddChild cRoot.Key & "_4." & i, "4." & Right$("000" & i, 4)
''''''''''        Next
''''''''''
'''''''''''            If cbxTestSort.Value = True Then
'''''''''''                SortTest cNode, False
'''''''''''            End If
''''''''''
''''''''''    Next
''''''''''
''''''''''    'create the node controls and display the tree
''''''''''    .Refresh

  End With

  Exit Sub

ErrorHandler:
    #If DEBUGMODE = 1 Then
        Debug.Print Err.Source, Err.Description
        Stop
        Resume
    #End If

    Debug.Print Err.Source, Err.Description

    If Not mcTree Is Nothing Then
        mcTree.NodesClear
    End If

End Sub

Private Sub LoadRootElementAndAllDescendants(UIAElement As IUIAutomationElement)

  Dim RuntimeId As String
  Dim Caption As String
  RuntimeId = UIACommon.GetElementRuntimeId(UIAElement)
'Make common
  Caption = UIAElement.CurrentName & " (" & UIAProps.ControlTypeName(UIAElement.CurrentControlType) & ")"

  ' Add the root node and make it bold
  Dim Root As clsNode
  Set Root = mcTree.AddRoot(sKey:=RuntimeId, vCaption:=Caption)
  Root.Bold = True
  Root.ControlTipText = "RuntimeID:=" & RuntimeId
  Root.Tag = UIAElement.CurrentName
  Root.Expanded = True

  LoadAllRootElementDescendants UIAElement, 1, Root

End Sub

Private Sub LoadAllRootElementDescendants(UIAElement As IUIAutomationElement, LevelCounter As Integer, ParentNode As clsNode)

  'If iMaxNumberOfLevels=0 show all levels
  If iMaxNumberOfLevels = 0 Or (iMaxNumberOfLevels >= LevelCounter) Then
        
    Dim AllElements As IUIAutomationElementArray
    Set AllElements = UIAElement.FindAll(TreeScope.Children, UIA.CreateTrueCondition)
    
    If AllElements.Length = 0 Then
      ' No children, just exit
      Exit Sub
    End If
    
    Dim i As Long
    Dim RuntimeId As String
    Dim Caption As String
    Dim CurrentUIAElement As IUIAutomationElement
    Dim ChildNode As clsNode
    For i = 0 To AllElements.Length - 1
      Set CurrentUIAElement = AllElements.GetElement(i)
      RuntimeId = UIACommon.GetElementRuntimeId(CurrentUIAElement)
'Make common
      Caption = CurrentUIAElement.CurrentName & " (" & UIAProps.ControlTypeName(UIAElement.CurrentControlType) & ")"
      Set ChildNode = ParentNode.AddChild(sKey:=RuntimeId, vCaption:=Caption)
      ChildNode.ControlTipText = "RuntimeID:=" & RuntimeId
      ChildNode.Tag = CurrentUIAElement.CurrentName
      ChildNode.Expanded = False
      LoadAllRootElementDescendants CurrentUIAElement, LevelCounter + 1, ChildNode
    Next i
    
  End If

End Sub
'    Debug.Print VBA.Strings.String(1 * 2, " ") & "Element #" & i
'    Debug.Print VBA.Strings.String(1 * 2, " ") & "AriaRole: " & CurrentElement.GetCurrentPropertyValue(UIAProperties.AriaRole)
'    Debug.Print VBA.Strings.String(1 * 2, " ") & "ControlType: " & CurrentElement.GetCurrentPropertyValue(UIAProperties.ControlType) & " (" & ControlTypeName(CurrentElement.GetCurrentPropertyValue(UIAProperties.ControlType)) & ")"
'    Debug.Print VBA.Strings.String(1 * 2, " ") & "ClassName: " & CurrentElement.GetCurrentPropertyValue(UIAProperties.ClassName)
'    Debug.Print VBA.Strings.String(1 * 2, " ") & "Name: " & CurrentElement.GetCurrentPropertyValue(UIAProperties.Name)

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

'This gets fired after a node has been edited
'Private Sub mcTree_AfterLabelEdit(Cancel As Boolean, NewString As String, cNode As clsNode)
'
'' Validate user's manually edited node here
'
'    If Len(NewString) >= 3 Then
'        If NewString Like ("*123*") Then
'            MsgBox "123   is not allowed !" & vbCr & "but try 456", , AppName
'            Cancel = True    ' undo user's change
'        ElseIf NewString Like "*456*" Then
'            NewString = Replace(NewString, "456", "789")
'        End If
'    End If
'
'End Sub

'This gets fired when a node is clicked
Private Sub mcTree_Click(cNode As pExternals.clsNode)
    Dim s As String
    With cNode
'PJG Index Method not recognised???
'        s = .Caption & IIf(mcTree.CheckBoxes, "   Checked:" & .Checked, "") & vbNewLine & _
'            "Key:  " & .Key & vbNewLine & _
'            "Index:  " & .Index & "    VisIndex:  " & .VisIndex & "    Level:  " & .Level
s = "???"
    End With
'Add A text box to display the properties of the current node
'    Me.labInfo.Caption = s
End Sub

'This gets fired when a key is pressed down
Private Sub mcTree_KeyDown(cNode As pExternals.clsNode, ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
' PT demo
    Dim bMove As Boolean
    Dim sMsg As String
    Dim cSource As clsNode
    
    Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, _
         48 To 57, 96 To 105, vbKeyF2, 20, 93, _
         vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd
        ' these keys are already trapped in clsTreeView for navigation, expand/collapse, edit mode

    Case vbKeyC, vbKeyC + 48
        If Shift = 2 Then    ' Ctrl-X move
            ' code here to validate if user can copy this node
            Set mcTree.MoveCopyNode(False) = mcTree.ActiveNode
        End If

    Case vbKeyX, vbKeyX + 48
        If Shift = 2 Then    ' Ctrl-X move
            ' code here to validate if user can move this node
            Set mcTree.MoveCopyNode(True) = mcTree.ActiveNode
        End If

    Case vbKeyV, vbKeyV + 48
        If Shift = 2 Then    ' Ctrl-V paste
            Set cSource = mcTree.MoveCopyNode(bMove)
            If Not cSource Is Nothing Then
                ' code to validate if the stored 'MoveCopyNode' can be Moved or Copied to the selected node
                If bMove Then
                    mcTree.Move cSource, mcTree.ActiveNode, bShowError:=True
                Else
                    mcTree.Copy cSource, mcTree.ActiveNode, bShowError:=True
                End If

                mcTree.ActiveNode.Sort    ' assume user wants move/copy to locate as sorted
                mcTree.ActiveNode.Expanded = True    ' assume user wants to see the moved/copied node if behind a collapsed node
                mcTree.Refresh
            End If
        End If
    Case vbKeyDelete
        sMsg = "Are you sure you want to delete node ''" & cNode.Caption & "'' and all it's child-nodes?" & vbCr & _
                        vbCr & "(press Ctrl-break now and click Debug to see this event code)"
        If MsgBox(sMsg, vbOKCancel, AppName) = vbOK Then
            mcTree.NodeRemove cNode
            mcTree.Refresh
        End If
    End Select
End Sub

'This gets fired when a node is checked
Private Sub mcTree_NodeCheck(cNode As clsNode)
'???
'    Dim s As String
'    With cNode
'        s = .Caption & "     Check changed: " & .Checked
'    End With

'    Me.labInfo.Caption = s

End Sub

'/########## Treeview Events, raised in clsTreeView ##########


'########## Set some Treeview properties with controls on the demo form ##########

'Private Sub txtchildren_AfterUpdate()
'    On Error Resume Next
'    mlCntChildren = CLng(txtChildren.Value)
'    If mlCntChildren > 5000 Then
'        MsgBox mlCntChildren & " is a lot, press Ctrl-Break to abort loading the demo", , AppName
'    End If
'    txtChildren = mlCntChildren
'End Sub

'''''' checkboxes ''''''''
'Private Sub cbxFullWidth_Click()
'' PT True: node labels strech full width, will use separate image controls for icons
''    False: node labels sized to text, but no extra time to show icons with a big tree
''    If value is changed during runtime must clear rebuild the tree
'
'    Dim lDemoNo As Long
'
'    If Not mcTree Is Nothing Then
'        lDemoNo = mlDemoNo
'        cmdStop_Click
'        If lDemoNo = 2 Then
'            cmdDemo2_Click
'        ElseIf lDemoNo = 1 Then
'            cmdDemo1_Click
'        End If
'    End If
'
'End Sub

'Private Sub cbxNarrator_Click()
'    If Not mcTree Is Nothing Then
'        mcTree.NarratorReaderControl = cbxNarrator.Value
'        mcTree.Refresh
'    End If
'End Sub

'Private Sub cbxShowlines_Click()
'    If Not mcTree Is Nothing Then
'        mcTree.ShowLines = cbxShowlines
'    End If
'End Sub

'Private Sub cbxRootButton_Click()
'' PT equivalent to the original Treeview's LineStyle property,
'' TreeView.LineStyle = tvwTreeLines or tvwRootLines
''
'    If Not mcTree Is Nothing Then
'        mcTree.RootButton = cbxRootButton
'        UpdateInfoLabel
'    End If
'End Sub

'Private Sub cbxCheckBoxes_Click()
'    If mbExit Then Exit Sub
'    On Error GoTo locErr
'    If cbxCheckboxIcons Then
'        cbxCheckboxIcons = False
'    ElseIf Not mcTree Is Nothing Then
'        If mlDemoNo = 1 Then
'            ' demo checkboxes with triState
'            mcTree.CheckBoxes(bTriState:=True) = cbxCheckBoxes.Value
'        Else
'            mcTree.CheckBoxes = cbxCheckBoxes.Value
'        End If
'    End If
'    Exit Sub
'locErr:
'    MsgBox Err.Description, , Err.Source
'End Sub

'Private Sub cbxCheckboxIcons_Click()
''PT
'    If cbxCheckboxIcons And Not cbxCheckBoxes Then
'        mbExit = True
'        cbxCheckBoxes = True
'        mbExit = False
'    End If
'
'    If Not mcTree Is Nothing Then
'        If mlDemoNo = 2 Then
'            cmdDemo2_Click
'        ElseIf mlDemoNo = 1 Then
'            cmdDemo1_Click
'        End If
'    End If
'
'End Sub

'Private Sub cbxShowExpanders_Click()
'    If mbExit Then Exit Sub
'
'    If cbxExpanderIcons Then
'        cbxExpanderIcons = False
'    ElseIf Not mcTree Is Nothing Then
'        mcTree.ShowExpanders = cbxShowExpanders
'    End If
'End Sub

'Private Sub cbxExpanderIcons_Click()
''PT
'    If cbxExpanderIcons And Not cbxShowExpanders Then
'        mbExit = True
'        cbxShowExpanders = True
'        mbExit = False
'    End If
'
'    If Not mcTree Is Nothing Then
'        If mlDemoNo = 2 Then
'            cmdDemo2_Click
'        ElseIf mlDemoNo = 1 Then
'            cmdDemo1_Click
'        End If
'    End If
'
'End Sub

'Private Sub cbxAllowEditing_Click()
'' PT
'    If Not mcTree Is Nothing Then
'        mcTree.EnableLabelEdit(bAutoSort:=(mlDemoNo = 1), bMultiLine:=(mlDemoNo = 1)) = cbxAllowEditing.Value
'    End If
'End Sub

'Private Sub cbxTestSort_Click()
''PT
'    If mlDemoNo Then
'        cmdTestSort.Visible = cbxTestSort
'    End If
'End Sub

'Private Sub cmdTestSort_Click()
'' PT
'Dim cNode As clsNode
'    On Error GoTo errH
'    If Not mcTree Is Nothing And mbExit = False Then
'
'        Set cNode = mcTree.ActiveNode
'        If cNode.ChildNodes Is Nothing Then
'            MsgBox "The selected node does not have any child-nodes", , AppName
'        ElseIf cNode.ChildNodes.Count = 1 Then
'            MsgBox "The selected node only has one child-node", , AppName
'        Else
'            SortTest cNode, True
'        End If
'
'    End If
'    Exit Sub
'errH:
'    MsgBox Err.Description, , AppName
'End Sub

'Private Sub SortTest(cNode As clsNode, bRefresh As Boolean)
''PT
'    Dim i As Long, j As Long
'    Dim sText As String
'    Dim cChild As clsNode
'
'    If Not cNode.ChildNodes Is Nothing Then
'        For Each cChild In cNode.ChildNodes
'            i = i + 1
'            sText = ""
'            For j = 1 To 10
'                sText = sText & Chr(Int(65 + Rnd * 26)) & " "
'            Next
'            cChild.Caption = sText & "   " & vbTab & Right$("000" & i, 4)
'        Next
'
'        cNode.Sort ndAscending, ndTextCompare
'
'        If bRefresh Then
'            mcTree.Refresh
'        End If
'    End If
'End Sub

'''''''' Spinners''''''''''''''
'Private Sub SpinButton1_Change()
'' PT change Indentation during at run-time
''   equivalent to the original Treeview's Indentation
'    Dim sngIndent As Single
'
'    sngIndent = SpinButton1.Value * mcPtPixel
'    If Not mbExit Then
'        If Not mcTree Is Nothing Then
'            sngIndent = SpinButton1.Value * mcPtPixel
'            mcTree.Indentation = sngIndent
'            If mcTree.Indentation <> sngIndent Then    ' tried to change beyond limits
'                SpinButton1.Value = sngIndent / mcPtPixel
'            End If
'            UpdateInfoLabel
'        End If
'
'    End If
'    labIndent = labIndent.Tag & ": " & sngIndent
'End Sub

'Private Sub SpinButton2_Change()
'' PT change NodeHeight during run-time
''    the original Treeview's NodeHeight is governed by the Font
''    in ours the NodeHeight can be set but will "auto-increase" for Font as needs
'
'    Dim sngNodeHt As Single
'
'    sngNodeHt = SpinButton2.Value * mcPtPixel
'    If Not mbExit Then
'        If Not mcTree Is Nothing Then
'            mcTree.NodeHeight = sngNodeHt
'            If mcTree.NodeHeight <> sngNodeHt Then    ' tried to change beyond limits
'                sngNodeHt = mcTree.NodeHeight
'                mbExit = True
'                SpinButton2.Value = sngNodeHt / mcPtPixel
'                mbExit = False
'            End If
'            UpdateInfoLabel
'        End If
'    End If
'    labNodeHeight = labNodeHeight.Tag & ": " & sngNodeHt
'
'End Sub

'Private Sub SpinButton3_Change()
'' PT When the node labels are created font properties are inherited from
''    the parent container, ie the Frame (frTreeControl)
''    Note the minimum NodeHeight will autosize to the font size, but if decreasing the
''    font size might also want to decrease the nodeHeight
''    This control is disabled at while a treeview demo is displayed
'
'    If Not mbExit Then
'        Me.frTreeControl.Font.Size = SpinButton3.Value
'    End If
'    Me.labFont.Caption = Me.labFont.Tag & ":   " & SpinButton3.Value
'End Sub

'/########## Set some Treeview properties with controls on the demo form ##########


'########### Some Read/Write examples behind the buttons #########################

Private Sub cmdFullPath_Click()
' PT similar to the original Treeview's FullPath function
Dim cNode As clsNode
    Dim s As String
    If Not mcTree Is Nothing Then
        Set cNode = mcTree.ActiveNode
        If cNode Is Nothing Then Exit Sub
        s = mcTree.ActiveNode.Caption
        s = s & vbCr & mcTree.ActiveNode.FullPath
        MsgBox s, , AppName
    End If

End Sub

Private Sub cmdReset_Click()
' PT change formats and clear Checked
    Dim cNode As clsNode
    If Not mcTree Is Nothing Then
        For Each cNode In mcTree.Nodes
            With cNode
                .BackColor = vbWindowBackground
                .ForeColor = vbWindowText
                .Bold = False
                .Checked = False
            End With
        Next
        Set mcTree.ActiveNode = mcTree.ActiveNode
    End If

End Sub

Private Sub cmdRemoveNode_Click()
' PT Remove the selected Node and all its children
'    the approach is slightly different to the original Treeview's

    Dim sMsg As String
    Dim cNode As clsNode

    On Error GoTo errH
    If mcTree Is Nothing Then Exit Sub

    Set cNode = mcTree.ActiveNode
    If cNode Is Nothing Then Exit Sub

    sMsg = "Are you sure you want to Remove " & cNode.Caption
    If Not cNode.GetChild(-1) Is Nothing Then
        sMsg = sMsg & vbCr & "and all its Child Nodes"
    End If
    sMsg = sMsg & " ?"

    If MsgBox(sMsg, vbOKCancel, AppName) = vbOK Then

        mcTree.NodeRemove cNode

        mcTree.Refresh  ' refresh after all deleted nodes are removed

    End If

    Exit Sub
errH:
    MsgBox Err.Source & vbCr & Err.Description, , AppName
End Sub

Private Sub cmdAddSibling_Click()
' PT add a sibling to the active node
'    there are two approaches -

    Dim cNode As clsNode
    Dim cParent As clsNode
    Dim cSibling As clsNode
    Dim vIcon1, vIcon2
    Static lNewItem As Long

    If mcTree Is Nothing Then Exit Sub

    On Error GoTo errH
    Set cNode = mcTree.ActiveNode
    If cNode Is Nothing Then Exit Sub

    ' as this is a demo assume we want similar icons (if any)
    vIcon1 = cNode.ImageMain
    vIcon2 = cNode.ImageExpanded
    lNewItem = lNewItem + 1

    '''' Approach-1: use the Node.AddChild method, add as a child to the parent''''''''''''
    ''''             in this example we will sort all the child nodes when done

    '    Set cParent = cNode.ParentNode ' we add to the parent
    '
    '    Set cSibling = cParent.AddChild("MyUniqueSiblingKey" & lNewItem, _
         '                                  "New Sibling of " & cNode.Caption & " #" & lNewItem, _
         '                                  vIcon1, vIcon2)
    '    cSibling.Expanded = False
    '    cParent.Sort

    ''''''''''''''''''end approach-1''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''Approach-2: use the old style NodeAdd method
    ''''            in this example we can include vRelationship:=tvNext t add after the activenode

    If cNode.ParentNode.Caption = "RootHolder" Then
        'We have a root node, add another root node
        Set cSibling = mcTree.NodeAdd(, , _
                                    "MyUniqueSiblingKey" & lNewItem, _
                                    "New Sibling of " & cNode.Caption & " #" & lNewItem, _
                                    vIcon1, vIcon2)
    Else
    'mcTree.tvTreeRelationship.tvNext
        Set cSibling = mcTree.NodeAdd(cNode, 2, _
                                    "MyUniqueSiblingKey" & lNewItem, _
                                    "New Sibling of " & cNode.Caption & " #" & lNewItem, _
                                    vIcon1, vIcon2)
    End If
    cSibling.Expanded = False
    
    ''''''''''''''''' end approach-2''''''''''''''''''''''''''''''''''''''''''''''''''''


    mcTree.Refresh  ' refresh the tree after adding all new nodes

    '  Set mcTree.ActiveNode = cSibling ' could activate the new sibling
    mcTree.ScrollToView cSibling, Top1Bottom2:=2    ' reset scrolltop if necessary to view the new sibling
    Exit Sub

errH:

    MsgBox Err.Source & vbCr & Err.Description, , AppName
    Stop
    Resume
End Sub

Private Sub cmdAddChild_Click()
' PT add a child node to the active node

    Dim cNode As clsNode
    Dim cChild As clsNode
    Dim vIcon1, vIcon2
    Static lNewItem As Long

    On Error GoTo errH
    If mcTree Is Nothing Then Exit Sub

    Set cNode = mcTree.ActiveNode
    If cNode Is Nothing Then Exit Sub

    If Not cNode.ChildNodes Is Nothing Then
        If cNode.ChildNodes.Count Then
            ' as this is a demo assume we want similar icons to the first child (if any)
            vIcon1 = cNode.ChildNodes(1).ImageMain
            vIcon2 = cNode.ChildNodes(1).ImageExpanded
        End If
    End If

    lNewItem = lNewItem + 1

    Set cChild = cNode.AddChild("Key" & lNewItem, _
                                "New Child of " & cNode.Caption & " #" & lNewItem, _
                                vIcon1, vIcon2)

    ''' we could also use the NodeAdd method

    '    Set cChild = mcTree.NodeAdd(cNode, tvChild, "Key" & lNewItem, _
         '                                "New Child of " & cNode.Caption & " #" & lNewItem, _
         '                                vIcon1, vIcon2)

    If Len(vIcon2) = 0 Then
        cChild.Expanded = False    ' don't want to shownded icon
    End If
    cNode.Expanded = True
    mcTree.Refresh    ' refresh the tree after adding all new nodes

    '  Set mcTree.ActiveNode = cChild ' could activate the new child
    mcTree.ScrollToView cChild, 2    ' reset scrolltop if necessary to view the new child
    Exit Sub
errH:
    MsgBox Err.Source & vbCr & Err.Description, , AppName
End Sub

Private Sub cmdSiblings_Click()
' PT similar to the original Treeview's FirstSibling & LastSibling
    Dim s1 As String, s2 As String
    Dim cNode As clsNode

    If Not mcTree Is Nothing Then

        Set cNode = mcTree.ActiveNode
        If cNode Is Nothing Then Exit Sub
        
        s1 = cNode.Caption & vbCr
        Set cNode = cNode.FirstSibling
        If Not cNode Is Nothing Then
            s2 = "First sibling: " & cNode.Caption & vbCr
        End If
        Set cNode = mcTree.ActiveNode.LastSibling
        If Not cNode Is Nothing Then

            s2 = s2 & "Last sibling: " & cNode.Caption
        End If
        If Len(s2) = 0 Then s2 = "no siblings"

        MsgBox s1 & s2, , AppName
    End If
End Sub

Private Sub cmdGetData_Click()
    Dim lCnt As Long
    Dim ws As Worksheet
    Dim cRoot As clsNode
    If mcTree Is Nothing Then
        MsgBox "Tree must contain nodes"
        Exit Sub
    End If
    mcTree.ActiveNode.Key = Now
    If mcTree Is Nothing Then Exit Sub

    On Error Resume Next
    Set ws = Worksheets("DataDump")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(, ActiveSheet)
        ws.Name = "DataDump"
    Else
        '  ws.UsedRange.Clear
    End If

    '' Dump all treeview data
    For Each cRoot In mcTree.RootNodes
        GetData1 cRoot, lCnt, 0, ws.Range("A1")
    Next

    '    ''' Dump the selected node's + childnodes' data
    '    GetData1 mcTree.ActiveNode, 0, 0, ws.Range("A1")

    '    '' Dump a snapshot of the visible/expanded nodes
    '    For Each cRoot In mcTree.RootNodes
    '       GetData2 cRoot, 0, 0, ws.Range("A1")
    '    Next

    ws.Activate
End Sub

Sub GetData1(cParent As clsNode, lCt As Long, ByVal lLevel As Long, rng As Range)
' PT, a recursively retrieve the node's decendent data
    Dim cChild As clsNode

    lLevel = lLevel + 1
    lCt = lCt + 1

    rng(lCt, lLevel) = cParent.Caption
    If Not cParent.ChildNodes Is Nothing Then
        For Each cChild In cParent.ChildNodes
            GetData1 cChild, lCt, lLevel, rng
        Next
    End If

End Sub

Sub GetData2(cParent As clsNode, lCt As Long, _
             ByVal lLevel As Long, rng As Range)
' PT, a recursively retrieve the node's decendent data
'     but only if visible, a snapshot of the current view
Dim cChild As clsNode
    lLevel = lLevel + 1
    lCt = lCt + 1

    rng(lCt, lLevel) = cParent.Caption
    If Not cParent.ChildNodes Is Nothing Then
        If cParent.Expanded Then
            For Each cChild In cParent.ChildNodes
                GetData2 cChild, lCt, lLevel, rng
            Next
        End If
    End If
End Sub

'Sub ShowMyNodes(strSnippet As String, cParent As clsNode, lCt As Long, _
'                ByVal lLevel As Long)
'Dim cChild As clsNode
'
'    lLevel = lLevel + 1
'    lCt = lCt + 1
'
'    If InStr(1, cParent.Caption, strSnippet, vbTextCompare) Then
'        mcTree.ScrollToView cParent
'    End If
'    If Not cParent.ChildNodes Is Nothing Then
'
'            For Each cChild In cParent.ChildNodes
'                ShowMyNodes strSnippet, cChild, lCt, lLevel
'            Next
'
'    End If
'End Sub

'Private Sub UpdateInfoLabel()
'    If Not mcTree Is Nothing Then
'        With mcTree
'            Me.labInfo.Caption = "Indentation:  " & .Indentation & vbCr & _
'                                 "NodeHeight:   " & .NodeHeight
'        End With
'    End If
'End Sub

'/########### Some Read/Write examples behind the buttons #########################



