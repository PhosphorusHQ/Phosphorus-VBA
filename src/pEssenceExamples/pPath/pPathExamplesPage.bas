VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pPathExamplesPage"
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

Private Type PageAttributes
  Workbook As Workbook
  CurrentpPath As String
  MasterWindow As pLocator
  TestLocator As pLocator
  MatchingElementspPaths() As String
  MatchingElements() As UIAutomationClient.IUIAutomationElement
End Type

Private This As PageAttributes

Private Sub Class_Initialize()
  OpenNewWorkBook
  InitialiseAllLocators
End Sub

Private Sub Class_Terminate()
  CloseWorkbook
  DestroyLocators
End Sub

Private Sub DestroyLocators()
  Set This.MasterWindow = Nothing
  Set This.TestLocator = Nothing
End Sub

Private Sub OpenNewWorkBook()
  Set This.Workbook = Workbooks.Add
  This.Workbook.Activate
  
  'Populate the function test data
  Dim ws As Worksheet
  Dim myVariantArray As Variant
  Set ws = This.Workbook.Sheets("Sheet1")
  With ws
  
    .Range("A1") = "Test Data"
  
    myVariantArray = Array("Cat", "Dog", "", "Rabbit", " " & "One" & vbCrLf & "Two" & vbCr & "Three" & vbLf & "Four" & vbTab & "Five" & "        " & "Six" & " ")
    .Range("A2:E2") = myVariantArray
  
    myVariantArray = Array(1, 2, 3, 4, -1)
    .Range("A3:E3") = myVariantArray

    myVariantArray = Array(101.45, 22.7, 365.25, 404.36, 521.01)
    .Range("A4:E4") = myVariantArray

    myVariantArray = Array("https://www.google.co.uk/", "https://www.bbc.co.uk/")
    .Range("A5:B5") = myVariantArray

    myVariantArray = Array("31/5/2021", "1st April 1992", "19/5/2025 14:43", "April 6, 2003")
    .Range("A6:D6") = myVariantArray
  
    myVariantArray = Array(True, False, "=1=1", "=1=2", "'True", "'False")
    .Range("A7:F7") = myVariantArray
    
    .Cells.EntireColumn.AutoFit
  
  End With
  
  VBA.Interaction.DoEvents

End Sub

Public Sub CloseWorkbook()
  This.Workbook.Close SaveChanges:=False
  Set This.Workbook = Nothing
End Sub

Private Sub InitialiseAllLocators()
  Set This.MasterWindow = Factory.GetNewLocator
  This.CurrentpPath = "/Window[And(@Name=""" & This.Workbook.Name & " - Excel"",@ClassName=""XLMAIN"")]"
  With This.MasterWindow
    .Initialise "MasterWindow", Nothing, None, pPath, This.CurrentpPath
    On Error Resume Next
    .Find 10
    On Error GoTo 0
    Debug.Assert .ReturnedValue = True
    Debug.Assert .NumberOfMatchingElements = 1
    Debug.Assert .Element.UIAElement.CurrentName = This.Workbook.Name & " - Excel"
    This.MatchingElementspPaths = .MatchingElementspPaths
    'Square brackets are reserved characters for 'Like' so we need to escape these with []!
    Debug.Assert This.MatchingElementspPaths(1) Like "/Window[[]*[]][[]@Name='" & This.Workbook.Name & " - Excel'[]]"
    Window.HighlightElement .Element.UIAElement, BorderColor:=&H808000, MultiHighlight:=True, DelayMs:=0
    Snooze 1000
    Window.ReleaseHighlighting
  End With
  'These test assume that we are on the HOME tab unless specified otherwise
  pEssence.OfficeRibbon.SelectTab This.MasterWindow, "Home"

End Sub

Private Sub PreValidation(ElementName As String, pPathString As String, ExpectedErrorMessage As String)
  Set This.TestLocator = Nothing
  Set This.TestLocator = Factory.GetNewLocator
  This.CurrentpPath = pPathString
  With This.TestLocator
    .Initialise ElementName, Nothing, None, pPath, This.CurrentpPath
    On Error Resume Next
    .Find 0
    On Error GoTo 0
    If .ReturnedValue = "" Then
      Debug.Assert .ReturnedValue = ""
    Else
      Debug.Assert .ReturnedValue = False
    End If
    Debug.Assert .ErrorMessage = ExpectedErrorMessage
    Debug.Assert .NumberOfMatchingElements = 0
  End With
End Sub

Public Sub PreValidation_Test01()
  PreValidation "PreValidation_Test01", "((((()))", "Unequal number of left & right parentheses '(' & ')' in pPath!"
End Sub

Public Sub PreValidation_Test02()
  PreValidation "PreValidation_Test02", "[[[]]]]]", "Unequal number of left & right square brackets '[' & ']' in pPath!"
End Sub

Public Sub PreValidation_Test03()
  PreValidation "PreValidation_Test03", "@NotAUIProperty", "Unrecognised property in pPath!"
End Sub

Public Sub PreValidation_Test04()
  PreValidation "PreValidation_Test04", "./..*", "pPath is relative but no context node is provided!"
End Sub

Public Sub PreValidation_Test05()
  Set This.TestLocator = Nothing
  Set This.TestLocator = Factory.GetNewLocator
  This.CurrentpPath = "/Edit"
  With This.TestLocator
    .Initialise "PreValidation_Test05", Nothing, None, pPath, This.CurrentpPath
    .AddpPathContextNode This.MasterWindow, ""
    On Error Resume Next
    .Find 0
    On Error GoTo 0
    Debug.Assert .ErrorMessage = "A context node is provided but isn't used!"
    Debug.Assert .NumberOfMatchingElements = 0
  End With
End Sub

Public Sub PreValidation_Test06()
  Set This.TestLocator = Nothing
  Set This.TestLocator = Factory.GetNewLocator
  This.CurrentpPath = ".//* | .//* | .//*"
  With This.TestLocator
    .Initialise "PreValidation_Test06", Nothing, None, pPath, This.CurrentpPath
    .AddpPathContextNode This.MasterWindow, ""
    .AddpPathContextNode This.MasterWindow, ""
    On Error Resume Next
    .Find 0
    On Error GoTo 0
    Debug.Assert .ReturnedValue = ""
    Debug.Assert .ErrorMessage = "The number of nodes and relative pPaths do not match!"
    Debug.Assert .NumberOfMatchingElements = 0
  End With
End Sub

Public Sub PreValidation_Test07()
  PreValidation "PreValidation_Test07", "([)]", "Unexpected ')' Bracket at position 3"
End Sub

Public Sub PreValidation_Test08()
  PreValidation "PreValidation_Test08", "(()])[", "Unexpected ']' Bracket at position 4"
End Sub

Public Sub PreValidation_Test09()
  PreValidation "PreValidation_Test09", "][", "Unexpected ']' Bracket at position 1"
End Sub

Public Sub PreValidation_Test10()
  PreValidation "PreValidation_Test10", "[)](", "Unexpected ')' Bracket at position 2"
End Sub

Public Sub PreValidation_Test11()
  PreValidation "PreValidation_Test11", "//*[And(count(./Button)>=1,count(./SplitButton)>1,count(./ComboBox)>1)]/Button[position)(=6]", "Unexpected ')' Bracket at position 88"
End Sub

Public Sub PreValidation_Test12()
  PreValidation "PreValidation_Test12", "/button", "Invalid Node Test 'button'!"
End Sub

Public Sub PreValidation_Test13()
  PreValidation "PreValidation_Test13", "/abc", "Invalid Node Test 'abc'!"
End Sub

Public Sub PreValidation_Test14()
  PreValidation "PreValidation_Test14", "//Button/attribute::isenabled", "Invalid Node Test 'isenabled'!"
End Sub

Public Sub PreValidation_Test15()
  PreValidation "PreValidation_Test15", "//Button/attribute::NotANavigableAttribute", "Invalid Node Test 'NotANavigableAttribute'!"
End Sub

Private Sub Evaluation(ElementName As String, pPathString As String, ExpectedNumberOfMatchingElements As Long, Optional RootUIAElementLocator As pLocator, Optional ContextNodes As Variant, Optional InitialpPaths As Variant, Optional ReturnedValue As Variant, Optional UnitTestingMode As Boolean)
  Set This.TestLocator = Nothing
  Set This.TestLocator = Factory.GetNewLocator
  This.CurrentpPath = pPathString
  With This.TestLocator
    .SetUnitTestingMode UnitTestingMode
    .Initialise ElementName, RootUIAElementLocator, None, pPath, This.CurrentpPath
    Dim NumberOfContextNodes As Integer
    NumberOfContextNodes = Phosphorus.Utils.GetSizeOfArray(ContextNodes) + 1
    If Not IsMissing(ContextNodes) And NumberOfContextNodes > 0 Then
      Dim Counter As Integer
      For Counter = 1 To NumberOfContextNodes
        Dim CurrentContextNode As pLocator
        Dim CurrentInitialpPath As String
        Set CurrentContextNode = ContextNodes(Counter - 1)
        CurrentInitialpPath = InitialpPaths(Counter - 1)
        .AddpPathContextNode CurrentContextNode, CurrentInitialpPath
      Next Counter
    End If
    .FindAll TimeoutInSeconds:=0, AcceptNoElements:=True
    If IsMissing(ReturnedValue) Then
      ReturnedValue = (.NumberOfMatchingElements > 0)
    End If
    Debug.Assert .ReturnedValue = ReturnedValue
    Debug.Assert .ErrorMessage = ""
    Debug.Assert .NumberOfMatchingElements = ExpectedNumberOfMatchingElements
    This.MatchingElements = This.TestLocator.MatchingElements
    Dim i As Long
    If .NumberOfMatchingElements >= 1 Then
      For i = 1 To .NumberOfMatchingElements
        Window.HighlightElement This.MatchingElements(i), BorderColor:=&H808000, MultiHighlight:=True, DelayMs:=0
      Next i
      Snooze 1000
      Window.ReleaseHighlighting
    End If
  End With
End Sub

Public Sub Evaluation_Test001()
  PreValidation "Evaluation_Test001", "", "Some PPath expression must be specified!"
End Sub

Public Sub Evaluation_Test002()
  PreValidation "Evaluation_Test002", "( | )", "PPath #1: Some PPath expression must be specified!"
End Sub

Public Sub Evaluation_Test003()
  PreValidation "Evaluation_Test003", "/", "Missing Node Test!"
End Sub

Public Sub Evaluation_Test004()
  PreValidation "Evaluation_Test004", "    /    ", "Missing Node Test!"
End Sub

Public Sub Evaluation_Test005()
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]/Pane/ScrollBar"
    .Find 10
  End With
  Evaluation "Evaluation_Test005", "/*", 5, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/Button[1][@Name='Line up']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Thumb[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(5) = "/Button[4][@Name='Line down']"
  Set RootTestLocator = Nothing
End Sub

Public Sub Evaluation_Test006()
  PreValidation "Evaluation_Test006", "/*a", "Illegal start Of predicate! 'a'"
End Sub

Public Sub Evaluation_Test007()
  PreValidation "Evaluation_Test007", " /* | / ", "Missing Node Test!"
End Sub

Public Sub Evaluation_Test008()
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]/Pane/ScrollBar"
    .Find 10
  End With
  Evaluation "Evaluation_Test008", "(/* | /*)", 5, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/Button[1][@Name='Line up']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Thumb[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(5) = "/Button[4][@Name='Line down']"
  Set RootTestLocator = Nothing
End Sub

Public Sub Evaluation_Test009()
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]"
    .Find 10
  End With
  Evaluation "Evaluation_Test009", "//*", 7, RootTestLocator
  Evaluation_Test009_pPaths
  Set RootTestLocator = Nothing
End Sub

Private Sub Evaluation_Test009_pPaths()
  'NOTE: The zoom & zoom slider seems to display a thumbnail dynamically so this has been disabled from the status bar ... Status Bar > Right Mouse Click > Unticked
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[1][@Name='Line up']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(5) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Thumb[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(6) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(7) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[4][@Name='Line down']"
End Sub

Public Sub Evaluation_Test010()
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]"
    .Find 10
  End With
  Evaluation "Evaluation_Test010", "//* | //*", 7, RootTestLocator
  Evaluation_Test009_pPaths
  Set RootTestLocator = Nothing
End Sub

Public Sub Evaluation_Test011()
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, "/Pane[@ClassName=""XLDESK""]"
    .Find 10
  End With
  Evaluation "Evaluation_Test011", "/Pane", 4, RootTestLocator
  Evaluation_Test011_pPaths
  Set RootTestLocator = Nothing
End Sub

Private Sub Evaluation_Test011_pPaths()
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Vertical']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[2][@Name='Horizontal']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Pane[3][@Name='']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Pane[4][@Name='" & This.Workbook.Name & "']"
End Sub

Public Sub Evaluation_Test012()
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, "/Pane[@ClassName=""XLDESK""]"
    .Find 10
  End With
  Evaluation "Evaluation_Test012", "/Pane | /Pane", 4, RootTestLocator
  Evaluation_Test011_pPaths
  Set RootTestLocator = Nothing
End Sub

Public Sub Evaluation_Test013()
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, "/Pane[@ClassName=""XLDESK""]"
    .Find 10
  End With
  Evaluation "Evaluation_Test013", "//Pane", 7, RootTestLocator
  Evaluation_Test013_pPaths
  Set RootTestLocator = Nothing
End Sub

Private Sub Evaluation_Test013_pPaths()
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Vertical']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[1][@Name='Vertical']/Pane[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Pane[2][@Name='Horizontal']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Pane[2][@Name='Horizontal']/Pane[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(5) = "/Pane[3][@Name='']"
  Debug.Assert This.MatchingElementspPaths(6) = "/Pane[4][@Name='" & This.Workbook.Name & "']"
  Debug.Assert This.MatchingElementspPaths(7) = "/Pane[4][@Name='" & This.Workbook.Name & "']/Pane[1][@Name='Sheet Sheet1']"
End Sub

Public Sub Evaluation_Test014()
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, "/Pane[@ClassName=""XLDESK""]"
    .Find 10
  End With
  Evaluation "Evaluation_Test014", "/child::*", 4, RootTestLocator
  Evaluation_Test011_pPaths
  Set RootTestLocator = Nothing
End Sub

Public Sub Evaluation_Test015()
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]"
    .Find 10
  End With
  Evaluation "Evaluation_Test015", "/descendant::*", 7, RootTestLocator
  Evaluation_Test009_pPaths
  Set RootTestLocator = Nothing
End Sub

Public Sub Evaluation_Test016()
  Evaluation "Evaluation_Test016", "/.*", 1, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "."
End Sub

Public Sub Evaluation_Test017()
  Evaluation "Evaluation_Test017", "/self::*", 1, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "."
End Sub

Public Sub Evaluation_Test018()
  Evaluation "Evaluation_Test018", "/..*", 0, This.MasterWindow
End Sub

Public Sub Evaluation_Test019()
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  Dim pPathString As String
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[1]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test019", "./..*", 1, This.MasterWindow, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "(/Pane[@ClassName=""EXCEL2""])[1]/../Window[@Name='" & This.Workbook.Name & " - Excel']"
  Set ContextNodeLocator = Nothing
End Sub

Public Sub Evaluation_Test020()
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  Dim pPathString As String
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[1]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test020", "./parent::*", 1, This.MasterWindow, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = pPathString & "/../Window[@Name='" & This.Workbook.Name & " - Excel']"
  Set ContextNodeLocator = Nothing
End Sub

Public Sub Evaluation_Test021()
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  Dim pPathString As String
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//TabItem[@Name=""Home""]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test021", "./ancestor::*", 8, This.MasterWindow, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = pPathString & "/../Tab[@Name='Ribbon Tabs']"
  Debug.Assert This.MatchingElementspPaths(2) = This.MatchingElementspPaths(1) & "/../Pane[@Name='Ribbon']"
  Debug.Assert This.MatchingElementspPaths(3) = This.MatchingElementspPaths(2) & "/../Pane[@Name='']"
  Debug.Assert This.MatchingElementspPaths(4) = This.MatchingElementspPaths(3) & "/../Pane[@Name='']"
  Debug.Assert This.MatchingElementspPaths(5) = This.MatchingElementspPaths(4) & "/../Pane[@Name='Ribbon']"
  Debug.Assert This.MatchingElementspPaths(6) = This.MatchingElementspPaths(5) & "/../ToolBar[@Name='']"
  Debug.Assert This.MatchingElementspPaths(7) = This.MatchingElementspPaths(6) & "/../Pane[@Name='']"
  Debug.Assert This.MatchingElementspPaths(8) = This.MatchingElementspPaths(7) & "/../Window[@Name='" & This.Workbook.Name & " - Excel']"
  Set ContextNodeLocator = Nothing
End Sub

Public Sub Evaluation_Test022()
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  Dim pPathString As String
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//TabItem[@Name=""Home""]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test022", "./ancestor-or-self::*", 9, This.MasterWindow, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = pPathString & "/."
  Debug.Assert This.MatchingElementspPaths(2) = pPathString & "/../Tab[@Name='Ribbon Tabs']"
  Debug.Assert This.MatchingElementspPaths(3) = This.MatchingElementspPaths(2) & "/../Pane[@Name='Ribbon']"
  Debug.Assert This.MatchingElementspPaths(4) = This.MatchingElementspPaths(3) & "/../Pane[@Name='']"
  Debug.Assert This.MatchingElementspPaths(5) = This.MatchingElementspPaths(4) & "/../Pane[@Name='']"
  Debug.Assert This.MatchingElementspPaths(6) = This.MatchingElementspPaths(5) & "/../Pane[@Name='Ribbon']"
  Debug.Assert This.MatchingElementspPaths(7) = This.MatchingElementspPaths(6) & "/../ToolBar[@Name='']"
  Debug.Assert This.MatchingElementspPaths(8) = This.MatchingElementspPaths(7) & "/../Pane[@Name='']"
  Debug.Assert This.MatchingElementspPaths(9) = This.MatchingElementspPaths(8) & "/../Window[@Name='" & This.Workbook.Name & " - Excel']"
  Set ContextNodeLocator = Nothing
End Sub

Public Sub Evaluation_Test023()
'See: Evaluation_Test015
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]"
    .Find 10
  End With
  Evaluation "Evaluation_Test023", "/descendant-or-self::*", 8, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "."
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[1][@Name='Line up']"
  Debug.Assert This.MatchingElementspPaths(5) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(6) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Thumb[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(7) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(8) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[4][@Name='Line down']"
  Set RootTestLocator = Nothing
End Sub

Public Sub Evaluation_Test024()
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  Dim pPathString As String
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]//Button[4]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test024", "./preceding-sibling::*", 4, This.MasterWindow, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = pPathString & "/../Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(2) = pPathString & "/../Thumb[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(3) = pPathString & "/../Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(4) = pPathString & "/../Button[1][@Name='Line up']"
  Set ContextNodeLocator = Nothing
End Sub

Public Sub Evaluation_Test025()
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  Dim pPathString As String
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]//Button[1]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test025", "./following-sibling::*", 4, This.MasterWindow, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = pPathString & "/../Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(2) = pPathString & "/../Thumb[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(3) = pPathString & "/../Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(4) = pPathString & "/../Button[4][@Name='Line down']"
  Set ContextNodeLocator = Nothing
End Sub

Public Sub Evaluation_Test026()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[1]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  'Nothing before first element test
  Evaluation "Evaluation_Test026", "/preceding::*", 0, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test027()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  'Lots before last child root element test
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  pPathString = "//Button[4]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", RootTestLocator, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test027", "./preceding::*", 6, RootTestLocator, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Thumb[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[1][@Name='Line up']"
  Debug.Assert This.MatchingElementspPaths(5) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(6) = "/Pane[1][@Name='']"
  
  Set RootTestLocator = Nothing
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test028()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  'Lots before middle low level element test
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  pPathString = "//Button[4]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", RootTestLocator, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test028", "./preceding::*", 6, RootTestLocator, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Thumb[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[1][@Name='Line up']"
  Debug.Assert This.MatchingElementspPaths(5) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(6) = "/Pane[1][@Name='']"

  Set RootTestLocator = Nothing
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test029()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/*[last()]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  'Nothing after last root element test
  Evaluation "Evaluation_Test029", "/following::*", 0, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test030()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  pPathString = "//Button[1]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", RootTestLocator, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test030", "./following::*", 4, RootTestLocator, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Thumb[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[4][@Name='Line down']"
  
  Set RootTestLocator = Nothing
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test031()
   
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  pPathString = "//Button[@Name=""Line up""]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", RootTestLocator, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test031", "./following::*", 4, RootTestLocator, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Thumb[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[4][@Name='Line down']"
  
  Set RootTestLocator = Nothing
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test032()
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, "/Pane[@ClassName=""XLDESK""]"
    .Find 10
  End With
  Evaluation "Evaluation_Test032", "/child::Pane", 4, RootTestLocator
  Evaluation_Test011_pPaths
  Set RootTestLocator = Nothing
End Sub

Public Sub Evaluation_Test033()
   
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar//ToolBar[@Name=""Quick Access Toolbar""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test033", "/descendant::Button", 4, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/Button[1][@Name='AutoSave']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Button[2][@Name='Save']"
  Debug.Assert This.MatchingElementspPaths(3) = "/SplitButton[1][@Name='Undo']/Button[1][@Name='Undo']" Or _
               This.MatchingElementspPaths(3) = "/Button[3][@Name='Undo']"
  Debug.Assert This.MatchingElementspPaths(4) = "/SplitButton[2][@Name='Redo']/Button[1][@Name='Redo']" Or _
               This.MatchingElementspPaths(4) = "/Button[4][@Name='Redo']"
 
 Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test034()
  Evaluation "Evaluation_Test034", "/.Window", 1, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "."
End Sub

Public Sub Evaluation_Test035()
  Evaluation "Evaluation_Test035", "/self::Window", 1, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "."
End Sub

Public Sub Evaluation_Test036()
  Evaluation "Evaluation_Test036", "/..Window", 0, This.MasterWindow
End Sub

Public Sub Evaluation_Test037()
   
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  Dim pPathString As String
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test037", "./..Window", 1, This.MasterWindow, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "(/Pane[@ClassName=""EXCEL2""])[2]/../Window[@Name='" & This.Workbook.Name & " - Excel']"
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test038()
   
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  Dim pPathString As String
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar//ToolBar[@Name=""Quick Access Toolbar""]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test038", "./parent::*", 1, This.MasterWindow, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar//ToolBar[@Name=""Quick Access Toolbar""]/../Pane[@Name='Ribbon']"
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test039()
   
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  Dim pPathString As String
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar//ToolBar[@Name=""Quick Access Toolbar""]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test039", "./ancestor::ToolBar", 1, This.MasterWindow, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar//ToolBar[@Name=""Quick Access Toolbar""]/../Pane[@Name='Ribbon']/../Pane[@Name='']/../Pane[@Name='']/../Pane[@Name='Ribbon']/../ToolBar[@Name='']"
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test040()
   
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  Dim pPathString As String
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar//ToolBar[@Name=""Quick Access Toolbar""]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test040", "./ancestor-or-self::Window", 1, This.MasterWindow, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar//ToolBar[@Name=""Quick Access Toolbar""]/../Pane[@Name='Ribbon']/../Pane[@Name='']/../Pane[@Name='']/../Pane[@Name='Ribbon']/../ToolBar[@Name='']/../Pane[@Name='']/../Window[@Name='" & This.Workbook.Name & " - Excel']"
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test041()
   
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  Dim pPathString As String
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test041", "./descendant-or-self::Pane", 6, This.MasterWindow, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "(/Pane[@ClassName=""EXCEL2""])[2]/."
  Debug.Assert This.MatchingElementspPaths(2) = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']"
  Debug.Assert This.MatchingElementspPaths(3) = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(4) = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(5) = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']"
  Debug.Assert This.MatchingElementspPaths(6) = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']"
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test042()
   
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  Dim pPathString As String
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]//Button[4]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test042", "./preceding-sibling::Button", 3, This.MasterWindow, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = pPathString & "/../Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(2) = pPathString & "/../Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(3) = pPathString & "/../Button[1][@Name='Line up']"
  
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test043()
   
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  Dim pPathString As String
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]//Button[1]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test043", "./following-sibling::Button", 3, This.MasterWindow, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = pPathString & "/../Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(2) = pPathString & "/../Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(3) = pPathString & "/../Button[4][@Name='Line down']"
  
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test044()
   
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  pPathString = "//Button[4]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", RootTestLocator, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test044", "./preceding::Button", 3, RootTestLocator, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[1][@Name='Line up']"
 
  Set RootTestLocator = Nothing
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test045()
   
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""Vertical""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  pPathString = "//Button[1]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", RootTestLocator, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test045", "./following::Button", 3, RootTestLocator, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[4][@Name='Line down']"
  
  Set RootTestLocator = Nothing
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test046()
   
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  pPathString = "/Pane[@Name=""Vertical""]//Button[1]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", RootTestLocator, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test046", "./attribute::*", 118, RootTestLocator, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
'  Debug.Assert This.MatchingElementspPaths(1) = pPathString & "/@BoundingRectangle='{1897, 272, 21, 21}'"
  Debug.Assert This.MatchingElementspPaths(2) = pPathString & "/@ClassName='NetUIRepeatButton'"
  Debug.Assert This.MatchingElementspPaths(3) = pPathString & "/@ControlType=50000"
  '...
  Debug.Assert This.MatchingElementspPaths(116) = pPathString & "/@WindowIsTopmost=False"
  Debug.Assert This.MatchingElementspPaths(117) = pPathString & "/@WindowWindowInteractionState=0"
  Debug.Assert This.MatchingElementspPaths(118) = pPathString & "/@WindowWindowVisualState=0"
  
  Set RootTestLocator = Nothing
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test047()
   
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  pPathString = "//ToolBar//Pane[And(@Name=""Ribbon"",@ClassName=""NetUInetpane"")]/Button[4]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", RootTestLocator, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test047", "./attribute::IsEnabled", 1, RootTestLocator, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = pPathString & "/@IsEnabled=True"
  Set RootTestLocator = Nothing
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test048()
   
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  pPathString = "//ToolBar//ToolBar[@Name=""Quick Access Toolbar""]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", RootTestLocator, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test048", "./node()", 5, RootTestLocator, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = pPathString & "//Button[1][@Name='AutoSave']"
  Debug.Assert This.MatchingElementspPaths(2) = pPathString & "//Button[2][@Name='Save']"
  Debug.Assert This.MatchingElementspPaths(3) = pPathString & "//SplitButton[1][@Name='Undo']" Or _
               This.MatchingElementspPaths(3) = pPathString & "//Button[3][@Name='Undo']"
  Debug.Assert This.MatchingElementspPaths(4) = pPathString & "//SplitButton[2][@Name='Redo']" Or _
               This.MatchingElementspPaths(4) = pPathString & "//Button[4][@Name='Redo']"
  Debug.Assert This.MatchingElementspPaths(5) = pPathString & "//MenuItem[1][@Name='Customize Quick Access Toolbar']"
  Set RootTestLocator = Nothing
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test049()
   
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  pPathString = "//ToolBar//ToolBar[@Name=""Quick Access Toolbar""]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", RootTestLocator, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test049", "./element()", 5, RootTestLocator, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = pPathString & "//Button[1][@Name='AutoSave']"
  Debug.Assert This.MatchingElementspPaths(2) = pPathString & "//Button[2][@Name='Save']"
  Debug.Assert This.MatchingElementspPaths(3) = pPathString & "//SplitButton[1][@Name='Undo']" Or _
               This.MatchingElementspPaths(3) = pPathString & "//Button[3][@Name='Undo']"
  Debug.Assert This.MatchingElementspPaths(4) = pPathString & "//SplitButton[2][@Name='Redo']" Or _
               This.MatchingElementspPaths(4) = pPathString & "//Button[4][@Name='Redo']"
  Debug.Assert This.MatchingElementspPaths(5) = pPathString & "//MenuItem[1][@Name='Customize Quick Access Toolbar']"
  Set RootTestLocator = Nothing
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test050()
Exit Sub
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  pPathString = "//ToolBar//Pane[And(@Name=""Ribbon"",@ClassName=""NetUInetpane"")]//Group[@Name=""Font""]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", RootTestLocator, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_Test050", ".//text()", 2, RootTestLocator, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  If This.TestLocator.NumberOfMatchingElements = 2 Then
    Debug.Assert This.MatchingElementspPaths(1) = pPathString & "//ComboBox[1][@Name='Font']/Edit[1][@Name='Font']/text()"
    Debug.Assert This.MatchingElementspPaths(2) = pPathString & "//ComboBox[2][@Name='Font Size']/Edit[1][@Name='Font Size']/text()"
  End If
  Set RootTestLocator = Nothing
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test051()

End Sub

Public Sub Evaluation_Test052()
      
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test052", "//element(*, integer)", 5, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[4][@Name='" & This.Workbook.Name & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[67][@Name='A3']/@value"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[4][@Name='" & This.Workbook.Name & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[68][@Name='B3']/@value"
  Debug.Assert This.MatchingElementspPaths(3) = "/Pane[4][@Name='" & This.Workbook.Name & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[69][@Name='C3']/@value"
  Debug.Assert This.MatchingElementspPaths(4) = "/Pane[4][@Name='" & This.Workbook.Name & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[70][@Name='D3']/@value"
  Debug.Assert This.MatchingElementspPaths(5) = "/Pane[4][@Name='" & This.Workbook.Name & "']/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[71][@Name='E3']/@value"
    
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test053()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test053", "//element(*, string)", 6, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[23][@Name='A1']/@value"
  Debug.Assert This.MatchingElementspPaths(2) = "/DataItem[45][@Name='A2']/@value"
  Debug.Assert This.MatchingElementspPaths(3) = "/DataItem[46][@Name='B2']/@value"
  Debug.Assert This.MatchingElementspPaths(4) = "/DataItem[48][@Name='D2']/@value"
  Debug.Assert This.MatchingElementspPaths(5) = "/DataItem[49][@Name='E2']/@value"
  Debug.Assert This.MatchingElementspPaths(6) = "/DataItem[134][@Name='B6']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test054()
  Evaluation "Evaluation_Test054", "/TitleBar/*", 4, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']"
  Debug.Assert This.MatchingElementspPaths(2) = "/TitleBar[1][@Name='']/Button[1][@Name='Minimise']"
  Debug.Assert This.MatchingElementspPaths(3) = "/TitleBar[1][@Name='']/Button[2][@Name='Restore']"
  Debug.Assert This.MatchingElementspPaths(4) = "/TitleBar[1][@Name='']/Button[3][@Name='Close']"
End Sub

Public Sub Evaluation_Test055()
    
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar//Pane[And(@Name=""Ribbon"",@ClassName=""NetUInetpane"")]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_Test055", "//Group/SplitButton", 11, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']"
  Debug.Assert This.MatchingElementspPaths(5) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']"
  Debug.Assert This.MatchingElementspPaths(6) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']"
  Debug.Assert This.MatchingElementspPaths(7) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']"
  Debug.Assert This.MatchingElementspPaths(8) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']"
  Debug.Assert This.MatchingElementspPaths(9) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']"
  Debug.Assert This.MatchingElementspPaths(10) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']"
  Debug.Assert This.MatchingElementspPaths(11) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test056()
    
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar//Pane[And(@Name=""Ribbon"",@ClassName=""NetUInetpane"")]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_Test056", "//ToolBar/attribute::*", 116, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='Quick Access Toolbar']/@BoundingRectangle='{53, 0, 300, 60}'"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='Quick Access Toolbar']/@ClassName='NetUIElement'"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='Quick Access Toolbar']/@ControlType=50021"
  '...
  Debug.Assert This.MatchingElementspPaths(114) = "/ToolBar[1][@Name='Quick Access Toolbar']/@WindowIsTopmost=False"
  Debug.Assert This.MatchingElementspPaths(115) = "/ToolBar[1][@Name='Quick Access Toolbar']/@WindowWindowInteractionState=0"
  Debug.Assert This.MatchingElementspPaths(116) = "/ToolBar[1][@Name='Quick Access Toolbar']/@WindowWindowVisualState=0"

End Sub

Public Sub Evaluation_Test057()
  Evaluation "Evaluation_Test057", "/TitleBar/node()/attribute::AutomationId", 1, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']/@AutomationId='MenuBar'"
End Sub

Public Sub Evaluation_Test058()
'Union of expressions, no context node

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_Test058", "/Pane[@Name=""Vertical""] | /Pane[@Name=""Horizontal""]", 2, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Vertical']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[2][@Name='Horizontal']"

End Sub

Public Sub Evaluation_Test059()
'Union of expressions, context node, 1 use

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  pPathString = "/Pane[@Name=""Horizontal""]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", RootTestLocator, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test059", "/Pane[@Name=""Vertical""] | .//*", 8, RootTestLocator, Array(ContextNodeLocator), Array(pPathString)
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Vertical']"
  Debug.Assert This.MatchingElementspPaths(2) = pPathString & "/Pane[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(3) = pPathString & "/Pane[1][@Name='']/ScrollBar[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(4) = pPathString & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[1][@Name='Column left']"
  Debug.Assert This.MatchingElementspPaths(5) = pPathString & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page left']"
  Debug.Assert This.MatchingElementspPaths(6) = pPathString & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Thumb[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(7) = pPathString & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[3][@Name='Page right']"
  Debug.Assert This.MatchingElementspPaths(8) = pPathString & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[4][@Name='Column right']"
  
  Set RootTestLocator = Nothing
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test060()
'Union of expressions, 1 context node, 2 uses
  
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  Dim pPathString As String
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test060", "./Pane[@Name=""Horizontal""] | ./Pane[@Name=""Vertical""]", 2, ContextNodeLocator, Array(ContextNodeLocator), Array(pPathString)
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = pPathString & "/Pane[2][@Name='Horizontal']"
  Debug.Assert This.MatchingElementspPaths(2) = pPathString & "/Pane[1][@Name='Vertical']"
  
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test061()
'Missing Context mode Initial pPath
'pLocators force us to add the Initial pPath
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  pPathString = "/Pane[@Name=""Horizontal""]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", RootTestLocator, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test061", "./ancestor::*", 1, RootTestLocator, Array(ContextNodeLocator), Array(pPathString & "/")
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = pPathString & "/../Pane[@Name='']"
  
  Set RootTestLocator = Nothing
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test062()
'Union of expressions, 3 context nodes, 3 uses
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Dim ContextNodeLocator1 As pLocator
  Dim pPathString1 As String
  Set ContextNodeLocator1 = Factory.GetNewLocator
  pPathString1 = "/Pane[@Name=""Horizontal""]"
  With ContextNodeLocator1
    .Initialise "ContextNodeLocator1", RootTestLocator, None, pPath, pPathString1
    .Find 10
  End With
   
  Dim ContextNodeLocator2 As pLocator
  Dim pPathString2 As String
  Set ContextNodeLocator2 = Factory.GetNewLocator
  pPathString2 = "/Pane[@Name=""Vertical""]"
  With ContextNodeLocator2
    .Initialise "ContextNodeLocator2", RootTestLocator, None, pPath, pPathString2
    .Find 10
  End With
   
  Dim ContextNodeLocator3 As pLocator
  Dim pPathString3 As String
  Set ContextNodeLocator3 = Factory.GetNewLocator
  pPathString3 = "/Pane[@Name=""" & This.Workbook.Name & """]"
  With ContextNodeLocator3
    .Initialise "ContextNodeLocator3", RootTestLocator, None, pPath, pPathString3
    .Find 10
  End With
   
  Evaluation "Evaluation_Test062", ".//* | .//* | ./*", 16, RootTestLocator, Array(ContextNodeLocator1, ContextNodeLocator2, ContextNodeLocator3), Array(pPathString1, pPathString2, pPathString3)
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = pPathString1 & "/Pane[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(2) = pPathString1 & "/Pane[1][@Name='']/ScrollBar[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(3) = pPathString1 & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[1][@Name='Column left']"
  Debug.Assert This.MatchingElementspPaths(4) = pPathString1 & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page left']"
  Debug.Assert This.MatchingElementspPaths(5) = pPathString1 & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Thumb[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(6) = pPathString1 & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[3][@Name='Page right']"
  Debug.Assert This.MatchingElementspPaths(7) = pPathString1 & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[4][@Name='Column right']"
  
  Debug.Assert This.MatchingElementspPaths(8) = pPathString2 & "/Pane[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(9) = pPathString2 & "/Pane[1][@Name='']/ScrollBar[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(10) = pPathString2 & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[1][@Name='Line up']"
  Debug.Assert This.MatchingElementspPaths(11) = pPathString2 & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(12) = pPathString2 & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Thumb[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(13) = pPathString2 & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(14) = pPathString2 & "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[4][@Name='Line down']"
  
  Debug.Assert This.MatchingElementspPaths(15) = pPathString3 & "/Tab[1][@Name='" & This.Workbook.Name & "']"
  Debug.Assert This.MatchingElementspPaths(16) = pPathString3 & "/Pane[1][@Name='Sheet Sheet1']"
  
  Set RootTestLocator = Nothing
  Set ContextNodeLocator1 = Nothing
  Set ContextNodeLocator2 = Nothing
  Set ContextNodeLocator3 = Nothing

End Sub

Public Sub Evaluation_Test063()
  Evaluation "Evaluation_Test063", "/TitleBar[1]", 1, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/TitleBar[1][@Name='']"
End Sub

Public Sub Evaluation_Test064()
  Evaluation "Evaluation_Test064", "(/TitleBar)[1]", 1, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/TitleBar[1][@Name='']"
End Sub

Public Sub Evaluation_Test065()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_Test065", "//Button[4]", 2, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Vertical']/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[4][@Name='Line down']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[2][@Name='Horizontal']/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[4][@Name='Column right']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test066()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test066", "(//Button[2])[2]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[2][@Name='Horizontal']/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page left']"
  
  Set RootTestLocator = Nothing
  
End Sub

Public Sub Evaluation_Test067()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test067", "/Pane[position()=2]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[2][@Name='Horizontal']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test068()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test068", "/Pane[first()]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Vertical']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test069()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test069", "(//Button[2])[first()]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Vertical']/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page up']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test070()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test070", "(//Button[2])[first()+1]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[2][@Name='Horizontal']/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page left']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test071()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test071", "(/Pane)[last()]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[4][@Name='" & This.Workbook.Name & "']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test072()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar//ToolBar[@Name=""Quick Access Toolbar""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
     
  Evaluation "Evaluation_Test072", "//Button[last()]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Button[2][@Name='Save']"
  Debug.Assert This.MatchingElementspPaths(2) = "/SplitButton[1][@Name='Undo']/Button[1][@Name='Undo']"
  Debug.Assert This.MatchingElementspPaths(3) = "/SplitButton[2][@Name='Redo']/Button[1][@Name='Redo']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test073()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//ToolBar//ToolBar[@Name=""Quick Access Toolbar""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
     
  Evaluation "Evaluation_Test073", "//Button[last()-1]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Button[1][@Name='AutoSave']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test074()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//Pane[@Name=""Lower Ribbon""]//Group[@Name=""Font""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
     
  Evaluation "Evaluation_Test074", "//Button[And(position()>1,position()<last())]", 2, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Button[2][@Name='Decrease Font Size']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Button[3][@Name='Bold']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test075()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//Pane[@Name=""Lower Ribbon""]//Group[@Name=""Font""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
     
  Evaluation "Evaluation_Test075", "(//SplitButton[Or(position()=1,position()=last())])[And(position()>1,Or(position()=2,position()=last()-1))]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/SplitButton[4][@Name='Font Color']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test076()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
     
  Evaluation "Evaluation_Test076", "count(//ScrollBar/Button)", 8, RootTestLocator, , ReturnedValue:=8
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
'  Debug.Assert This.TestLocator.ReturnedValue = 5
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Vertical']/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[1][@Name='Line up']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[1][@Name='Vertical']/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page up']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Pane[1][@Name='Vertical']/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[3][@Name='Page down']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Pane[1][@Name='Vertical']/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[4][@Name='Line down']"
  Debug.Assert This.MatchingElementspPaths(5) = "/Pane[2][@Name='Horizontal']/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[1][@Name='Column left']"
  Debug.Assert This.MatchingElementspPaths(6) = "/Pane[2][@Name='Horizontal']/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page left']"
  Debug.Assert This.MatchingElementspPaths(7) = "/Pane[2][@Name='Horizontal']/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[3][@Name='Page right']"
  Debug.Assert This.MatchingElementspPaths(8) = "/Pane[2][@Name='Horizontal']/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[4][@Name='Column right']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test077()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
     
  Evaluation "Evaluation_Test077", "//ScrollBar[count(./Button)=4]", 2, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Vertical']/Pane[1][@Name='']/ScrollBar[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[2][@Name='Horizontal']/Pane[1][@Name='']/ScrollBar[1][@Name='']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test078()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
       
  Set This.TestLocator = Nothing
  Set This.TestLocator = Factory.GetNewLocator
  This.CurrentpPath = "//ScrollBar[count(/Button)=6]"
  With This.TestLocator
    .Initialise "Evaluation_Test078", RootTestLocator, None, pPath, This.CurrentpPath
    .FindAll TimeoutInSeconds:=0, AcceptNoElements:=True
    Debug.Assert .ReturnedValue = True 'TODO: Shoudl be false!?
    Debug.Assert .ErrorMessage = "A context node is provided but isn't used!"
    Debug.Assert .NumberOfMatchingElements = 2 'TODO: This should be reset to zero before processing the Predicates
  End With
  
  Set RootTestLocator = Nothing
  Set This.TestLocator = Nothing

End Sub

Public Sub Evaluation_Test079()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[1]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
       
  Set This.TestLocator = Nothing
  Set This.TestLocator = Factory.GetNewLocator
  This.CurrentpPath = "//StatusBar[notafunction(/Button)=6]"
  With This.TestLocator
    .Initialise "Evaluation_Test079", RootTestLocator, None, pPath, This.CurrentpPath
    .FindAll TimeoutInSeconds:=0, AcceptNoElements:=True
    Debug.Assert .ReturnedValue = False
    Debug.Assert .ErrorMessage = "Invalid Predicate [notafunction(/Button)=6] => notafunction(/Button)=6"
    Debug.Assert .NumberOfMatchingElements = 0
  End With
  
  Set RootTestLocator = Nothing
  Set This.TestLocator = Nothing

End Sub

Public Sub Evaluation_Test080()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
       
  Evaluation "Evaluation_Test080a", "//*[And(count(./Button)>=1,count(./SplitButton)>1,count(./ComboBox)>1)]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']"
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test081a()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
       
  Evaluation "Evaluation_Test081a", "//Tab[./Button]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test081b()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
       
  Evaluation "Evaluation_Test081b", "//Pane[not(./Button)]", 4, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(4) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test081c()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
       
  Evaluation "Evaluation_Test081c", "/*[not(./Button)]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test081d()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
       
  Evaluation "Evaluation_Test081d", "//Group[not(./Button)]", 4, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']"
  Debug.Assert This.MatchingElementspPaths(4) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test081e()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
       
  Evaluation "Evaluation_Test081e", "//Group[count(./Button)=0]", 4, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']"
  Debug.Assert This.MatchingElementspPaths(4) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test082()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
       
  Evaluation "Evaluation_Test082", "/ToolBar/Pane/Pane", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test083()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
       
  Evaluation "Evaluation_Test083", "/ToolBar[count(./Pane)=1]/Pane/Pane[1]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test084()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//Pane[@Name=""Lower Ribbon""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test084", "//MenuItem[./preceding-sibling::Button]", 12, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(5) = "/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(6) = "/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(7) = "/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']"
  Debug.Assert This.MatchingElementspPaths(8) = "/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(9) = "/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(10) = "/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(11) = "/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(12) = "/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/MenuItem[1][@Name='More Options']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test085()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//Pane[@Name=""Lower Ribbon""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test085", "//MenuItem[./following-sibling::MenuItem]", 5, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[1][@Name='Conditional Formatting']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[2][@Name='Format as Table']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[1][@Name='Fill']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[2][@Name='Clear']"
  Debug.Assert This.MatchingElementspPaths(5) = "/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[3][@Name='Sort & Filter']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test086()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]//Pane[@Name=""Lower Ribbon""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test086", "//MenuItem[Or(./preceding-sibling::Button, ./following-sibling::MenuItem)]", 17, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(4) = "/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(5) = "/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(6) = "/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(7) = "/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']"
  Debug.Assert This.MatchingElementspPaths(8) = "/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(9) = "/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(10) = "/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[1][@Name='Conditional Formatting']"
  Debug.Assert This.MatchingElementspPaths(11) = "/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[2][@Name='Format as Table']"
  Debug.Assert This.MatchingElementspPaths(12) = "/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(13) = "/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(14) = "/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(15) = "/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[1][@Name='Fill']"
  Debug.Assert This.MatchingElementspPaths(16) = "/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[2][@Name='Clear']"
  Debug.Assert This.MatchingElementspPaths(17) = "/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[3][@Name='Sort & Filter']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test087()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test087", "//ToolBar[@Name=""Quick Access Toolbar""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test088()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test088", "//Button[@AcceleratorKey=""Ctrl+S""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[2][@Name='Save']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test089()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane/ComboBox[@Name=""Name Box""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test089", "//Button[@AccessKey=""Alt+Down Arrow""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/Button[1][@Name='Open']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test090()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane/ToolBar)[1]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test090", "//Text[@AriaProperties=""""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[1][@Name='Cell Mode Ready']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test091()
    
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test091", "//ComboBox[@AriaRole=""""]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test092()
    
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test092", "//*[@AutomationId=""AutoSaveSwitch""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']/Button[1][@Name='AutoSave']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test093()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test093", "/*[@BoundingRectangle<>""{627, 818, 1303, 201}""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test094()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test094", "//ComboBox[@CenterPoint=""""]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test095()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_Test095", "/*[@Name=""Vertical""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Vertical']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test096()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test096", "//ComboBox[@Culture=0]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test097()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test097", "//ComboBox[@FillColor=""""]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test098()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test098", "//ComboBox[@FillType=0]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test099()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test099", "//ComboBox[@FrameworkId=""Win32""]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test100()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test100", "//*[@FullDescription=""Just start typing here to bring features to your fingertips and get help.""]", 2, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test101()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test101", "//Button[@FullDescription=""Like the look of a particular selection? You can apply that look to other content in the document.\n\nTo get started: \n1. Select content with the formatting you like\n2. Click Format Painter\n3. Select something else to automatically apply the formatting\n\nFYI: To apply the formatting in multiple places, double-click Format Painter.""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/Button[2][@Name='Format Painter']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test102()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test102", "//Pane[@HasKeyboardFocus=False]", 5, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(4) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']"
  Debug.Assert This.MatchingElementspPaths(5) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test103a()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test103a", "//ToolBar[@HeadingLevel=80050]", 2, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/ToolBar[1][@Name='Quick Access Toolbar']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test103b()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test103b", "/*[@HeadingLevel<>80049]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test104a()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""EXCEL2""][2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test104a", "/ToolBar[@HelpText=""Ribbon toolbar""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test104b()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""EXCEL2""][1]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test104b ", "//StatusBar[@HelpText<>""Ribbon toolbar""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test105()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""EXCEL2""][2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test105", "//Tab[@IsContentElement=True]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test106()
  Evaluation "Evaluation_Test106", "/TitleBar//MenuItem[@IsControlElement=True]", 1, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']/MenuItem[1][@Name='System']"
End Sub

Public Sub Evaluation_Test107()
  Evaluation "Evaluation_Test107", "/TitleBar//MenuItem[@IsDataValidForForm=False]", 1, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']/MenuItem[1][@Name='System']"
End Sub

Public Sub Evaluation_Test108()
  Evaluation "Evaluation_Test108", "/TitleBar//MenuItem[@IsDialog=False]", 1, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/TitleBar[1][@Name='']/MenuBar[1][@Name='System']/MenuItem[1][@Name='System']"
End Sub

Public Sub Evaluation_Test109()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test109", "//Button[@IsEnabled=False]", 2, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[4][@Name='" & This.Workbook.Name & "']/Tab[1][@Name='" & This.Workbook.Name & "']/Button[1][@Name='Scroll Left']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[4][@Name='" & This.Workbook.Name & "']/Tab[1][@Name='" & This.Workbook.Name & "']/Button[2][@Name='Scroll Right']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test110()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane/ToolBar)[1]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test110", "//Text[@IsKeyboardFocusable=True]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[1][@Name='Cell Mode Ready']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test111()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane/Pane[@Name=""Vertical""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test111", "//*[@IsOffscreen=True]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page up']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test112()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane/Pane[@Name=""Vertical""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test112", "//*[@IsOffscreen=True]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='']/ScrollBar[1][@Name='']/Button[2][@Name='Page up']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test113()
  Evaluation "Evaluation_Test113", "/Edit[@IsPeripheral=False]", 1, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/Edit[1][@Name='Formula Bar']"
End Sub

Public Sub Evaluation_Test114()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane/ToolBar)[1]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test114", "//Text[@IsRequiredForForm=False]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Text[1][@Name='Cell Mode Ready']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test115()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane/Pane[@Name=""" & This.Workbook.Name & """]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test115", "//*[@ItemStatus<>""""]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Sheet Sheet1']"
  Debug.Assert This.MatchingElementspPaths(2) = "/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[157][@Name='C7']"
  Debug.Assert This.MatchingElementspPaths(3) = "/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[158][@Name='D7']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test116()
  Evaluation "Evaluation_Test116", "/Edit[@ItemType=""Edit Formula""]", 1, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/Edit[1][@Name='Formula Bar']"
End Sub

Public Sub Evaluation_Test117()
  Evaluation "Evaluation_Test117", "/Edit[@LandmarkType=0]", 1, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/Edit[1][@Name='Formula Bar']"
End Sub

Public Sub Evaluation_Test118()
  Evaluation "Evaluation_Test11", "/Edit[@Level=0]", 1, This.MasterWindow
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  Debug.Assert This.MatchingElementspPaths(1) = "/Edit[1][@Name='Formula Bar']"
End Sub

Public Sub Evaluation_Test119()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test119", "//ComboBox[@LiveSetting=2]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"
 
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test120()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test120", "//ComboBox[@LocalizedControlType=""ComboBox""]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"
 
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test121()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test121", "//ComboBox[@LocalizedLandmarkType=""""]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"
 
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test122()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane/ToolBar)[1]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  Evaluation "Evaluation_Test122", "//*[@Name=""Macro Recording Not Recording""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Status Bar']/Pane[1][@Name='']/Pane[1][@Name='']/StatusBar[1][@Name='Status Bar']/Button[1][@Name='Macro Recording Not Recording']"
 
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test123()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test123", "//ComboBox[@NativeWindowHandle=""#""]", 3, RootTestLocator, UnitTestingMode:=True
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"
 
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test124()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test124", "//ComboBox[@OptimizeForVisualContent=False]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test125()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test125", "//ComboBox[@Orientation=1]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test126()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test126", "//*[@PositionInSet=13]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Tab[1][@Name='Ribbon Tabs']/TabItem[11][@Name='Help']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test127()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test127", "//ComboBox[@ProcessId=""#""]", 3, RootTestLocator, UnitTestingMode:=True
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test128()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test128", "//ComboBox[@ProviderDescription=""#""]", 3, RootTestLocator, UnitTestingMode:=True
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test129()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test129", "//ComboBox[@Rotation=""""]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test130()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test130", "//*[@SizeOfSet=3]", 6, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[1][@Name='Conditional Formatting']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[2][@Name='Format as Table']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[3][@Name='Cell Styles']"
  Debug.Assert This.MatchingElementspPaths(4) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']"
  Debug.Assert This.MatchingElementspPaths(5) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']"
  Debug.Assert This.MatchingElementspPaths(6) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/MenuItem[1][@Name='Format']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test131()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test131", "//ComboBox[@VisualEffects=0]", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[1][@Name='Font']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/ComboBox[2][@Name='Font Size']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='']/Pane[1][@Name='']/Pane[1][@Name='Ribbon']/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/ComboBox[1][@Name='Number Format']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test132a()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "((/Pane[@ClassName=""EXCEL2""])[2]/ToolBar[1]/Pane[@Name=""Ribbon""])[1]//Pane[@Name=""Ribbon""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test132a", "//MenuItem[position()<=2][1]", 20, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(2) = "/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[2][@Name='Redo']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(3) = "/ToolBar[1][@Name='Quick Access Toolbar']/MenuItem[1][@Name='Customize Quick Access Toolbar']"
  Debug.Assert This.MatchingElementspPaths(4) = "/MenuItem[1][@Name='Type to search and use the up and down arrow keys to navigate']"
  Debug.Assert This.MatchingElementspPaths(5) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[1][@Name='Paste']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(6) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[1][@Name='Clipboard']/SplitButton[2][@Name='Copy']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(7) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(8) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(9) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(10) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[4][@Name='Font Color']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(11) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']"
  Debug.Assert This.MatchingElementspPaths(12) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/SplitButton[1][@Name='Merge & Center']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(13) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[4][@Name='Number']/SplitButton[1][@Name='Accounting Number Format']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(14) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[5][@Name='Styles']/MenuItem[1][@Name='Conditional Formatting']"
  Debug.Assert This.MatchingElementspPaths(15) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[1][@Name='Insert']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(16) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/SplitButton[2][@Name='Delete']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(17) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[6][@Name='Cells']/MenuItem[1][@Name='Format']"
  Debug.Assert This.MatchingElementspPaths(18) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/SplitButton[1][@Name='AutoSum']/MenuItem[1][@Name='More Options']"
  Debug.Assert This.MatchingElementspPaths(19) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[7][@Name='Editing']/MenuItem[1][@Name='Fill']"
  Debug.Assert This.MatchingElementspPaths(20) = "/MenuItem[3][@Name='Ribbon Display Options']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test132b()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "((/Pane[@ClassName=""EXCEL2""])[2]/ToolBar[1]/Pane[@Name=""Ribbon""])[1]//Pane[@Name=""Ribbon""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test132b", "(//MenuItem[position()<=2])[1]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']/MenuItem[1][@Name='More Options']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test133a()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "((/Pane[@ClassName=""EXCEL2""])[2]/ToolBar[1]/Pane[@Name=""Ribbon""])[1]//Pane[@Name=""Ribbon""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test133a", "//MenuItem[And(./preceding-sibling::Button,./following-sibling::SplitButton)]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_Test133b()
  
  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "(/Pane[@ClassName=""EXCEL2""])[2]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Dim ContextNodeLocator As pLocator
  Set ContextNodeLocator = Factory.GetNewLocator
  pPathString = "/ToolBar[1]/Pane[@Name=""Ribbon""]//Pane[@Name=""Ribbon""]"
  With ContextNodeLocator
    .Initialise "ContextNodeLocator", RootTestLocator, None, pPath, pPathString
    .Find 10
  End With
   
  Evaluation "Evaluation_Test133b", ".//*[And(./preceding-sibling::Button,./following-sibling::SplitButton)]", 17, RootTestLocator, Array(ContextNodeLocator), Array(pPathString)
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
  
  Debug.Assert This.MatchingElementspPaths(1) = pPathString & "/ToolBar[1][@Name='Quick Access Toolbar']/Button[2][@Name='Save']"
  Debug.Assert This.MatchingElementspPaths(2) = pPathString & "/ToolBar[1][@Name='Quick Access Toolbar']/SplitButton[1][@Name='Undo']"
  Debug.Assert This.MatchingElementspPaths(3) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[2][@Name='Decrease Font Size']"
  Debug.Assert This.MatchingElementspPaths(4) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[3][@Name='Bold']"
  Debug.Assert This.MatchingElementspPaths(5) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/Button[4][@Name='Italic']"
  Debug.Assert This.MatchingElementspPaths(6) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[1][@Name='Underline']"
  Debug.Assert This.MatchingElementspPaths(7) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[2][@Name='Borders']"
  Debug.Assert This.MatchingElementspPaths(8) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[2][@Name='Font']/SplitButton[3][@Name='Fill Color']"
  Debug.Assert This.MatchingElementspPaths(9) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[2][@Name='Middle Align']"
  Debug.Assert This.MatchingElementspPaths(10) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[3][@Name='Bottom Align']"
  Debug.Assert This.MatchingElementspPaths(11) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']"
  Debug.Assert This.MatchingElementspPaths(12) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[4][@Name='Align Left']"
  Debug.Assert This.MatchingElementspPaths(13) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[5][@Name='Center']"
  Debug.Assert This.MatchingElementspPaths(14) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[6][@Name='Align Right']"
  Debug.Assert This.MatchingElementspPaths(15) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[7][@Name='Decrease Indent']"
  Debug.Assert This.MatchingElementspPaths(16) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[8][@Name='Increase Indent']"
  Debug.Assert This.MatchingElementspPaths(17) = pPathString & "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/Button[9][@Name='Wrap Text']"
  
  Set RootTestLocator = Nothing
  Set ContextNodeLocator = Nothing

End Sub

Public Sub Evaluation_Test134()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "((/Pane[@ClassName=""EXCEL2""])[2]/ToolBar[1]/Pane[@Name=""Ribbon""])[1]//Pane[@Name=""Ribbon""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_Test134", "//MenuItem[./preceding-sibling::Button][./following-sibling::SplitButton]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Lower Ribbon']/Group[1][@Name='Home']/Group[3][@Name='Alignment']/MenuItem[1][@Name='Orientation']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel001()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_TestExcel001", "/Pane[@Name=""Sheet Sheet1""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Sheet Sheet1']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel002()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_TestExcel002", "//DataItem[@Name=""A1""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Sheet Sheet1']/DataGrid[1][@Name='Grid']/DataItem[23][@Name='A1']"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel003()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_TestExcel003", "//element(*, string)", 6, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[23][@Name='A1']/@value"
  Debug.Assert This.MatchingElementspPaths(2) = "/DataItem[45][@Name='A2']/@value"
  Debug.Assert This.MatchingElementspPaths(3) = "/DataItem[46][@Name='B2']/@value"
  Debug.Assert This.MatchingElementspPaths(4) = "/DataItem[48][@Name='D2']/@value"
  Debug.Assert This.MatchingElementspPaths(5) = "/DataItem[49][@Name='E2']/@value"
  Debug.Assert This.MatchingElementspPaths(6) = "/DataItem[134][@Name='B6']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel004()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_TestExcel004", "//element(*, integer)", 5, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[67][@Name='A3']/@value"
  Debug.Assert This.MatchingElementspPaths(2) = "/DataItem[68][@Name='B3']/@value"
  Debug.Assert This.MatchingElementspPaths(3) = "/DataItem[69][@Name='C3']/@value"
  Debug.Assert This.MatchingElementspPaths(4) = "/DataItem[70][@Name='D3']/@value"
  Debug.Assert This.MatchingElementspPaths(5) = "/DataItem[71][@Name='E3']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel005()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_TestExcel005", "//element(*, decimal)", 10, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Evaluation_TestExcel005_pPaths
    
  Set RootTestLocator = Nothing

End Sub

Private Sub Evaluation_TestExcel005_pPaths()
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[67][@Name='A3']/@value"
  Debug.Assert This.MatchingElementspPaths(2) = "/DataItem[68][@Name='B3']/@value"
  Debug.Assert This.MatchingElementspPaths(3) = "/DataItem[69][@Name='C3']/@value"
  Debug.Assert This.MatchingElementspPaths(4) = "/DataItem[70][@Name='D3']/@value"
  Debug.Assert This.MatchingElementspPaths(5) = "/DataItem[71][@Name='E3']/@value"
  Debug.Assert This.MatchingElementspPaths(6) = "/DataItem[89][@Name='A4']/@value"
  Debug.Assert This.MatchingElementspPaths(7) = "/DataItem[90][@Name='B4']/@value"
  Debug.Assert This.MatchingElementspPaths(8) = "/DataItem[91][@Name='C4']/@value"
  Debug.Assert This.MatchingElementspPaths(9) = "/DataItem[92][@Name='D4']/@value"
  Debug.Assert This.MatchingElementspPaths(10) = "/DataItem[93][@Name='E4']/@value"
End Sub

Public Sub Evaluation_TestExcel006()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_TestExcel006", "//element(*, hyperlink)", 2, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[111][@Name='A5']/@value"
  Debug.Assert This.MatchingElementspPaths(2) = "/DataItem[112][@Name='B5']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel007()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_TestExcel007", "//element(*, date)", 3, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[133][@Name='A6']/@value"
  Debug.Assert This.MatchingElementspPaths(2) = "/DataItem[135][@Name='C6']/@value"
  Debug.Assert This.MatchingElementspPaths(3) = "/DataItem[136][@Name='D6']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel008()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_TestExcel008", "//element(*, boolean)", 6, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths

  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[155][@Name='A7']/@value"
  Debug.Assert This.MatchingElementspPaths(2) = "/DataItem[156][@Name='B7']/@value"
  Debug.Assert This.MatchingElementspPaths(3) = "/DataItem[157][@Name='C7']/@value"
  Debug.Assert This.MatchingElementspPaths(4) = "/DataItem[158][@Name='D7']/@value"
  Debug.Assert This.MatchingElementspPaths(5) = "/DataItem[159][@Name='E7']/@value"
  Debug.Assert This.MatchingElementspPaths(6) = "/DataItem[160][@Name='F7']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel009()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
  
  Evaluation "Evaluation_TestExcel009", "count(//element(*, decimal))", 10, RootTestLocator, ReturnedValue:=10
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Evaluation_TestExcel005_pPaths
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel010()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel010", "sum(//element(*, decimal))", 10, RootTestLocator, ReturnedValue:=1423.77
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Evaluation_TestExcel005_pPaths
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel011()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel011", "average(//element(*, decimal))", 10, RootTestLocator, ReturnedValue:=142.377
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Evaluation_TestExcel005_pPaths
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel012()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel012", "stdeva(//element(*, decimal))", 10, RootTestLocator, ReturnedValue:=204.529890075863
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Evaluation_TestExcel005_pPaths
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel013()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel013", "round(stdeva(//element(*, decimal)),2)", 10, RootTestLocator, ReturnedValue:=204.53
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Evaluation_TestExcel005_pPaths
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel014()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel014", "round(stdeva(//element(*, decimal)),6)", 10, RootTestLocator, ReturnedValue:=204.52989
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Evaluation_TestExcel005_pPaths
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel015()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel015", "min(//element(*, decimal))", 10, RootTestLocator, ReturnedValue:=-1
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Evaluation_TestExcel005_pPaths
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel016()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel016", "max(//element(*, decimal))", 10, RootTestLocator, ReturnedValue:=521.01
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Evaluation_TestExcel005_pPaths
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel017()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel017", "roundup(stdeva(//element(*, decimal)),0)", 10, RootTestLocator, ReturnedValue:=205
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Evaluation_TestExcel005_pPaths
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel018()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel018", "rounddown(stdeva(//element(*, decimal)),0)", 10, RootTestLocator, ReturnedValue:=204
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Evaluation_TestExcel005_pPaths
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel019()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel019", "product(//element(*, decimal))", 10, RootTestLocator, ReturnedValue:=-4252982755186.95
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Evaluation_TestExcel005_pPaths
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel020()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel020", "//Pane[xp:starts-with(@Name,""Sheet Sheet1"")]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/Pane[1][@Name='Sheet Sheet1']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel100()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel100", "//DataItem[text()=""Cat""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[45][@Name='A2']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel101()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel101", "//DataItem[lower(text())=""cat""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[45][@Name='A2']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel102()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel102", "//DataItem[upper(text())=""CAT""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[45][@Name='A2']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel103()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel103", "//DataItem[@Name=""A2""][upper(text())=""CAT""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[45][@Name='A2']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel104()
Exit Sub
'Not working

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    '//DataItem[@Name=""A2""]
  Evaluation "Evaluation_TestExcel104", "//DataItem[translate(text(),""en"",""fr"")=""Chat""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[45][@Name='A2']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel105()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With
    
  Evaluation "Evaluation_TestExcel105", "//DataItem[@Name=""A2""][udf:reversetext(text())=""taC""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[45][@Name='A2']"

  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel106()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel106", "sum(//element(*, decimal)[ceiling_math(value())=405])", 1, RootTestLocator, ReturnedValue:=404.36
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[92][@Name='D4']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel107()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel107", "sum(//element(*, integer)[ceiling_math(value(),6,TRUE)=-6])", 1, RootTestLocator, ReturnedValue:=-1
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[71][@Name='E3']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel108()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel108", "sum(//element(*, decimal)[floor_math(value(),5,false)=-5])", 1, RootTestLocator, ReturnedValue:=-1
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[71][@Name='E3']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel109()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel109", "//element(*, string)[concat(text(),text(),text())=""RabbitRabbitRabbit""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[48][@Name='D2']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel110()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel110", "//element(*, hyperlink)[xp:contains(text(),""google"")]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[111][@Name='A5']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel111()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel111", "//element(*, date)[xp:format-number(text(),""ddmmyyyy"")=""06042003""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[136][@Name='D6']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel112()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel112", "//element(*, string)[xp:starts-with(text(),""1st April"")]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[134][@Name='B6']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel113()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel113", "//element(*, string)[xp:normalize-space(text())=""OneTwoThreeFourFive Six""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[49][@Name='E2']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel114()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel114", "//element(*, string)[xp:string-length(text())=6]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[48][@Name='D2']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel115()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel115", "//element(*, string)[xp:substring(text(),4)=""bit""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[48][@Name='D2']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel116()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel116", "//element(*, string)[xp:substring(text(),2,5)=""abbi""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[48][@Name='D2']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel117()
'TODO: End-user defined functions
End Sub

Public Sub Evaluation_TestExcel118()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel118", "//element(*, string)[xp:string-after(text(),""abb"")=""it""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[48][@Name='D2']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel119()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel119", "//element(*, string)[xp:string-before(text(),""bit"")=""Rab""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[48][@Name='D2']/@value"
  
  Set RootTestLocator = Nothing

End Sub

'//Pane[@Name=""Sheet Sheet1""]
Public Sub Evaluation_TestExcel120()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel120", "//element(*, string)[xp:translate(text(),""bit"",""BIT"")=""RabBIT""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[48][@Name='D2']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel121()

  Dim pPathString As String
  Dim RootTestLocator As pLocator
  Set RootTestLocator = Factory.GetNewLocator
  pPathString = "/Pane[@ClassName=""XLDESK""]/Pane[@Name=""" & This.Workbook.Name & """]//Pane[@Name=""Sheet Sheet1""]/DataGrid[1][@Name=""Grid""]"
  With RootTestLocator
    .Initialise "RootTestLocator", This.MasterWindow, None, pPath, pPathString
    .Find 10
  End With

  Evaluation "Evaluation_TestExcel121", "//element(*, string)[udf:text-between(text(),""Ra"",""it"")=""bb""]", 1, RootTestLocator
  This.MatchingElementspPaths = This.TestLocator.MatchingElementspPaths
    
  Debug.Assert This.MatchingElementspPaths(1) = "/DataItem[48][@Name='D2']/@value"
  
  Set RootTestLocator = Nothing

End Sub

Public Sub Evaluation_TestExcel122()
'TODO: End-user defined functions
End Sub

Private Sub Compare(SourcepPath As String, TargetpPath As String)
  Dim Length As Integer
  Length = Len(SourcepPath)
  Dim i As Integer
  For i = 1 To Length
    If Not Mid(SourcepPath, i, 1) = Mid(TargetpPath, i, 1) Then
      Debug.Print i; "'" & Mid(SourcepPath, i, 1) & "': "; Mid(SourcepPath, i, 1) = Mid(TargetpPath, i, 1); AscW(Mid(SourcepPath, i, 1))
    End If
  Next i
End Sub













